import pytest
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unittest.mock import MagicMock, patch
from agents.planner import PlannerAgent, normalize_audience, suggest_audience_label
from models.schemas import OutlinePlan


@pytest.fixture
def mock_planner():
    """返回一个 mock 掉 API 和文件读取的 PlannerAgent"""
    with patch("agents.planner.OpenAI"):
        with patch("agents.planner.assert_skill_present"):
            with patch("agents.planner.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock prompt"
                planner = PlannerAgent()
                planner.client = MagicMock()
                planner._skill_md = "## Design Ideas\n- test"
                planner._pptxgenjs_md = "# PptxGenJS Tutorial\n- test"
                planner._user_template = (
                    "topic={topic} lang={language} style={style} style_ref={style_reference} "
                    "audience={audience} audience_ref={audience_reference} "
                    "profile={audience_profile} outline={outline_context} "
                    "research={research_context} "
                    "min={min_slides} max={max_slides}"
                )
                yield planner


def test_extract_code_tag(mock_planner):
    raw = "Here is the code:\n<code>\nconst pptxgen = require('pptxgenjs');\n</code>\nDone."
    code = mock_planner._extract_code(raw)
    assert "pptxgenjs" in code


def test_extract_javascript_block(mock_planner):
    raw = "```javascript\nconst pptxgen = require('pptxgenjs');\n```"
    code = mock_planner._extract_code(raw)
    assert "pptxgenjs" in code


def test_extract_generic_block(mock_planner):
    raw = "```\nconst pptxgen = require('pptxgenjs');\n```"
    code = mock_planner._extract_code(raw)
    assert "pptxgenjs" in code


def test_no_code_raises(mock_planner):
    with pytest.raises(ValueError, match="未找到"):
        mock_planner._extract_code("这里没有代码")


def test_inject_output_path_existing_writefile(mock_planner):
    code = 'pres.writeFile({ fileName: "old.pptx" });'
    new_code = mock_planner._inject_output_path(code, "/tmp/new.pptx")
    assert 'fileName: "/tmp/new.pptx"' in new_code
    assert "old.pptx" not in new_code


def test_inject_output_path_append_writefile(mock_planner):
    code = 'const pptxgen = require("pptxgenjs");\nlet pres = new pptxgen();'
    new_code = mock_planner._inject_output_path(code, "/tmp/new.pptx")
    assert 'pres.writeFile({ fileName: "/tmp/new.pptx" });' in new_code


def test_build_user_prompt_includes_style(mock_planner):
    prompt = mock_planner._build_user_prompt(
        "测试主题",
        6,
        8,
        language="English",
        style="minimal",
        audience="投资人",
        outline_context="第0页 cover",
        research_context="主题摘要：行业增长快\n- 要点A",
    )
    assert "style=minimal" in prompt
    assert "style_ref=minimal" in prompt
    assert "topic=测试主题" in prompt
    assert "lang=English" in prompt
    assert "audience=投资人" in prompt
    assert "audience_ref=investor" in prompt
    assert "市场空间" in prompt
    assert "outline=第0页 cover" in prompt
    assert "research=主题摘要：行业增长快\n- 要点A" in prompt


def test_extract_json_handles_code_fence(mock_planner):
    data = mock_planner._extract_json('```json\n{"title":"测试"}\n```')
    assert data["title"] == "测试"


def test_sanitize_generated_code_fixes_shape_aliases(mock_planner):
    code = (
        'slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 1, h: 1 });\n'
        'slide.addShape("RECTANGLE", { x: 1, y: 1, w: 1, h: 1 });'
    )
    sanitized = mock_planner._sanitize_generated_code(code)
    assert 'addShape("rect"' in sanitized


def test_parse_outline_plan_validates_structure(mock_planner):
    outline = mock_planner._parse_outline_plan(
        {
            "title": "人工智能",
            "topic": "人工智能",
            "slides": [
                {"slide_index": 0, "layout": "cover", "topic": "人工智能概览", "objective": "开场"},
                {"slide_index": 1, "layout": "toc", "topic": "目录", "objective": "建立结构"},
                {"slide_index": 2, "layout": "content", "topic": "核心技术", "objective": "解释技术"},
                {"slide_index": 3, "layout": "closing", "topic": "总结", "objective": "收尾"},
            ],
        },
        "人工智能",
    )
    assert isinstance(outline, OutlinePlan)
    assert outline.slides[2].topic == "核心技术"


def test_outline_to_research_slides(mock_planner):
    outline = OutlinePlan.model_validate(
        {
            "title": "人工智能",
            "topic": "人工智能",
            "slides": [
                {"slide_index": 0, "layout": "cover", "topic": "封面"},
                {"slide_index": 1, "layout": "toc", "topic": "目录"},
                {"slide_index": 2, "layout": "content", "topic": "应用场景"},
                {"slide_index": 3, "layout": "closing", "topic": "结束"},
            ],
        }
    )
    slides = mock_planner.outline_to_research_slides(outline)
    assert len(slides) == 4
    assert slides[2].topic == "应用场景"
    assert slides[2].elements[0].type == "title"


def test_normalize_audience_aliases():
    assert normalize_audience("大学生") == "大学生"
    assert suggest_audience_label("大学生") == "student"
    assert suggest_audience_label("老板") == "boss"
    assert suggest_audience_label("投资人") == "investor"
    assert suggest_audience_label("unknown-audience") is None


def test_plan_with_real_api():
    if not os.getenv("GLM_API_KEY"):
        pytest.skip("未设置 GLM_API_KEY，跳过真实 API 测试")

    planner = PlannerAgent()
    output_path = planner.plan("量子计算入门")

    assert isinstance(output_path, str)
    assert output_path.endswith(".pptx")
    assert os.path.exists(output_path)

    print(f"\n生成文件: {output_path}")
