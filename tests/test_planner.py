import pytest
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unittest.mock import MagicMock, patch
from agents.planner import PlannerAgent
from models.schemas import PresentationPlan, SlideLayout, LayoutValidationError


@pytest.fixture
def mock_planner():
    """返回一个 mock 掉 API 和文件读取的 PlannerAgent"""
    with patch("agents.planner.OpenAI") as mock_openai:
        with patch("agents.planner.Path") as mock_path:
            mock_path.return_value.read_text.return_value = "mock prompt"
            planner = PlannerAgent()
            planner.client = MagicMock()
            planner._system_template = "s {slide_width} {slide_height}"
            planner._user_template = "u {topic} {slide_width} {slide_height} {language} {min_slides} {max_slides}"
            yield planner


# ─── 代码提取测试 ───


def test_extract_code_tag(mock_planner):
    """从 <code>...</code> 提取代码"""
    raw = "Here is the code:\n<code>\ndef generate_slides(output_dir):\n    return 0\n</code>\nDone."
    code = mock_planner._extract_code(raw)
    assert "def generate_slides" in code


def test_extract_python_block(mock_planner):
    """从 ```python...``` 提取代码"""
    raw = "```python\ndef generate_slides(output_dir):\n    return 0\n```"
    code = mock_planner._extract_code(raw)
    assert "def generate_slides" in code


def test_extract_generic_block(mock_planner):
    """从 ```...``` 提取代码"""
    raw = "```\ndef generate_slides(output_dir):\n    return 0\n```"
    code = mock_planner._extract_code(raw)
    assert "def generate_slides" in code


def test_no_code_raises(mock_planner):
    """无代码块应报错"""
    with pytest.raises(ValueError, match="未找到"):
        mock_planner._extract_code("这里没有代码")


# ─── 代码执行测试 ───


def test_execute_valid_code(mock_planner, tmp_path):
    """合法代码执行后生成 XML 文件"""
    xml_dir = str(tmp_path)
    code = '''
def generate_slides(output_dir):
    import os, xml.etree.ElementTree as ET
    os.makedirs(output_dir, exist_ok=True)
    s = ET.Element("slide", index="0", layout="cover", topic="封面", background_color="#1E2761")
    t = ET.SubElement(s, "element", type="title", x="1.5", y="2.0", width="10.333", height="1.8",
                      font_size="44", bold="true", color="#FFFFFF", align="center")
    t.text = "测试"
    ET.ElementTree(s).write(os.path.join(output_dir, "slide_0.xml"), encoding="unicode", xml_declaration=True)
    s1 = ET.Element("slide", index="1", layout="closing", topic="结尾", background_color="#1E2761")
    t1 = ET.SubElement(s1, "element", type="title", x="1.5", y="2.5", width="10.333", height="1.5",
                       font_size="44", bold="true", color="#FFFFFF", align="center")
    t1.text = "结束"
    ET.ElementTree(s1).write(os.path.join(output_dir, "slide_1.xml"), encoding="unicode", xml_declaration=True)
    p = ET.Element("presentation", title="测试", topic="测试", theme_color="#1E2761",
                   accent_color="#CADCFC", font_family="Microsoft YaHei", slide_count="2")
    ET.ElementTree(p).write(os.path.join(output_dir, "presentation.xml"), encoding="unicode", xml_declaration=True)
    return 2
'''
    mock_planner._execute_code(code, xml_dir)
    assert os.path.isfile(os.path.join(xml_dir, "slide_0.xml"))
    assert os.path.isfile(os.path.join(xml_dir, "presentation.xml"))


def test_execute_bad_code_raises(mock_planner, tmp_path):
    """语法错误的代码应抛出 RuntimeError"""
    code = "def generate_slides(output_dir):\n    raise ValueError('故意失败')"
    with pytest.raises(RuntimeError, match="代码执行失败"):
        mock_planner._execute_code(code, str(tmp_path))


# ─── 集成测试（需要真实 API） ───


def test_plan_with_real_api():
    """集成测试：调用真实 GLM API"""
    if not os.getenv("GLM_API_KEY"):
        pytest.skip("未设置 GLM_API_KEY，跳过真实 API 测试")

    planner = PlannerAgent()
    plan = planner.plan("量子计算入门")

    assert isinstance(plan, PresentationPlan)
    assert len(plan.slides) >= 2
    assert plan.slides[0].layout == SlideLayout.COVER

    print(f"\n生成 {len(plan.slides)} 页幻灯片")
    for s in plan.slides:
        print(f"  第 {s.slide_index} 页：{s.layout.value} - {s.topic}")
