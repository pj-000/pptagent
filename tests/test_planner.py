import pytest
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unittest.mock import MagicMock, patch
from agents.planner import PlannerAgent
from models.schemas import PresentationPlan, SlideLayout


@pytest.fixture
def mock_planner():
    """返回一个 mock 掉 API 调用的 PlannerAgent，不消耗真实 token"""
    with patch("agents.planner.OpenAI") as mock_openai:
        planner = PlannerAgent()
        planner.client = MagicMock()
        yield planner


MOCK_VALID_RESPONSE = '''{
  "title": "人工智能发展趋势",
  "topic": "人工智能发展趋势",
  "theme_color": "#1F3864",
  "accent_color": "#2E75B6",
  "slides": [
    {
      "slide_index": 0,
      "layout": "cover",
      "topic": "封面",
      "content": {"title": "人工智能发展趋势", "subtitle": "2024年度报告"}
    },
    {
      "slide_index": 1,
      "layout": "toc",
      "topic": "目录",
      "content": {"title": "目录", "body": ["AI 技术现状", "应用场景", "未来展望"]}
    },
    {
      "slide_index": 2,
      "layout": "content",
      "topic": "AI 技术现状",
      "content": {"title": "AI 技术现状", "body": "• 大语言模型快速发展\\n• 多模态能力突破\\n• 推理效率大幅提升"}
    },
    {
      "slide_index": 3,
      "layout": "two_column",
      "topic": "应用场景对比",
      "content": {
        "title": "行业应用对比",
        "left": "成熟应用\\n• 文本生成\\n• 代码辅助\\n• 图像生成",
        "right": "新兴方向\\n• 具身智能\\n• 科学发现\\n• 自主代理"
      }
    },
    {
      "slide_index": 4,
      "layout": "closing",
      "topic": "结尾",
      "content": {"title": "感谢聆听", "subtitle": "欢迎交流讨论"}
    }
  ]
}'''


def test_parse_valid_json(mock_planner):
    """测试合法 JSON 能正确解析"""
    result = mock_planner._parse_json(MOCK_VALID_RESPONSE)
    assert result["title"] == "人工智能发展趋势"
    assert len(result["slides"]) == 5


def test_parse_json_with_markdown_wrapper(mock_planner):
    """测试带 ```json 包裹的响应也能正确解析"""
    wrapped = f"```json\n{MOCK_VALID_RESPONSE}\n```"
    result = mock_planner._parse_json(wrapped)
    assert result["title"] == "人工智能发展趋势"


def test_parse_invalid_json_raises(mock_planner):
    """测试非法 JSON 会抛出有意义的错误"""
    with pytest.raises(ValueError, match="不是合法 JSON"):
        mock_planner._parse_json("这不是 JSON")


def test_build_plan_returns_presentation_plan(mock_planner):
    """测试 _build_plan 返回正确类型"""
    data = mock_planner._parse_json(MOCK_VALID_RESPONSE)
    plan = mock_planner._build_plan(data)
    assert isinstance(plan, PresentationPlan)
    assert plan.title == "人工智能发展趋势"
    assert len(plan.slides) == 5


def test_build_plan_correct_layouts(mock_planner):
    """测试每页的布局类型被正确映射"""
    data = mock_planner._parse_json(MOCK_VALID_RESPONSE)
    plan = mock_planner._build_plan(data)
    assert plan.slides[0].layout == SlideLayout.COVER
    assert plan.slides[1].layout == SlideLayout.TOC
    assert plan.slides[4].layout == SlideLayout.CLOSING


def test_build_plan_elements_have_coordinates(mock_planner):
    """测试模板注入后每个元素都有坐标"""
    data = mock_planner._parse_json(MOCK_VALID_RESPONSE)
    plan = mock_planner._build_plan(data)
    for slide in plan.slides:
        assert len(slide.elements) > 0, f"第 {slide.slide_index} 页没有元素"
        for elem in slide.elements:
            assert elem.x >= 0
            assert elem.y >= 0
            assert elem.width > 0
            assert elem.height > 0


def test_build_plan_unknown_layout_fallback(mock_planner):
    """测试未知布局类型会 fallback 到 content"""
    data = mock_planner._parse_json(MOCK_VALID_RESPONSE)
    data["slides"][2]["layout"] = "unknown_layout_type"
    plan = mock_planner._build_plan(data)
    assert plan.slides[2].layout == SlideLayout.CONTENT


def test_plan_with_real_api():
    """
    集成测试：调用真实 GLM API（需要设置 GLM_API_KEY）。
    运行命令：pytest tests/test_planner.py::test_plan_with_real_api -v -s
    """
    if not os.getenv("GLM_API_KEY"):
        pytest.skip("未设置 GLM_API_KEY，跳过真实 API 测试")

    planner = PlannerAgent()
    plan = planner.plan("量子计算入门")

    assert isinstance(plan, PresentationPlan)
    assert len(plan.slides) >= 4
    assert plan.slides[0].layout == SlideLayout.COVER
    assert plan.slides[-1].layout == SlideLayout.CLOSING

    print(f"\n生成 {len(plan.slides)} 页幻灯片")
    for s in plan.slides:
        print(f"  第 {s.slide_index + 1} 页：{s.layout.value} - {s.topic}")
