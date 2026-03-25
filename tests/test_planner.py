import pytest
import os
import sys
import json

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unittest.mock import MagicMock, patch
from agents.planner import PlannerAgent
from models.schemas import PresentationPlan, SlideLayout


@pytest.fixture
def mock_planner():
    """返回一个 mock 掉 API 调用的 PlannerAgent，不消耗真实 token"""
    with patch("agents.planner.OpenAI") as mock_openai:
        with patch("agents.planner.Path") as mock_path:
            mock_path.return_value.read_text.return_value = "mock prompt"
            planner = PlannerAgent()
            planner.client = MagicMock()
            planner._system_template = "system {slide_width} {slide_height}"
            planner._user_template = "user {topic} {slide_width} {slide_height} {language} {min_slides} {max_slides}"
            yield planner


# Phase 2 格式的 mock 响应（含完整坐标）
MOCK_VALID_RESPONSE = json.dumps({
    "title": "人工智能发展趋势",
    "topic": "人工智能发展趋势",
    "theme_color": "#1F3864",
    "accent_color": "#2E75B6",
    "slides": [
        {
            "slide_index": 0, "layout": "cover", "topic": "封面",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "人工智能发展趋势", "x": 1.5, "y": 2.0,
                 "width": 10.333, "height": 1.8, "font_size": 48, "bold": True,
                 "color": "#1F3864", "align": "center"},
                {"type": "subtitle", "content": "2024年度报告", "x": 2.0, "y": 4.2,
                 "width": 9.333, "height": 0.8, "font_size": 24, "bold": False,
                 "color": "#2E75B6", "align": "center"},
            ]
        },
        {
            "slide_index": 1, "layout": "toc", "topic": "目录",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "目录", "x": 0.5, "y": 0.3,
                 "width": 12.333, "height": 0.9, "font_size": 32, "bold": True,
                 "color": "#1F3864", "align": "left"},
                {"type": "body", "content": "◆ AI 技术现状\n◆ 应用场景\n◆ 未来展望",
                 "x": 1.0, "y": 1.5, "width": 11.333, "height": 5.5,
                 "font_size": 20, "bold": False, "color": "#333333", "align": "left"},
            ]
        },
        {
            "slide_index": 2, "layout": "content", "topic": "AI 技术现状",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "AI 技术现状", "x": 0.5, "y": 0.3,
                 "width": 12.333, "height": 0.9, "font_size": 32, "bold": True,
                 "color": "#1F3864", "align": "left"},
                {"type": "body", "content": "• 大语言模型快速发展\n• 多模态能力突破\n• 推理效率大幅提升",
                 "x": 0.5, "y": 1.5, "width": 12.333, "height": 5.5,
                 "font_size": 18, "bold": False, "color": "#333333", "align": "left"},
            ]
        },
        {
            "slide_index": 3, "layout": "two_column", "topic": "应用场景对比",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "行业应用对比", "x": 0.5, "y": 0.3,
                 "width": 12.333, "height": 0.9, "font_size": 32, "bold": True,
                 "color": "#1F3864", "align": "left"},
                {"type": "body", "content": "成熟应用\n• 文本生成\n• 代码辅助",
                 "x": 0.5, "y": 1.5, "width": 5.9, "height": 5.5,
                 "font_size": 17, "bold": False, "color": "#333333", "align": "left"},
                {"type": "body", "content": "新兴方向\n• 具身智能\n• 科学发现",
                 "x": 6.9, "y": 1.5, "width": 5.9, "height": 5.5,
                 "font_size": 17, "bold": False, "color": "#333333", "align": "left"},
            ]
        },
        {
            "slide_index": 4, "layout": "closing", "topic": "结尾",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "感谢聆听", "x": 1.5, "y": 2.2,
                 "width": 10.333, "height": 1.5, "font_size": 44, "bold": True,
                 "color": "#1F3864", "align": "center"},
                {"type": "subtitle", "content": "欢迎交流讨论", "x": 2.0, "y": 4.2,
                 "width": 9.333, "height": 0.8, "font_size": 22, "bold": False,
                 "color": "#2E75B6", "align": "center"},
            ]
        },
    ]
}, ensure_ascii=False)


def _mock_response(content: str):
    mock_resp = MagicMock()
    mock_resp.choices = [MagicMock()]
    mock_resp.choices[0].message.content = content
    mock_resp.usage = MagicMock()
    return mock_resp


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


def test_plan_returns_presentation_plan(mock_planner):
    """测试 plan() 返回正确类型"""
    mock_planner.client.chat.completions.create.return_value = _mock_response(MOCK_VALID_RESPONSE)
    plan = mock_planner.plan("人工智能发展趋势")
    assert isinstance(plan, PresentationPlan)
    assert plan.title == "人工智能发展趋势"
    assert len(plan.slides) == 5


def test_plan_correct_layouts(mock_planner):
    """测试每页的布局类型被正确映射"""
    mock_planner.client.chat.completions.create.return_value = _mock_response(MOCK_VALID_RESPONSE)
    plan = mock_planner.plan("人工智能发展趋势")
    assert plan.slides[0].layout == SlideLayout.COVER
    assert plan.slides[1].layout == SlideLayout.TOC
    assert plan.slides[4].layout == SlideLayout.CLOSING


def test_plan_elements_have_coordinates(mock_planner):
    """测试每个元素都有坐标"""
    mock_planner.client.chat.completions.create.return_value = _mock_response(MOCK_VALID_RESPONSE)
    plan = mock_planner.plan("人工智能发展趋势")
    for slide in plan.slides:
        assert len(slide.elements) > 0, f"第 {slide.slide_index} 页没有元素"
        for elem in slide.elements:
            assert elem.x >= 0
            assert elem.y >= 0
            assert elem.width > 0
            assert elem.height > 0


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
