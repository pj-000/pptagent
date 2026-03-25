import pytest
import os
import sys
import json

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unittest.mock import MagicMock, patch
from pydantic import ValidationError
from models.schemas import (
    TextElement, SlideSpec, SlideLayout, PresentationPlan,
    LayoutValidationError, SLIDE_WIDTH, SLIDE_HEIGHT,
)
from agents.planner import PlannerAgent


# ─── ElementSpec 坐标校验 ───


class TestElementBoundary:

    def test_valid_coordinates_pass(self):
        """合法坐标应通过校验"""
        elem = TextElement(
            content="测试", type="title",
            x=0.5, y=0.3, width=12.333, height=0.9,
            font_size=32, bold=True, color="#1F3864", align="left",
        )
        assert elem.x == 0.5
        assert elem.width == 12.333

    def test_right_edge_overflow_raises(self):
        """右边界超出时应报错"""
        with pytest.raises(ValidationError, match="右边界越界"):
            TextElement(
                content="测试", type="title",
                x=10.0, y=0.0, width=5.0, height=1.0,
                font_size=32, color="#000000",
            )

    def test_bottom_edge_overflow_raises(self):
        """下边界超出时应报错"""
        with pytest.raises(ValidationError, match="下边界越界"):
            TextElement(
                content="测试", type="title",
                x=0.0, y=5.0, width=5.0, height=4.0,
                font_size=32, color="#000000",
            )

    def test_exact_boundary_passes(self):
        """恰好贴边界时应通过"""
        elem = TextElement(
            content="测试", type="title",
            x=0.0, y=0.0, width=SLIDE_WIDTH, height=SLIDE_HEIGHT,
            font_size=32, color="#000000",
        )
        assert elem.x + elem.width == SLIDE_WIDTH
        assert elem.y + elem.height == SLIDE_HEIGHT

    def test_within_tolerance_passes(self):
        """在浮点误差容忍范围内应通过"""
        elem = TextElement(
            content="测试", type="title",
            x=0.0, y=0.0, width=13.34, height=7.5,
            font_size=32, color="#000000",
        )
        assert elem is not None

    def test_negative_x_raises(self):
        """负数 x 坐标应报错"""
        with pytest.raises(ValidationError, match="非负"):
            TextElement(
                content="测试", type="title",
                x=-1.0, y=0.0, width=5.0, height=1.0,
                font_size=32, color="#000000",
            )

    def test_zero_width_raises(self):
        """宽度为 0 应报错"""
        with pytest.raises(ValidationError, match="正数"):
            TextElement(
                content="测试", type="title",
                x=0.0, y=0.0, width=0, height=1.0,
                font_size=32, color="#000000",
            )


# ─── SlideSpec title 元素校验 ───


class TestSlideSpecValidation:

    def test_no_title_element_raises(self):
        """无 title 元素的 slide 应校验失败"""
        with pytest.raises(ValidationError, match="缺少 type='title'"):
            SlideSpec(
                slide_index=0,
                layout=SlideLayout.CONTENT,
                topic="测试",
                elements=[
                    TextElement(
                        content="正文", type="body",
                        x=0.5, y=1.5, width=12.0, height=5.0,
                        font_size=18, color="#333333",
                    )
                ],
            )

    def test_with_title_element_passes(self):
        """有 title 元素的 slide 应通过"""
        slide = SlideSpec(
            slide_index=0,
            layout=SlideLayout.CONTENT,
            topic="测试",
            elements=[
                TextElement(
                    content="标题", type="title",
                    x=0.5, y=0.3, width=12.0, height=0.9,
                    font_size=32, bold=True, color="#1F3864",
                ),
                TextElement(
                    content="正文", type="body",
                    x=0.5, y=1.5, width=12.0, height=5.0,
                    font_size=18, color="#333333",
                ),
            ],
        )
        assert len(slide.elements) == 2


# ─── PresentationPlan 页数边界 ───


class TestPresentationPlanValidation:

    def _make_slide(self, index, layout=SlideLayout.CONTENT):
        return SlideSpec(
            slide_index=index,
            layout=layout,
            topic=f"页 {index}",
            elements=[
                TextElement(
                    content=f"标题 {index}", type="title",
                    x=0.5, y=0.3, width=12.0, height=0.9,
                    font_size=32, bold=True, color="#1F3864",
                )
            ],
        )

    def test_too_few_slides_raises(self):
        """少于 2 页应报错"""
        with pytest.raises(ValidationError, match="至少需要 2 页"):
            PresentationPlan(
                title="测试", topic="测试",
                slides=[self._make_slide(0)],
            )

    def test_too_many_slides_raises(self):
        """超过 20 页应报错"""
        with pytest.raises(ValidationError, match="最多 20 页"):
            PresentationPlan(
                title="测试", topic="测试",
                slides=[self._make_slide(i) for i in range(21)],
            )

    def test_valid_slide_count_passes(self):
        """正常页数应通过"""
        plan = PresentationPlan(
            title="测试", topic="测试",
            slides=[self._make_slide(i) for i in range(5)],
        )
        assert len(plan.slides) == 5


# ─── LayoutValidationError ───


class TestLayoutValidationError:

    def test_carries_errors_and_raw_json(self):
        """LayoutValidationError 应携带详细错误和原始 JSON"""
        errors = ["右边界越界", "缺少 title 元素"]
        raw = '{"bad": "json"}'
        exc = LayoutValidationError(errors=errors, raw_json=raw)
        assert exc.errors == errors
        assert exc.raw_json == raw
        assert "2 个错误" in str(exc)
        assert "右边界越界" in str(exc)

    def test_empty_raw_json(self):
        """raw_json 可以为空"""
        exc = LayoutValidationError(errors=["错误1"])
        assert exc.raw_json == ""


# ─── Planner 重试逻辑（mock） ───


MOCK_VALID_PLAN = json.dumps({
    "title": "Python编程入门",
    "topic": "Python编程入门",
    "theme_color": "#1F3864",
    "accent_color": "#2E75B6",
    "slides": [
        {
            "slide_index": 0, "layout": "cover", "topic": "封面",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "Python编程入门", "x": 1.5, "y": 2.0,
                 "width": 10.333, "height": 1.8, "font_size": 48, "bold": True,
                 "color": "#1F3864", "align": "center"},
                {"type": "subtitle", "content": "从零开始学编程", "x": 2.0, "y": 4.2,
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
                {"type": "body", "content": "◆ 基础语法\n◆ 数据类型\n◆ 控制流程",
                 "x": 1.0, "y": 1.5, "width": 11.333, "height": 5.5,
                 "font_size": 20, "bold": False, "color": "#333333", "align": "left"},
            ]
        },
        {
            "slide_index": 2, "layout": "content", "topic": "基础语法",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "Python 基础语法", "x": 0.5, "y": 0.3,
                 "width": 12.333, "height": 0.9, "font_size": 32, "bold": True,
                 "color": "#1F3864", "align": "left"},
                {"type": "body", "content": "• 变量与赋值\n• 缩进规则\n• 注释写法",
                 "x": 0.5, "y": 1.5, "width": 12.333, "height": 5.5,
                 "font_size": 18, "bold": False, "color": "#333333", "align": "left"},
            ]
        },
        {
            "slide_index": 3, "layout": "two_column", "topic": "数据类型对比",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "数据类型对比", "x": 0.5, "y": 0.3,
                 "width": 12.333, "height": 0.9, "font_size": 32, "bold": True,
                 "color": "#1F3864", "align": "left"},
                {"type": "body", "content": "基本类型\n• int\n• float\n• str",
                 "x": 0.5, "y": 1.5, "width": 5.9, "height": 5.5,
                 "font_size": 17, "bold": False, "color": "#333333", "align": "left"},
                {"type": "body", "content": "容器类型\n• list\n• dict\n• tuple",
                 "x": 6.9, "y": 1.5, "width": 5.9, "height": 5.5,
                 "font_size": 17, "bold": False, "color": "#333333", "align": "left"},
            ]
        },
        {
            "slide_index": 4, "layout": "content", "topic": "控制流程",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "控制流程", "x": 0.6, "y": 0.4,
                 "width": 12.0, "height": 0.9, "font_size": 30, "bold": True,
                 "color": "#1F3864", "align": "left"},
                {"type": "body", "content": "• if/elif/else 条件判断\n• for 循环\n• while 循环",
                 "x": 0.6, "y": 1.6, "width": 12.0, "height": 5.4,
                 "font_size": 18, "bold": False, "color": "#333333", "align": "left"},
            ]
        },
        {
            "slide_index": 5, "layout": "closing", "topic": "结尾",
            "background_color": "#FFFFFF",
            "elements": [
                {"type": "title", "content": "感谢聆听", "x": 1.5, "y": 2.2,
                 "width": 10.333, "height": 1.5, "font_size": 44, "bold": True,
                 "color": "#1F3864", "align": "center"},
                {"type": "subtitle", "content": "欢迎交流与提问", "x": 2.0, "y": 4.2,
                 "width": 9.333, "height": 0.8, "font_size": 22, "bold": False,
                 "color": "#2E75B6", "align": "center"},
            ]
        },
    ]
}, ensure_ascii=False)


@pytest.fixture
def mock_planner():
    """返回一个 mock 掉 API 调用的 PlannerAgent"""
    with patch("agents.planner.OpenAI") as mock_openai:
        with patch("agents.planner.Path") as mock_path:
            mock_path.return_value.read_text.return_value = "mock prompt"
            planner = PlannerAgent()
            planner.client = MagicMock()
            planner._system_template = "system {slide_width} {slide_height}"
            planner._user_template = "user {topic} {slide_width} {slide_height} {language} {min_slides} {max_slides}"
            yield planner


def _mock_response(content: str):
    """构造 mock API 响应"""
    mock_resp = MagicMock()
    mock_resp.choices = [MagicMock()]
    mock_resp.choices[0].message.content = content
    mock_resp.usage = MagicMock()
    return mock_resp


class TestPlannerRetry:

    def test_valid_response_first_try(self, mock_planner):
        """第一次就返回合法 JSON 应直接成功"""
        mock_planner.client.chat.completions.create.return_value = _mock_response(MOCK_VALID_PLAN)
        plan = mock_planner.plan("Python编程入门")
        assert isinstance(plan, PresentationPlan)
        assert len(plan.slides) == 6
        assert mock_planner.client.chat.completions.create.call_count == 1

    def test_retry_on_invalid_json(self, mock_planner):
        """第一次返回非法 JSON，第二次成功"""
        mock_planner.client.chat.completions.create.side_effect = [
            _mock_response("这不是 JSON"),
            _mock_response(MOCK_VALID_PLAN),
        ]
        plan = mock_planner.plan("Python编程入门")
        assert isinstance(plan, PresentationPlan)
        assert mock_planner.client.chat.completions.create.call_count == 2

    def test_all_retries_fail_raises_layout_error(self, mock_planner):
        """连续 3 次失败应抛出 LayoutValidationError"""
        mock_planner.client.chat.completions.create.return_value = _mock_response("bad json")
        with pytest.raises(LayoutValidationError) as exc_info:
            mock_planner.plan("Python编程入门")
        assert len(exc_info.value.errors) > 0
        assert exc_info.value.raw_json == "bad json"

    def test_json_wrapped_in_markdown(self, mock_planner):
        """```json 包裹的响应也能正确解析"""
        wrapped = f"```json\n{MOCK_VALID_PLAN}\n```"
        mock_planner.client.chat.completions.create.return_value = _mock_response(wrapped)
        plan = mock_planner.plan("Python编程入门")
        assert isinstance(plan, PresentationPlan)


# ─── 集成测试（需要真实 API） ───


class TestPlannerIntegration:

    @pytest.fixture(autouse=True)
    def check_api_key(self):
        if not os.getenv("GLM_API_KEY"):
            pytest.skip("未设置 GLM_API_KEY，跳过集成测试")

    def test_first_slide_is_cover(self):
        """第一页必须是 cover"""
        planner = PlannerAgent()
        plan = planner.plan("Python编程入门")
        assert plan.slides[0].layout == SlideLayout.COVER

    def test_all_elements_within_boundary(self):
        """所有元素必须在边界内"""
        planner = PlannerAgent()
        plan = planner.plan("数据科学基础")
        for slide in plan.slides:
            for elem in slide.elements:
                assert elem.x >= 0
                assert elem.y >= 0
                assert elem.x + elem.width <= SLIDE_WIDTH + 0.01
                assert elem.y + elem.height <= SLIDE_HEIGHT + 0.01

    def test_different_topics_different_layouts(self):
        """不同主题生成的布局不能完全相同"""
        planner = PlannerAgent()
        plan_a = planner.plan("量子计算入门")
        plan_b = planner.plan("美食烹饪技巧")

        # 收集所有元素坐标
        def extract_coords(plan):
            coords = []
            for slide in plan.slides:
                for elem in slide.elements:
                    coords.append((round(elem.x, 2), round(elem.y, 2),
                                   round(elem.width, 2), round(elem.height, 2)))
            return coords

        coords_a = extract_coords(plan_a)
        coords_b = extract_coords(plan_b)
        # 至少有一些坐标不同
        assert coords_a != coords_b, "两个不同主题的布局坐标完全相同，说明布局没有变化"

    def test_last_slide_is_closing(self):
        """最后一页必须是 closing"""
        planner = PlannerAgent()
        plan = planner.plan("现代艺术鉴赏")
        assert plan.slides[-1].layout == SlideLayout.CLOSING

    def test_every_slide_has_title_element(self):
        """每页都有 title 元素"""
        planner = PlannerAgent()
        plan = planner.plan("机器学习入门")
        for slide in plan.slides:
            has_title = any(e.type == "title" for e in slide.elements)
            assert has_title, f"第 {slide.slide_index} 页缺少 title 元素"
