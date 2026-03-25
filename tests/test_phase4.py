import pytest
import os
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from pydantic import ValidationError
from models.schemas import (
    TextElement, SlideSpec, SlideLayout, PresentationPlan,
    LayoutValidationError, SLIDE_WIDTH, SLIDE_HEIGHT,
)
from tools.xml_parser import parse_presentation, _parse_slide, _parse_element
from tools.pptx_renderer import PPTXRenderer


# ─── Schema: shape 元素校验 ───


class TestShapeElementValidation:

    def test_shape_requires_shape_type(self):
        """shape 元素必须指定 shape_type"""
        with pytest.raises(ValidationError, match="shape_type"):
            TextElement(
                type="shape", x=0, y=0, width=5, height=1,
                font_size=14, color="#000000", fill_color="#FF0000",
            )

    def test_shape_requires_fill_color(self):
        """shape 元素必须指定 fill_color"""
        with pytest.raises(ValidationError, match="fill_color"):
            TextElement(
                type="shape", shape_type="rect",
                x=0, y=0, width=5, height=1,
                font_size=14, color="#000000",
            )

    def test_valid_rect_shape(self):
        """合法 rect shape 通过"""
        elem = TextElement(
            type="shape", shape_type="rect", fill_color="#1E2761",
            x=0, y=6.5, width=13.333, height=1.0,
            font_size=14, color="#000000",
        )
        assert elem.shape_type == "rect"
        assert elem.fill_color == "#1E2761"

    def test_valid_circle_shape(self):
        """合法 circle shape 通过"""
        elem = TextElement(
            type="shape", shape_type="circle", fill_color="#CADCFC",
            x=1, y=1, width=1, height=1,
            font_size=14, color="#000000",
        )
        assert elem.shape_type == "circle"

    def test_valid_line_shape(self):
        """合法 line shape 通过"""
        elem = TextElement(
            type="shape", shape_type="line", fill_color="#2E75B6",
            x=0.5, y=1.2, width=12.333, height=0.04,
            font_size=14, color="#000000",
        )
        assert elem.height == 0.04

    def test_non_shape_no_shape_type_required(self):
        """非 shape 元素不要求 shape_type"""
        elem = TextElement(
            type="body", content="正文",
            x=0.5, y=1.5, width=12.0, height=5.0,
            font_size=18, color="#333333",
        )
        assert elem.shape_type is None

    def test_fill_color_validation(self):
        """fill_color 格式错误应报错"""
        with pytest.raises(ValidationError, match="fill_color"):
            TextElement(
                type="shape", shape_type="rect", fill_color="red",
                x=0, y=0, width=5, height=1,
                font_size=14, color="#000000",
            )

    def test_callout_element(self):
        """callout 元素正常创建"""
        elem = TextElement(
            type="callout", content="85%",
            x=1, y=2, width=4, height=2,
            font_size=66, bold=True, color="#1E2761", align="center",
        )
        assert elem.type == "callout"
        assert elem.font_size == 66


# ─── Schema: SlideSpec 放宽 title 校验 ───


class TestSlideSpecShapeOnly:

    def test_shape_only_slide_no_title_ok(self):
        """纯 shape 页面不要求 title"""
        slide = SlideSpec(
            slide_index=0, layout=SlideLayout.CONTENT, topic="装饰",
            elements=[
                TextElement(
                    type="shape", shape_type="rect", fill_color="#1E2761",
                    x=0, y=0, width=13.333, height=7.5,
                    font_size=14, color="#000000",
                ),
            ],
        )
        assert len(slide.elements) == 1

    def test_text_without_title_still_fails(self):
        """有文字元素但无 title 仍然报错"""
        with pytest.raises(ValidationError, match="缺少 type='title'"):
            SlideSpec(
                slide_index=0, layout=SlideLayout.CONTENT, topic="测试",
                elements=[
                    TextElement(
                        type="body", content="正文",
                        x=0.5, y=1.5, width=12.0, height=5.0,
                        font_size=18, color="#333333",
                    ),
                ],
            )


# ─── XML Parser ───


class TestXMLParser:

    def _write_xml_files(self, tmp_path):
        """写入测试用 XML 文件"""
        xml_dir = str(tmp_path)

        # presentation.xml
        pres = ET.Element("presentation",
                          title="测试PPT", topic="测试主题",
                          theme_color="#1E2761", accent_color="#CADCFC",
                          font_family="Microsoft YaHei", slide_count="3")
        ET.ElementTree(pres).write(os.path.join(xml_dir, "presentation.xml"),
                                   encoding="unicode", xml_declaration=True)

        # slide_0.xml - cover
        s0 = ET.Element("slide", index="0", layout="cover", topic="封面",
                         background_color="#1E2761")
        ET.SubElement(s0, "element", type="shape", shape_type="rect",
                      x="0", y="6.5", width="13.333", height="1.0",
                      fill_color="#CADCFC")
        title = ET.SubElement(s0, "element", type="title",
                              x="1.5", y="2.0", width="10.333", height="1.8",
                              font_size="44", bold="true", color="#FFFFFF", align="center")
        title.text = "测试标题"
        ET.ElementTree(s0).write(os.path.join(xml_dir, "slide_0.xml"),
                                 encoding="unicode", xml_declaration=True)

        # slide_1.xml - content
        s1 = ET.Element("slide", index="1", layout="content", topic="内容页",
                         background_color="#F5F5F5")
        ET.SubElement(s1, "element", type="shape", shape_type="rect",
                      x="0", y="0", width="0.3", height="7.5",
                      fill_color="#1E2761")
        t1 = ET.SubElement(s1, "element", type="title",
                           x="0.8", y="0.3", width="12.0", height="0.9",
                           font_size="32", bold="true", color="#1E2761", align="left")
        t1.text = "内容标题"
        b1 = ET.SubElement(s1, "element", type="body",
                           x="0.8", y="1.5", width="12.0", height="5.5",
                           font_size="16", bold="false", color="#333333", align="left")
        b1.text = "• 要点一\n• 要点二"
        ET.ElementTree(s1).write(os.path.join(xml_dir, "slide_1.xml"),
                                 encoding="unicode", xml_declaration=True)

        # slide_2.xml - closing
        s2 = ET.Element("slide", index="2", layout="closing", topic="结尾",
                         background_color="#1E2761")
        ET.SubElement(s2, "element", type="shape", shape_type="rect",
                      x="0", y="0", width="13.333", height="1.0",
                      fill_color="#CADCFC")
        t2 = ET.SubElement(s2, "element", type="title",
                           x="1.5", y="2.5", width="10.333", height="1.5",
                           font_size="44", bold="true", color="#FFFFFF", align="center")
        t2.text = "感谢聆听"
        ET.ElementTree(s2).write(os.path.join(xml_dir, "slide_2.xml"),
                                 encoding="unicode", xml_declaration=True)

        return xml_dir

    def test_parse_presentation(self, tmp_path):
        """解析完整 presentation"""
        xml_dir = self._write_xml_files(tmp_path)
        plan = parse_presentation(xml_dir)
        assert isinstance(plan, PresentationPlan)
        assert plan.title == "测试PPT"
        assert len(plan.slides) == 3

    def test_parse_slide_layouts(self, tmp_path):
        """解析后布局类型正确"""
        xml_dir = self._write_xml_files(tmp_path)
        plan = parse_presentation(xml_dir)
        assert plan.slides[0].layout == SlideLayout.COVER
        assert plan.slides[1].layout == SlideLayout.CONTENT
        assert plan.slides[2].layout == SlideLayout.CLOSING

    def test_parse_background_colors(self, tmp_path):
        """解析后背景色正确"""
        xml_dir = self._write_xml_files(tmp_path)
        plan = parse_presentation(xml_dir)
        assert plan.slides[0].background_color == "#1E2761"
        assert plan.slides[1].background_color == "#F5F5F5"

    def test_parse_shape_elements(self, tmp_path):
        """解析后 shape 元素正确"""
        xml_dir = self._write_xml_files(tmp_path)
        plan = parse_presentation(xml_dir)
        shapes = [e for e in plan.slides[0].elements if e.type == "shape"]
        assert len(shapes) >= 1
        assert shapes[0].shape_type == "rect"
        assert shapes[0].fill_color == "#CADCFC"

    def test_parse_text_content(self, tmp_path):
        """解析后文字内容正确"""
        xml_dir = self._write_xml_files(tmp_path)
        plan = parse_presentation(xml_dir)
        titles = [e for e in plan.slides[0].elements if e.type == "title"]
        assert titles[0].content == "测试标题"

    def test_missing_presentation_xml_raises(self, tmp_path):
        """缺少 presentation.xml 应报错"""
        with pytest.raises(FileNotFoundError):
            parse_presentation(str(tmp_path))

    def test_parse_element_node(self):
        """直接测试 _parse_element"""
        node = ET.Element("element", type="shape", shape_type="circle",
                          x="2", y="3", width="1.5", height="1.5",
                          fill_color="#FF0000")
        elem = _parse_element(node)
        assert elem.type == "shape"
        assert elem.shape_type == "circle"
        assert elem.x == 2.0

    def test_parse_image_placeholder(self, tmp_path):
        """解析 image_placeholder 元素"""
        xml_dir = str(tmp_path)

        # presentation.xml
        pres = ET.Element("presentation", title="T", topic="T",
                          theme_color="#1E2761", accent_color="#CADCFC",
                          font_family="Microsoft YaHei", slide_count="2")
        ET.ElementTree(pres).write(os.path.join(xml_dir, "presentation.xml"),
                                   encoding="unicode", xml_declaration=True)

        # slide with image_placeholder
        s = ET.Element("slide", index="0", layout="content", topic="图片页",
                        background_color="#F5F5F5")
        ET.SubElement(s, "element", type="shape", shape_type="rect",
                      x="0", y="0", width="0.3", height="7.5", fill_color="#1E2761")
        t = ET.SubElement(s, "element", type="title",
                          x="0.8", y="0.3", width="12.0", height="0.9",
                          font_size="32", bold="true", color="#1E2761")
        t.text = "标题"
        img = ET.SubElement(s, "element", type="image_placeholder",
                            x="6.9", y="1.5", width="5.9", height="5.5",
                            font_size="14", color="#999999",
                            unsplash_query="solar panel",
                            dalle_prompt="A solar panel on rooftop")
        img.text = "太阳能板"
        ET.ElementTree(s).write(os.path.join(xml_dir, "slide_0.xml"),
                                encoding="unicode", xml_declaration=True)

        # slide_1 closing
        s1 = ET.Element("slide", index="1", layout="closing", topic="结尾",
                         background_color="#1E2761")
        t1 = ET.SubElement(s1, "element", type="title",
                           x="1.5", y="2.5", width="10.333", height="1.5",
                           font_size="44", bold="true", color="#FFFFFF", align="center")
        t1.text = "结束"
        ET.ElementTree(s1).write(os.path.join(xml_dir, "slide_1.xml"),
                                 encoding="unicode", xml_declaration=True)

        plan = parse_presentation(xml_dir)
        img_elems = [e for e in plan.slides[0].elements if e.type == "image_placeholder"]
        assert len(img_elems) == 1
        assert img_elems[0].unsplash_query == "solar panel"


# ─── Renderer: shape 渲染 ───


class TestRendererShapes:

    def test_render_with_shapes(self, tmp_path):
        """包含 shape 元素的 PPT 能正常渲染"""
        import config
        original_output = config.OUTPUT_DIR
        config.OUTPUT_DIR = str(tmp_path)

        try:
            plan = PresentationPlan(
                title="形状测试", topic="测试",
                slides=[
                    SlideSpec(
                        slide_index=0, layout=SlideLayout.COVER, topic="封面",
                        background_color="#1E2761",
                        elements=[
                            TextElement(
                                type="shape", shape_type="rect", fill_color="#CADCFC",
                                x=0, y=6.5, width=13.333, height=1.0,
                                font_size=14, color="#000000",
                            ),
                            TextElement(
                                type="title", content="测试标题",
                                x=1.5, y=2.0, width=10.333, height=1.8,
                                font_size=44, bold=True, color="#FFFFFF", align="center",
                            ),
                        ],
                    ),
                    SlideSpec(
                        slide_index=1, layout=SlideLayout.CONTENT, topic="内容",
                        background_color="#F5F5F5",
                        elements=[
                            TextElement(
                                type="shape", shape_type="circle", fill_color="#1E2761",
                                x=0.5, y=1.5, width=0.8, height=0.8,
                                font_size=14, color="#000000",
                            ),
                            TextElement(
                                type="shape", shape_type="line", fill_color="#CADCFC",
                                x=0.5, y=1.2, width=12.333, height=0.04,
                                font_size=14, color="#000000",
                            ),
                            TextElement(
                                type="title", content="内容标题",
                                x=0.5, y=0.3, width=12.0, height=0.9,
                                font_size=32, bold=True, color="#1E2761",
                            ),
                            TextElement(
                                type="body", content="• 要点一\n• 要点二",
                                x=1.8, y=1.5, width=10.5, height=5.5,
                                font_size=16, color="#333333",
                            ),
                        ],
                    ),
                    SlideSpec(
                        slide_index=2, layout=SlideLayout.CLOSING, topic="结尾",
                        background_color="#1E2761",
                        elements=[
                            TextElement(
                                type="shape", shape_type="rect", fill_color="#CADCFC",
                                x=0, y=0, width=13.333, height=1.0,
                                font_size=14, color="#000000",
                            ),
                            TextElement(
                                type="title", content="感谢聆听",
                                x=1.5, y=2.5, width=10.333, height=1.5,
                                font_size=44, bold=True, color="#FFFFFF", align="center",
                            ),
                        ],
                    ),
                ],
            )
            renderer = PPTXRenderer()
            output_path = renderer.render(plan, "test_shapes.pptx")
            assert os.path.exists(output_path)
            assert os.path.getsize(output_path) > 10 * 1024
        finally:
            config.OUTPUT_DIR = original_output

    def test_render_shape_with_text(self, tmp_path):
        """shape 内嵌文字能正常渲染"""
        import config
        original_output = config.OUTPUT_DIR
        config.OUTPUT_DIR = str(tmp_path)

        try:
            plan = PresentationPlan(
                title="callout测试", topic="测试",
                slides=[
                    SlideSpec(
                        slide_index=0, layout=SlideLayout.CONTENT, topic="数据",
                        background_color="#F5F5F5",
                        elements=[
                            TextElement(
                                type="title", content="关键数据",
                                x=0.5, y=0.3, width=12.0, height=0.9,
                                font_size=32, bold=True, color="#1E2761",
                            ),
                            TextElement(
                                type="shape", shape_type="rect", fill_color="#1E2761",
                                content="85%", x=2, y=2, width=4, height=3,
                                font_size=60, bold=True, color="#FFFFFF", align="center",
                            ),
                        ],
                    ),
                    SlideSpec(
                        slide_index=1, layout=SlideLayout.CLOSING, topic="结尾",
                        background_color="#1E2761",
                        elements=[
                            TextElement(
                                type="title", content="结束",
                                x=1.5, y=2.5, width=10.333, height=1.5,
                                font_size=44, bold=True, color="#FFFFFF", align="center",
                            ),
                        ],
                    ),
                ],
            )
            renderer = PPTXRenderer()
            output_path = renderer.render(plan, "test_callout.pptx")
            assert os.path.exists(output_path)
        finally:
            config.OUTPUT_DIR = original_output


# ─── XML → Render 端到端 ───


class TestXMLToRender:

    def test_xml_to_pptx_pipeline(self, tmp_path):
        """XML 文件 → parse → render 完整流水线"""
        import config
        original_output = config.OUTPUT_DIR

        xml_dir = os.path.join(str(tmp_path), "xml")
        os.makedirs(xml_dir)

        # 写 XML
        pres = ET.Element("presentation", title="端到端测试", topic="测试",
                          theme_color="#1E2761", accent_color="#CADCFC",
                          font_family="Microsoft YaHei", slide_count="2")
        ET.ElementTree(pres).write(os.path.join(xml_dir, "presentation.xml"),
                                   encoding="unicode", xml_declaration=True)

        s0 = ET.Element("slide", index="0", layout="cover", topic="封面",
                         background_color="#1E2761")
        ET.SubElement(s0, "element", type="shape", shape_type="rect",
                      x="0", y="6.5", width="13.333", height="1.0", fill_color="#CADCFC")
        t0 = ET.SubElement(s0, "element", type="title",
                           x="1.5", y="2.0", width="10.333", height="1.8",
                           font_size="44", bold="true", color="#FFFFFF", align="center")
        t0.text = "端到端测试"
        ET.ElementTree(s0).write(os.path.join(xml_dir, "slide_0.xml"),
                                 encoding="unicode", xml_declaration=True)

        s1 = ET.Element("slide", index="1", layout="closing", topic="结尾",
                         background_color="#1E2761")
        t1 = ET.SubElement(s1, "element", type="title",
                           x="1.5", y="2.5", width="10.333", height="1.5",
                           font_size="44", bold="true", color="#FFFFFF", align="center")
        t1.text = "结束"
        ET.SubElement(s1, "element", type="shape", shape_type="rect",
                      x="0", y="0", width="13.333", height="0.5", fill_color="#CADCFC")
        ET.ElementTree(s1).write(os.path.join(xml_dir, "slide_1.xml"),
                                 encoding="unicode", xml_declaration=True)

        # Parse
        plan = parse_presentation(xml_dir)
        assert len(plan.slides) == 2

        # Render
        config.OUTPUT_DIR = str(tmp_path)
        try:
            renderer = PPTXRenderer()
            output_path = renderer.render(plan, "test_e2e.pptx")
            assert os.path.exists(output_path)
            assert os.path.getsize(output_path) > 10 * 1024
        finally:
            config.OUTPUT_DIR = original_output


# ─── Planner code extraction (mock) ───


class TestPlannerCodeExtraction:

    def test_extract_code_tag(self):
        """从 <code>...</code> 提取代码"""
        from agents.planner import PlannerAgent
        from unittest.mock import patch, MagicMock

        with patch("agents.planner.OpenAI"):
            with patch("agents.planner.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock"
                planner = PlannerAgent()
                planner._system_template = "s {slide_width} {slide_height}"
                planner._user_template = "u {topic} {slide_width} {slide_height} {language} {min_slides} {max_slides}"

        raw = "Here is the code:\n<code>\ndef generate_slides(output_dir):\n    return 0\n</code>\nDone."
        code = planner._extract_code(raw)
        assert "def generate_slides" in code

    def test_extract_python_block(self):
        """从 ```python...``` 提取代码"""
        from agents.planner import PlannerAgent
        from unittest.mock import patch, MagicMock

        with patch("agents.planner.OpenAI"):
            with patch("agents.planner.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock"
                planner = PlannerAgent()
                planner._system_template = "s {slide_width} {slide_height}"
                planner._user_template = "u {topic} {slide_width} {slide_height} {language} {min_slides} {max_slides}"

        raw = "```python\ndef generate_slides(output_dir):\n    return 0\n```"
        code = planner._extract_code(raw)
        assert "def generate_slides" in code

    def test_no_code_raises(self):
        """无代码块应报错"""
        from agents.planner import PlannerAgent
        from unittest.mock import patch

        with patch("agents.planner.OpenAI"):
            with patch("agents.planner.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock"
                planner = PlannerAgent()
                planner._system_template = "s {slide_width} {slide_height}"
                planner._user_template = "u {topic} {slide_width} {slide_height} {language} {min_slides} {max_slides}"

        with pytest.raises(ValueError, match="未找到"):
            planner._extract_code("这里没有代码")


# ─── 集成测试 ───


class TestPhase4Integration:

    @pytest.fixture(autouse=True)
    def check_api_key(self):
        if not os.getenv("GLM_API_KEY"):
            pytest.skip("未设置 GLM_API_KEY，跳过集成测试")

    def test_full_pipeline(self):
        """完整 Phase 4 流程：Planner → XML → parse → render"""
        from agents.orchestrator import OrchestratorAgent
        orch = OrchestratorAgent(no_images=True, no_research=True)
        output = orch.generate("区块链技术应用", output_filename="test_p4.pptx")
        assert os.path.exists(output)

        # 验证 XML 文件存在
        xml_dir = os.path.join("outputs", "xml")
        xml_files = [f for f in os.listdir(xml_dir) if f.startswith("slide_")]
        assert len(xml_files) >= 2

    def test_cover_has_dark_background(self):
        """封面应使用深色背景"""
        from agents.planner import PlannerAgent
        planner = PlannerAgent()
        plan = planner.plan("人工智能发展趋势")
        cover = plan.slides[0]
        # 深色背景的 RGB 值总和应较低
        bg = cover.background_color.lstrip("#")
        r, g, b = int(bg[0:2], 16), int(bg[2:4], 16), int(bg[4:6], 16)
        assert r + g + b < 400, f"封面背景色 {cover.background_color} 不够深"

    def test_every_slide_has_shape(self):
        """每页至少有一个 shape 元素"""
        from agents.planner import PlannerAgent
        planner = PlannerAgent()
        plan = planner.plan("可再生能源")
        for slide in plan.slides:
            shapes = [e for e in slide.elements if e.type == "shape"]
            assert len(shapes) >= 1, f"第 {slide.slide_index} 页缺少 shape 元素"
