import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from models.schemas import PresentationPlan, SlideSpec, TextElement
import config


def hex_to_rgb(hex_color: str) -> RGBColor:
    """将十六进制颜色字符串转换为 python-pptx 的 RGBColor 对象"""
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    return RGBColor(r, g, b)


ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
}


class PPTXRenderer:

    def render(self, plan: PresentationPlan, filename: str = "output.pptx") -> str:
        prs = Presentation()
        prs.slide_width = Inches(plan.slide_width)
        prs.slide_height = Inches(plan.slide_height)

        for slide_spec in plan.slides:
            self._render_slide(prs, slide_spec, plan)

        os.makedirs(config.OUTPUT_DIR, exist_ok=True)
        output_path = os.path.join(config.OUTPUT_DIR, filename)
        prs.save(output_path)
        print(f"[Renderer] 已保存：{output_path}（共 {len(plan.slides)} 页）")
        return os.path.abspath(output_path)

    def _render_slide(self, prs: Presentation, spec: SlideSpec, plan: PresentationPlan):
        blank_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_layout)

        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hex_to_rgb(spec.background_color)

        for elem in spec.elements:
            if elem.type == "image_placeholder":
                self._add_image(slide, elem)
            else:
                self._add_text_box(slide, elem, plan.font_family)

        if spec.speaker_notes:
            slide.notes_slide.notes_text_frame.text = spec.speaker_notes

    def _add_text_box(self, slide, elem: TextElement, font_family: str):
        txBox = slide.shapes.add_textbox(
            Inches(elem.x),
            Inches(elem.y),
            Inches(elem.width),
            Inches(elem.height)
        )
        tf = txBox.text_frame
        tf.word_wrap = True

        lines = elem.content.split("\n")
        for i, line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.alignment = ALIGN_MAP.get(elem.align, PP_ALIGN.LEFT)
            run = p.add_run()
            run.text = line

            font = run.font
            font.name = font_family
            font.size = Pt(elem.font_size)
            font.bold = elem.bold
            font.color.rgb = hex_to_rgb(elem.color)

    def _add_image(self, slide, elem: TextElement):
        """插入图片或灰色占位矩形。"""
        left = Inches(elem.x)
        top = Inches(elem.y)
        width = Inches(elem.width)
        height = Inches(elem.height)

        if elem.local_image_path and os.path.isfile(elem.local_image_path):
            try:
                slide.shapes.add_picture(
                    elem.local_image_path, left, top, width, height
                )
                return
            except Exception as e:
                print(f"[Renderer] 图片插入失败，使用占位: {e}")

        # 灰色占位矩形
        from pptx.util import Emu
        shape = slide.shapes.add_shape(
            1,  # MSO_SHAPE.RECTANGLE
            left, top, width, height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
        shape.line.fill.background()

        # 占位文字
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        desc = elem.content or elem.unsplash_query or "图片"
        run.text = f"[图片] {desc}"
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
