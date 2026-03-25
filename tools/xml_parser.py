"""
tools/xml_parser.py
将 outputs/xml/ 下的 slide_N.xml + presentation.xml 解析为 PresentationPlan。
不调用任何 LLM。
"""
import os
import xml.etree.ElementTree as ET
from models.schemas import PresentationPlan, SlideSpec, SlideLayout, TextElement


def parse_presentation(xml_dir: str) -> PresentationPlan:
    """读取 xml_dir 下的 presentation.xml 和所有 slide_N.xml，返回 PresentationPlan。"""
    pres_path = os.path.join(xml_dir, "presentation.xml")
    if not os.path.isfile(pres_path):
        raise FileNotFoundError(f"找不到 {pres_path}")

    pres_tree = ET.parse(pres_path)
    pres_root = pres_tree.getroot()

    title = pres_root.get("title", "未命名")
    topic = pres_root.get("topic", "")
    theme_color = pres_root.get("theme_color", "#1E2761")
    accent_color = pres_root.get("accent_color", "#CADCFC")
    font_family = pres_root.get("font_family", "Microsoft YaHei")
    slide_count = int(pres_root.get("slide_count", "0"))

    # 收集所有 slide 文件并按 index 排序
    slide_files = sorted(
        [f for f in os.listdir(xml_dir) if f.startswith("slide_") and f.endswith(".xml")],
        key=lambda f: int(f.replace("slide_", "").replace(".xml", ""))
    )

    slides = []
    for sf in slide_files:
        slide_path = os.path.join(xml_dir, sf)
        slide_spec = _parse_slide(slide_path)
        slides.append(slide_spec)

    return PresentationPlan(
        title=title,
        topic=topic,
        theme_color=theme_color,
        accent_color=accent_color,
        font_family=font_family,
        slides=slides,
    )


def _parse_slide(path: str) -> SlideSpec:
    """解析单个 slide XML 文件。"""
    tree = ET.parse(path)
    root = tree.getroot()

    slide_index = int(root.get("index", "0"))
    layout_str = root.get("layout", "content")
    topic = root.get("topic", "")
    background_color = root.get("background_color", "#FFFFFF")

    # 规范化 layout
    try:
        layout = SlideLayout(layout_str)
    except ValueError:
        layout = SlideLayout.CONTENT

    elements = []
    for elem_node in root.findall("element"):
        te = _parse_element(elem_node)
        elements.append(te)

    # 读取 speaker_notes（如果有）
    notes_node = root.find("speaker_notes")
    speaker_notes = notes_node.text if notes_node is not None else None

    return SlideSpec(
        slide_index=slide_index,
        layout=layout,
        topic=topic,
        background_color=background_color,
        elements=elements,
        speaker_notes=speaker_notes,
    )


def _parse_element(node: ET.Element) -> TextElement:
    """解析单个 <element> 节点为 TextElement。"""
    elem_type = node.get("type", "body")
    content = (node.text or "").strip()

    kwargs = {
        "type": elem_type,
        "content": content,
        "x": float(node.get("x", "0")),
        "y": float(node.get("y", "0")),
        "width": float(node.get("width", "1")),
        "height": float(node.get("height", "1")),
        "font_size": int(node.get("font_size", "18")),
        "bold": node.get("bold", "false").lower() == "true",
        "color": node.get("color", "#000000"),
        "align": node.get("align", "left"),
    }

    # 形状字段
    if node.get("shape_type"):
        kwargs["shape_type"] = node.get("shape_type")
    if node.get("fill_color"):
        kwargs["fill_color"] = node.get("fill_color")
    if node.get("corner_radius"):
        kwargs["corner_radius"] = float(node.get("corner_radius"))

    # 图片字段
    if node.get("unsplash_query"):
        kwargs["unsplash_query"] = node.get("unsplash_query")
    if node.get("dalle_prompt"):
        kwargs["dalle_prompt"] = node.get("dalle_prompt")

    return TextElement(**kwargs)
