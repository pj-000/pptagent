from models.schemas import SlideLayout, TextElement, SlideSpec


def apply_template(layout: SlideLayout, content: dict, slide_index: int) -> SlideSpec:
    """
    根据布局类型和内容字典，生成带有硬编码坐标的 SlideSpec。

    content 字典的字段：
      - title: str（标题文字）
      - subtitle: str（副标题，仅 COVER / CLOSING 使用）
      - body: str 或 list（正文内容）
      - left: str（双栏左侧，仅 TWO_COLUMN 使用）
      - right: str（双栏右侧，仅 TWO_COLUMN 使用）
    """
    elements = []

    if layout == SlideLayout.COVER:
        elements = [
            TextElement(
                content=content.get("title", ""),
                x=1.0, y=2.2, width=11.333, height=1.5,
                font_size=44, bold=True,
                color="#1F3864", align="center"
            ),
            TextElement(
                content=content.get("subtitle", ""),
                x=1.0, y=4.0, width=11.333, height=0.8,
                font_size=24, bold=False,
                color="#2E75B6", align="center"
            ),
        ]

    elif layout == SlideLayout.TOC:
        items = content.get("body", [])
        if isinstance(items, str):
            items = items.split("\n")
        toc_text = "\n".join(
            f"◆  {item.strip()}" for item in items if item.strip()
        )
        elements = [
            TextElement(
                content=content.get("title", "目录"),
                x=0.5, y=0.3, width=12.333, height=0.9,
                font_size=32, bold=True,
                color="#1F3864", align="left"
            ),
            TextElement(
                content="─" * 60,
                x=0.5, y=1.1, width=12.333, height=0.2,
                font_size=10, color="#2E75B6", align="left"
            ),
            TextElement(
                content=toc_text,
                x=1.0, y=1.5, width=11.333, height=5.5,
                font_size=20, color="#333333", align="left"
            ),
        ]

    elif layout == SlideLayout.CONTENT:
        body = content.get("body", "")
        if isinstance(body, list):
            body = "\n".join(f"• {item}" for item in body)
        elements = [
            TextElement(
                content=content.get("title", ""),
                x=0.5, y=0.3, width=12.333, height=0.9,
                font_size=32, bold=True,
                color="#1F3864", align="left"
            ),
            TextElement(
                content="─" * 60,
                x=0.5, y=1.1, width=12.333, height=0.2,
                font_size=10, color="#2E75B6", align="left"
            ),
            TextElement(
                content=body,
                x=0.5, y=1.5, width=12.333, height=5.5,
                font_size=18, color="#333333", align="left"
            ),
        ]

    elif layout == SlideLayout.TWO_COLUMN:
        elements = [
            TextElement(
                content=content.get("title", ""),
                x=0.5, y=0.3, width=12.333, height=0.9,
                font_size=32, bold=True,
                color="#1F3864", align="left"
            ),
            TextElement(
                content="─" * 60,
                x=0.5, y=1.1, width=12.333, height=0.2,
                font_size=10, color="#2E75B6", align="left"
            ),
            TextElement(
                content=content.get("left", ""),
                x=0.5, y=1.5, width=5.9, height=5.5,
                font_size=17, color="#333333", align="left"
            ),
            TextElement(
                content=content.get("right", ""),
                x=6.9, y=1.5, width=5.9, height=5.5,
                font_size=17, color="#333333", align="left"
            ),
        ]

    elif layout == SlideLayout.CLOSING:
        elements = [
            TextElement(
                content=content.get("title", "感谢聆听"),
                x=1.0, y=2.5, width=11.333, height=1.2,
                font_size=44, bold=True,
                color="#1F3864", align="center"
            ),
            TextElement(
                content=content.get("subtitle", ""),
                x=1.0, y=4.0, width=11.333, height=0.8,
                font_size=20, color="#2E75B6", align="center"
            ),
        ]

    return SlideSpec(
        slide_index=slide_index,
        layout=layout,
        topic=content.get("title", ""),
        elements=elements,
        background_color="#FFFFFF"
    )
