from pydantic import BaseModel, Field, field_validator
from typing import Optional, List, Literal
from enum import Enum


class SlideLayout(str, Enum):
    COVER = "cover"
    TOC = "toc"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    CLOSING = "closing"


class TextElement(BaseModel):
    """单个文本元素，包含内容和位置信息"""
    content: str
    x: float          # 距幻灯片左边缘，单位：英寸
    y: float          # 距幻灯片上边缘，单位：英寸
    width: float      # 文本框宽度，单位：英寸
    height: float     # 文本框高度，单位：英寸
    font_size: int    # 字号，单位：pt
    bold: bool = False
    color: str = "#000000"
    align: Literal["left", "center", "right"] = "left"

    @field_validator("color")
    @classmethod
    def validate_color(cls, v: str) -> str:
        if not v.startswith("#") or len(v) != 7:
            raise ValueError(f"颜色必须是 7 位十六进制格式，如 #1F3864，收到：{v}")
        return v

    @field_validator("x", "y", "width", "height")
    @classmethod
    def validate_positive(cls, v: float) -> float:
        if v < 0:
            raise ValueError("坐标和尺寸必须为非负数")
        return v


class SlideSpec(BaseModel):
    """单页幻灯片的完整规格"""
    slide_index: int
    layout: SlideLayout
    topic: str
    background_color: str = "#FFFFFF"
    elements: List[TextElement] = []
    speaker_notes: Optional[str] = None

    @field_validator("background_color")
    @classmethod
    def validate_bg_color(cls, v: str) -> str:
        if not v.startswith("#") or len(v) != 7:
            raise ValueError(f"背景色格式错误：{v}")
        return v


class PresentationPlan(BaseModel):
    """完整 PPT 规划，从 Planner 输出，交给 Renderer 渲染"""
    title: str
    topic: str
    slide_width: float = 13.333
    slide_height: float = 7.5
    theme_color: str = "#1F3864"
    accent_color: str = "#2E75B6"
    font_family: str = "Microsoft YaHei"
    slides: List[SlideSpec] = []

    @field_validator("slides")
    @classmethod
    def validate_slide_count(cls, v: List[SlideSpec]) -> List[SlideSpec]:
        if len(v) < 2:
            raise ValueError("PPT 至少需要 2 页")
        if len(v) > 20:
            raise ValueError("PPT 最多 20 页")
        return v
