import uuid
from pydantic import BaseModel, Field, field_validator, model_validator
from typing import Optional, List, Literal
from enum import Enum

# 幻灯片边界常量
SLIDE_WIDTH = 13.333
SLIDE_HEIGHT = 7.5
BOUNDARY_TOLERANCE = 0.01  # 允许 0.01 英寸浮点误差


class LayoutValidationError(Exception):
    """布局校验失败异常，携带详细错误列表和原始 JSON"""

    def __init__(self, errors: List[str], raw_json: str = ""):
        self.errors = errors
        self.raw_json = raw_json
        msg = f"布局校验失败（{len(errors)} 个错误）:\n" + "\n".join(f"  - {e}" for e in errors)
        super().__init__(msg)


class SlideLayout(str, Enum):
    COVER = "cover"
    TOC = "toc"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    CLOSING = "closing"


class TextElement(BaseModel):
    """单个元素，包含内容和位置信息。支持文本和图片占位。"""
    content: str = ""
    x: float
    y: float
    width: float
    height: float
    font_size: int = 18
    bold: bool = False
    color: str = "#000000"
    align: Literal["left", "center", "right"] = "left"
    type: str = "body"

    # Phase 3: 图片相关字段
    unsplash_query: Optional[str] = None
    dalle_prompt: Optional[str] = None
    local_image_path: Optional[str] = None

    @field_validator("color")
    @classmethod
    def validate_color(cls, v: str) -> str:
        if not v.startswith("#") or len(v) != 7:
            raise ValueError(f"颜色必须是 7 位十六进制格式，如 #1F3864，收到：{v}")
        return v

    @field_validator("x", "y")
    @classmethod
    def validate_non_negative(cls, v: float) -> float:
        if v < 0:
            raise ValueError("坐标必须为非负数")
        return v

    @field_validator("width", "height")
    @classmethod
    def validate_positive(cls, v: float) -> float:
        if v <= 0:
            raise ValueError("尺寸必须为正数")
        return v

    @model_validator(mode="after")
    def validate_boundary_and_image(self):
        """校验边界 + image_placeholder 必须提供 query 或 prompt"""
        right_edge = self.x + self.width
        bottom_edge = self.y + self.height
        if right_edge > SLIDE_WIDTH + BOUNDARY_TOLERANCE:
            raise ValueError(
                f"元素右边界越界：x({self.x}) + width({self.width}) = {right_edge:.3f}，"
                f"超出幻灯片宽度 {SLIDE_WIDTH}"
            )
        if bottom_edge > SLIDE_HEIGHT + BOUNDARY_TOLERANCE:
            raise ValueError(
                f"元素下边界越界：y({self.y}) + height({self.height}) = {bottom_edge:.3f}，"
                f"超出幻灯片高度 {SLIDE_HEIGHT}"
            )
        if self.type == "image_placeholder":
            if not self.unsplash_query and not self.dalle_prompt:
                raise ValueError(
                    "image_placeholder 元素必须至少提供 unsplash_query 或 dalle_prompt 之一"
                )
        return self


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

    @model_validator(mode="after")
    def validate_has_title(self):
        """每页至少有一个 type='title' 的元素"""
        has_title = any(elem.type == "title" for elem in self.elements)
        if not has_title and len(self.elements) > 0:
            raise ValueError(
                f"第 {self.slide_index} 页缺少 type='title' 的元素"
            )
        return self


def _generate_short_id() -> str:
    return uuid.uuid4().hex[:8]


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
    job_id: str = Field(default_factory=_generate_short_id)

    @field_validator("slides")
    @classmethod
    def validate_slide_count(cls, v: List[SlideSpec]) -> List[SlideSpec]:
        if len(v) < 2:
            raise ValueError("PPT 至少需要 2 页")
        if len(v) > 20:
            raise ValueError("PPT 最多 20 页")
        return v
