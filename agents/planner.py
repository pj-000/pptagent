import json
import re
import os
import logging
from pathlib import Path
from openai import OpenAI
from tools.pptx_skill import run_js, skill_paths, assert_skill_present
from models.schemas import OutlinePlan, SlideOutline, SlideLayout
import config

logger = logging.getLogger(__name__)

MAX_RETRIES = 3
SUPPORTED_STYLES = (
    "auto", "executive", "ocean", "minimal", "coral",
    "terracotta", "teal", "forest", "berry", "cherry",
)
SUPPORTED_AUDIENCES = ("general", "boss", "investor", "student", "customer", "technical")

AUDIENCE_ALIASES = {
    "": "general",
    "general": "general",
    "default": "general",
    "public": "general",
    "大众": "general",
    "通用": "general",
    "普通": "general",
    "boss": "boss",
    "leader": "boss",
    "leaders": "boss",
    "manager": "boss",
    "management": "boss",
    "executive": "boss",
    "ceo": "boss",
    "老板": "boss",
    "领导": "boss",
    "管理层": "boss",
    "投资人": "investor",
    "投资者": "investor",
    "vc": "investor",
    "investor": "investor",
    "student": "student",
    "students": "student",
    "learner": "student",
    "learners": "student",
    "college student": "student",
    "university student": "student",
    "学员": "student",
    "学生": "student",
    "大学生": "student",
    "高校学生": "student",
    "customer": "customer",
    "customers": "customer",
    "client": "customer",
    "clients": "customer",
    "user": "customer",
    "users": "customer",
    "客户": "customer",
    "用户": "customer",
    "消费者": "customer",
    "technical": "technical",
    "tech": "technical",
    "engineer": "technical",
    "engineers": "technical",
    "developer": "technical",
    "developers": "technical",
    "技术": "technical",
    "工程师": "technical",
    "开发者": "technical",
}

AUDIENCE_PROFILES = {
    "general": "信息密度中等，先讲背景再讲重点，少用术语，兼顾概念解释与关键结论。",
    "boss": "结论先行，高信息压缩，比细节更重视决策点、ROI、风险、资源投入与下一步建议。",
    "investor": "突出市场空间、增长逻辑、竞争优势、商业化路径、关键指标与风险回报，语气克制专业。",
    "student": "循序渐进，术语要解释，例子要具体，信息密度偏低到中等，强调学习路径与核心概念。",
    "customer": "以用户价值、应用场景、收益、差异化和可信度为中心，少讲内部机制，避免堆砌术语。",
    "technical": "允许更高信息密度，强调原理、架构、实现方式、性能取舍、限制条件与技术准确性。",
}


def normalize_audience(audience: str | None) -> str:
    return (audience or "").strip() or "general"


def suggest_audience_label(audience: str | None) -> str | None:
    raw = normalize_audience(audience)
    normalized = AUDIENCE_ALIASES.get(raw.lower())
    if normalized:
        return normalized
    return raw.lower() if raw.lower() in AUDIENCE_PROFILES else None


def suggest_style_label(style: str | None) -> str | None:
    raw = (style or "").strip()
    if not raw:
        return "auto"
    lowered = raw.lower()
    return lowered if lowered in SUPPORTED_STYLES else None


SHAPE_VALUE_MAP = {
    "rect": "rect",
    "rectangle": "rect",
    "square": "rect",
    "oval": "ellipse",
    "ellipse": "ellipse",
    "circle": "ellipse",
    "line": "line",
    "roundrect": "roundRect",
    "roundedrect": "roundRect",
    "roundedrectangle": "roundRect",
    "rounded_rectangle": "roundRect",
    "chevron": "chevron",
    "diamond": "diamond",
    "triangle": "triangle",
    "hexagon": "hexagon",
    "arc": "arc",
    "pie": "pie",
    "cloud": "cloud",
    "pentagon": "pentagon",
    "rightarrow": "rightArrow",
    "right_arrow": "rightArrow",
    "arrow": "rightArrow",
    "star5": "star5",
    "star5point": "star5",
    "star_5_point": "star5",
}


# 本地 vendored skill 路径
_SKILL = skill_paths()


class PlannerAgent:
    """
    读取本地 Anthropic PPTX skill 文档作为 system prompt，
    让 GLM 生成 PptxGenJS 代码，通过 pptx_skill.run_js() 执行。
    """

    def __init__(self):
        assert_skill_present()
        self.client = OpenAI(api_key=config.GLM_API_KEY, base_url=config.GLM_BASE_URL)

        # 直接读取本地 vendored skill 文档
        self._skill_md = Path(_SKILL["skill_md"]).read_text(encoding="utf-8")
        self._pptxgenjs_md = Path(_SKILL["pptxgenjs_md"]).read_text(encoding="utf-8")
        self._local_rules_md = Path(_SKILL["local_rules_md"]).read_text(encoding="utf-8")
        self._user_template = Path("prompts/planner_user.txt").read_text(encoding="utf-8")

    def _build_system_prompt(self) -> str:
        """
        system prompt = 角色指令 + 官方 SKILL.md Before Starting + Design Ideas + Typography + Avoid + 完整 pptxgenjs.md
        """
        # 从 SKILL.md 中提取 "Before Starting"（在 Design Ideas 下）到 QA 之前的内容
        skill = self._skill_md
        start_marker = "### Before Starting"
        end_marker = "## QA"
        start_idx = skill.find(start_marker)
        end_idx = skill.find(end_marker)
        if start_idx >= 0 and end_idx > start_idx:
            design_section = skill[start_idx:end_idx].strip()
        elif start_idx >= 0:
            design_section = skill[start_idx:].strip()
        else:
            # fallback: Design Ideas 之后
            idx = skill.find("## Design Ideas")
            design_section = skill[idx:].strip() if idx >= 0 else skill

        return f"""你是一位顶级 PPT 视觉设计工程师。你使用 PptxGenJS（Node.js）生成精美的演示文稿。

以下是来自 Anthropic 官方 PPTX Design Skill 的设计规范，你必须严格遵守：

---
{design_section}
---

以下是本地增强的生成规则（硬约束 + 视觉设计原则 + 布局词汇表），你也必须严格遵守：

---
{self._local_rules_md}
---

以下是 PptxGenJS 的完整 API 教程，你生成的代码必须严格遵循这些用法和注意事项：

---
{self._pptxgenjs_md}
---

## 你的输出格式

输出一段完整的 Node.js 代码，用 <code> 标签包裹。代码要求：

1. `const pptxgen = require("pptxgenjs");` 开头
2. 使用 `pres.layout = "LAYOUT_WIDE";`（13.33" × 7.5"）
3. 最后调用 `pres.writeFile({{ fileName: "OUTPUT_PATH" }});`
4. 只使用 pptxgenjs，不要其他 npm 包
5. 严格遵循上面所有设计规则，包括字体配对、字号规范、间距、以及 Avoid 清单
6. 严格遵循上面 PptxGenJS Tutorial 中的所有 API 用法和 Common Pitfalls
"""

    def _build_outline_system_prompt(self) -> str:
        return """你是一位 PPT 内容架构师，负责先规划页级大纲，再交给后续模块做研究和设计。

你的任务是输出一个严格的 JSON，对应如下结构：
{
  "title": "整份 PPT 标题",
  "topic": "用户主题",
  "slides": [
    {
      "slide_index": 0,
      "layout": "cover",
      "topic": "本页主题",
      "objective": "本页想传达的目标"
    }
  ]
}

规则：
1. 只输出 JSON，不要 markdown，不要解释
2. `layout` 只能是：cover、toc、content、two_column、closing
3. 第 0 页必须是 cover，第 1 页必须是 toc，最后一页必须是 closing
4. 中间页需要围绕主题形成清晰叙事，页间不要重复
5. `topic` 要具体到适合单页研究和展开
6. `objective` 用一句话说明本页任务
7. 幻灯片总页数必须落在用户要求范围内
"""

    def _build_outline_user_prompt(
        self,
        topic: str,
        min_slides: int = 6,
        max_slides: int = 10,
        style: str = "auto",
        audience: str = "general",
        language: str = "中文",
    ) -> str:
        audience = normalize_audience(audience)
        audience_ref = suggest_audience_label(audience) or "无"
        style = (style or "").strip() or "auto"
        style_ref = suggest_style_label(style) or "无"
        return (
            f"主题：{topic}\n"
            f"语言：{language}\n"
            f"风格原文：{style}\n"
            f"风格参考标签：{style_ref}\n"
            f"受众原文：{audience}\n"
            f"受众参考标签：{audience_ref}\n"
            "请优先理解并保留用户输入的原始风格和受众描述，参考标签仅用于辅助，不得覆盖原意。\n"
            f"受众适配提示：{self._build_audience_profile(audience)}\n"
            f"页数范围：{min_slides}-{max_slides}\n\n"
            "请先规划一份页级大纲，让后续 Research Agent 能逐页研究。"
        )

    def _build_user_prompt(self, topic: str, min_slides: int = 6, max_slides: int = 10,
                           style: str = "auto", audience: str = "general",
                           language: str = "中文", research_context: str = "",
                           outline_context: str = "") -> str:
        audience = normalize_audience(audience)
        style = (style or "").strip() or "auto"
        audience_profile = self._build_audience_profile(audience)
        return (self._user_template
                .replace("{topic}", topic)
                .replace("{slide_width}", str(config.SLIDE_WIDTH_INCH))
                .replace("{slide_height}", str(config.SLIDE_HEIGHT_INCH))
                .replace("{language}", language)
                .replace("{style}", style)
                .replace("{audience}", audience)
                .replace("{audience_profile}", audience_profile)
                .replace("{style_reference}", suggest_style_label(style) or "无")
                .replace("{audience_reference}", suggest_audience_label(audience) or "无")
                .replace("{outline_context}", outline_context.strip() or "无显式页级大纲，请自行规划结构。")
                .replace("{research_context}", research_context.strip() or "无额外研究资料，请基于常识与准确性完成。")
                .replace("{min_slides}", str(min_slides))
                .replace("{max_slides}", str(max_slides)))

    def plan_outline(
        self,
        topic: str,
        min_slides: int = 6,
        max_slides: int = 10,
        style: str = "auto",
        audience: str = "general",
        language: str = "中文",
    ) -> OutlinePlan:
        print(f"[Planner] 开始规划页级大纲：{topic}")
        system_prompt = self._build_outline_system_prompt()
        user_prompt = self._build_outline_user_prompt(
            topic,
            min_slides=min_slides,
            max_slides=max_slides,
            style=style,
            audience=audience,
            language=language,
        )

        last_raw = ""
        for attempt in range(1, MAX_RETRIES + 1):
            print(f"[Planner] 大纲规划第 {attempt}/{MAX_RETRIES} 次尝试...")
            response = self.client.chat.completions.create(
                model=config.PLANNER_MODEL,
                max_tokens=4096,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            )
            last_raw = response.choices[0].message.content

            try:
                data = self._extract_json(last_raw)
                outline = self._parse_outline_plan(data, topic)
                print(f"[Planner] 大纲规划完成，共 {len(outline.slides)} 页")
                return outline
            except Exception as e:
                err_msg = str(e)
                logger.warning(f"[Planner] 大纲规划第 {attempt} 次失败: {err_msg[:200]}")
                print(f"[Planner] 大纲规划第 {attempt} 次失败: {err_msg[:200]}")

        raise RuntimeError(f"页级大纲规划失败，最后响应前500字：{last_raw[:500]}")

    def plan(self, topic: str, output_path: str = None,
             min_slides: int = 6, max_slides: int = 10,
             style: str = "auto", audience: str = "general",
             language: str = "中文", research_context: str = "",
             outline: OutlinePlan | None = None,
             research_results: list[dict | None] | None = None) -> str:
        """
        生成 PptxGenJS 代码并执行，直接产出 .pptx 文件。

        Returns:
            生成的 .pptx 文件绝对路径
        """
        if output_path is None:
            os.makedirs(config.OUTPUT_DIR, exist_ok=True)
            output_path = os.path.join(config.OUTPUT_DIR, "output.pptx")

        output_path = os.path.abspath(output_path)
        print(f"[Planner] 开始规划主题：{topic}")

        system_prompt = self._build_system_prompt()
        user_prompt = self._build_user_prompt(
            topic,
            min_slides,
            max_slides,
            style=style,
            audience=audience,
            language=language,
            outline_context=self._build_outline_context(outline),
            research_context=research_context or self._build_research_context(outline, research_results),
        )
        errors_so_far: list[str] = []
        last_raw = ""

        for attempt in range(1, MAX_RETRIES + 1):
            print(f"[Planner] 第 {attempt}/{MAX_RETRIES} 次尝试...")

            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ]

            if errors_so_far:
                error_feedback = (
                    "你上一次输出的代码执行失败，请修正后重新输出完整的 <code>...</code>：\n"
                    + "\n".join(f"- {e}" for e in errors_so_far)
                )
                messages.append({"role": "user", "content": error_feedback})

            response = self.client.chat.completions.create(
                model=config.PLANNER_MODEL,
                max_tokens=config.MAX_TOKENS_PLANNER,
                messages=messages,
            )

            last_raw = response.choices[0].message.content
            print(f"[Planner] API 调用完成，usage: {response.usage}")

            try:
                code = self._extract_code(last_raw)
                print(f"[Planner] 提取到代码（{len(code)} 字符）")

                sanitized_code = self._sanitize_generated_code(code)
                full_code = self._inject_output_path(sanitized_code, output_path)
                run_js(full_code, output_path)
                print(f"[Planner] 生成完成: {output_path}")
                return output_path

            except Exception as e:
                err_msg = str(e)
                errors_so_far = self._build_error_feedback(err_msg)
                logger.warning(f"[Planner] 第 {attempt} 次失败: {err_msg[:200]}")
                print(f"[Planner] 第 {attempt} 次失败: {err_msg[:200]}")

        raise RuntimeError(
            f"连续 {MAX_RETRIES} 次生成失败。\n"
            f"最后错误: {errors_so_far}\n"
            f"最后响应（前500字）: {last_raw[:500]}"
        )

    def outline_to_research_slides(self, outline: OutlinePlan) -> list:
        return [
            self._make_research_slide(slide)
            for slide in outline.slides
        ]

    def _build_audience_profile(self, audience: str) -> str:
        reference = suggest_audience_label(audience)
        lines = [
            "- 必须优先理解并保留用户输入的原始受众描述，不能粗暴压缩成固定桶。",
            "- `style` 主要控制视觉风格；如果 `style` 与 `audience` 有冲突，优先保证内容表达适配受众。",
            f"- 原始受众描述：{audience}",
            f"- 可选参考标签：{reference or '无'}",
        ]
        if reference:
            lines.append(f"- 可参考的适配策略：{AUDIENCE_PROFILES[reference]}")
        else:
            lines.append("- 如果原始描述不属于常见标签，请根据字面含义自行判断术语深度、语气和信息密度。")
        return "\n".join(lines)

    def _build_outline_context(self, outline: OutlinePlan | None) -> str:
        if not outline:
            return ""
        lines = [
            "以下是已经确认的页级大纲，请严格按此结构组织内容：",
        ]
        for slide in outline.slides:
            lines.append(
                f"- 第 {slide.slide_index} 页 | layout={slide.layout.value} | topic={slide.topic} | objective={slide.objective}"
            )
        return "\n".join(lines)

    def _build_research_context(
        self,
        outline: OutlinePlan | None,
        research_results: list[dict | None] | None,
    ) -> str:
        if not outline or not research_results:
            return ""

        lines = ["以下是逐页 ResearchAgent 结果，请吸收为对应页面的事实依据和表达素材："]
        for slide, result in zip(outline.slides, research_results):
            if not result:
                continue
            summary = result.get("summary") or slide.topic
            bullet_points = result.get("bullet_points") or []
            lines.append(f"- 第 {slide.slide_index} 页 {slide.topic}")
            lines.append(f"  - 摘要：{summary}")
            for point in bullet_points:
                lines.append(f"  - 要点：{point}")
        return "\n".join(lines)

    def _parse_outline_plan(self, data: dict, topic: str) -> OutlinePlan:
        slides = data.get("slides")
        if not isinstance(slides, list):
            raise ValueError("大纲 JSON 缺少 slides 数组")

        normalized_slides = []
        for idx, slide in enumerate(slides):
            if not isinstance(slide, dict):
                raise ValueError("slides 内元素必须是对象")
            normalized_slides.append(
                {
                    "slide_index": slide.get("slide_index", idx),
                    "layout": slide.get("layout", "content"),
                    "topic": slide.get("topic") or f"{topic} - 第{idx + 1}页",
                    "objective": slide.get("objective", ""),
                }
            )

        normalized = {
            "title": data.get("title") or topic,
            "topic": data.get("topic") or topic,
            "slides": normalized_slides,
        }
        outline = OutlinePlan.model_validate(normalized)
        self._validate_outline_structure(outline)
        return outline

    def _validate_outline_structure(self, outline: OutlinePlan) -> None:
        if not outline.slides:
            raise ValueError("大纲不能为空")
        first = outline.slides[0].layout.value
        second = outline.slides[1].layout.value if len(outline.slides) > 1 else None
        last = outline.slides[-1].layout.value
        if first != "cover":
            raise ValueError("第 0 页必须是 cover")
        if second != "toc":
            raise ValueError("第 1 页必须是 toc")
        if last != "closing":
            raise ValueError("最后一页必须是 closing")
        for idx, slide in enumerate(outline.slides):
            if slide.slide_index != idx:
                raise ValueError("slide_index 必须从 0 开始连续递增")
        middle_layouts = [s.layout for s in outline.slides[2:-1]]
        if middle_layouts and not any(layout in {SlideLayout.CONTENT, SlideLayout.TWO_COLUMN} for layout in middle_layouts):
            raise ValueError("中间页至少需要一个 content 或 two_column 布局")

    def _make_research_slide(self, slide: SlideOutline):
        from models.schemas import SlideSpec, TextElement

        return SlideSpec(
            slide_index=slide.slide_index,
            layout=slide.layout,
            topic=slide.topic,
            elements=[
                TextElement(
                    type="title",
                    content=slide.topic,
                    x=0.5,
                    y=0.3,
                    width=12.0,
                    height=0.9,
                    font_size=32,
                    bold=True,
                    color="#1F3864",
                )
            ],
        )

    def _extract_code(self, raw: str) -> str:
        """从 LLM 响应中提取 <code>...</code> 或 ```javascript...``` 中的代码。"""
        m = re.search(r"<code>(.*?)</code>", raw, re.DOTALL)
        if m:
            return m.group(1).strip()

        m = re.search(r"```(?:javascript|js)\s*(.*?)```", raw, re.DOTALL)
        if m:
            return m.group(1).strip()

        m = re.search(r"```\s*(.*?)```", raw, re.DOTALL)
        if m:
            return m.group(1).strip()

        raise ValueError("LLM 响应中未找到 <code> 或 ```javascript 代码块")

    def _extract_json(self, raw: str) -> dict:
        cleaned = raw.strip()
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            match = re.search(r"(\{.*\})", cleaned, re.DOTALL)
            if not match:
                raise
            return json.loads(match.group(1))

    def _build_error_feedback(self, err_msg: str) -> list[str]:
        feedback = [err_msg[:500]]
        if "Missing/Invalid shape parameter" in err_msg:
            feedback.append(
                "修正所有 `addShape()` 的第一个参数：请使用合法形状值，例如 "
                "`\"rect\"`、`\"ellipse\"`、`\"line\"`、`\"roundRect\"`，"
                "或等价的 `pres.shapes.RECTANGLE/OVAL/LINE/ROUNDED_RECTANGLE`。"
            )
        return feedback

    def _sanitize_generated_code(self, code: str) -> str:
        def shape_literal(token: str) -> str | None:
            normalized = re.sub(r"[^A-Za-z0-9]+", "", token).lower()
            value = SHAPE_VALUE_MAP.get(normalized)
            return f'"{value}"' if value else None

        def replace_shape_member(match: re.Match) -> str:
            token = match.group(1)
            return shape_literal(token) or match.group(0)

        def replace_addshape_string(match: re.Match) -> str:
            prefix, quote, token, suffix = match.groups()
            return f"{prefix}{shape_literal(token) or f'{quote}{token}{quote}'}{suffix}"

        code = re.sub(
            r"\b[\w$]+\.(?:ShapeType|shapes)\.([A-Za-z0-9_]+)\b",
            replace_shape_member,
            code,
        )
        code = re.sub(
            r"(addShape\(\s*)([\"'])([A-Za-z0-9_]+)\2(\s*,)",
            replace_addshape_string,
            code,
        )
        return code

    def _inject_output_path(self, code: str, output_path: str) -> str:
        """确保代码中的 writeFile 使用正确的 output_path。"""
        safe_path = output_path.replace("\\", "/")

        if "writeFile" in code or "writeToFile" in code:
            code = re.sub(
                r'fileName\s*:\s*["\'][^"\']*["\']',
                f'fileName: "{safe_path}"',
                code
            )
            return code

        code += f'\npres.writeFile({{ fileName: "{safe_path}" }});\n'
        return code
