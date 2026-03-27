import json
import re
import os
import logging
from pathlib import Path
from openai import OpenAI
from tools.pptx_skill import run_js, skill_paths, assert_skill_present
from models.schemas import OutlinePlan, SlideOutline, SlideLayout, SlideEvalResult
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
      "objective": "本页想传达的目标",
      "image_prompt": ""
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
8. `image_prompt`：content/two_column 页必须填写一句英文视觉描述（15-40词），描述具体画面主体、场景、氛围，适合图片搜索或 AI 生图；cover/toc/closing 页留空字符串 ""
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
                           outline_context: str = "", image_context: str = "") -> str:
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
                .replace("{image_context}", image_context.strip() or "无可用图片。")
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

    # ------------------------------------------------------------------ #
    #  逐页生成                                                             #
    # ------------------------------------------------------------------ #

    def decide_visual_theme(
        self,
        outline: OutlinePlan,
        style: str = "auto",
        audience: str = "general",
        language: str = "中文",
    ) -> dict:
        """
        一次 LLM 调用，确定整份 PPT 的视觉母题。
        返回 dict，包含 primary_color / secondary_color / accent_color /
        header_font / body_font / motif_description / pres_init_code。
        """
        style = (style or "").strip() or "auto"
        system = (
            "你是 PPT 视觉设计师。根据主题和风格，输出整份 PPT 的视觉母题 JSON。\n"
            "只输出 JSON，不要解释。格式：\n"
            "{\n"
            '  "primary_color": "1E2761",\n'
            '  "secondary_color": "CADCFC",\n'
            '  "accent_color": "FFFFFF",\n'
            '  "header_font": "Georgia",\n'
            '  "body_font": "Calibri",\n'
            '  "motif_description": "左侧深色色带 + 圆形图标 + 卡片内容区",\n'
            '  "pres_init_code": "pres.layout = \\"LAYOUT_WIDE\\";"\n'
            "}\n\n"
            "颜色不要带 # 号。primary_color 明度要低（深色），accent_color 要高对比。"
        )
        slide_topics = "\n".join(
            f"- 第{s.slide_index}页 [{s.layout.value}] {s.topic}"
            for s in outline.slides
        )
        user = (
            f"主题：{outline.title}\n"
            f"风格：{style}\n"
            f"受众：{audience}\n"
            f"语言：{language}\n\n"
            f"页面列表：\n{slide_topics}\n\n"
            "请选择与主题高度匹配的配色和视觉母题，不要默认蓝色。"
        )
        try:
            resp = self.client.chat.completions.create(
                model=config.PLANNER_MODEL,
                max_tokens=512,
                messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
            )
            data = self._extract_json(resp.choices[0].message.content)
            if isinstance(data, dict) and "primary_color" in data:
                print(f"[Planner] 视觉母题：{data.get('motif_description', '')}")
                return data
        except Exception as e:
            logger.warning(f"[Planner] decide_visual_theme 失败，使用默认: {e}")
        return {
            "primary_color": "1F3864",
            "secondary_color": "2E75B6",
            "accent_color": "FFFFFF",
            "header_font": "Arial Black",
            "body_font": "Calibri",
            "motif_description": "深色封面 + 浅色内容页 + 左侧色带装饰",
            "pres_init_code": 'pres.layout = "LAYOUT_WIDE";',
        }

    def plan_slide(
        self,
        slide: SlideOutline,
        theme: dict,
        research: dict | None,
        image_path: str | None,
        prev_slides_summary: str = "",
        revision_feedback: SlideEvalResult | None = None,
    ) -> str:
        """
        为单页生成 PptxGenJS 代码片段（不含 require / pres 初始化 / writeFile）。
        返回形如 `{ let slide = pres.addSlide(); ... }` 的代码块字符串。
        """
        SKIP = {SlideLayout.COVER, SlideLayout.TOC, SlideLayout.CLOSING}

        system = self._build_slide_system_prompt()

        # 构建本页 user prompt
        parts = [
            f"## 整份 PPT 视觉母题",
            f"- 主色：#{theme.get('primary_color', '1F3864')}",
            f"- 辅色：#{theme.get('secondary_color', '2E75B6')}",
            f"- 点缀色：#{theme.get('accent_color', 'FFFFFF')}",
            f"- 标题字体：{theme.get('header_font', 'Arial Black')}",
            f"- 正文字体：{theme.get('body_font', 'Calibri')}",
            f"- 视觉母题：{theme.get('motif_description', '')}",
            "",
            f"## 本页信息",
            f"- 页码：第 {slide.slide_index} 页",
            f"- 布局类型：{slide.layout.value}",
            f"- 页面主题：{slide.topic}",
            f"- 页面目标：{slide.objective}",
        ]

        if slide.image_prompt:
            parts.append(f"- 图片描述：{slide.image_prompt}")

        if research and slide.layout not in SKIP:
            summary = research.get("summary", "")
            bullets = research.get("bullet_points", [])
            key_data = research.get("key_data", [])
            if summary:
                parts.append(f"\n## 研究摘要\n{summary}")
            if bullets:
                parts.append("\n## 要点（必须完整呈现）")
                for b in bullets:
                    parts.append(f"- {b}")
            if key_data:
                parts.append("\n## 核心数据（用 stat-callout 大字展示）")
                for d in key_data:
                    parts.append(f"- {d}")

        if image_path:
            parts.append(f"\n## 图片资产\n本页有图片，必须用 addImage({{ path: \"{image_path}\" }}) 插入，优先双栏布局（左文右图）。")

        if prev_slides_summary:
            parts.append(f"\n## 已生成页布局摘要（避免重复）\n{prev_slides_summary}")

        if revision_feedback:
            parts.append(f"\n## 上一版问题（必须修复）")
            for issue in revision_feedback.issues:
                parts.append(f"- 问题：{issue}")
            for sug in revision_feedback.suggestions:
                parts.append(f"- 建议：{sug}")

        parts.append(
            "\n## 输出要求\n"
            "只输出本页的代码片段，用 <code> 标签包裹。\n"
            "代码以 `{` 开头，以 `}` 结尾，内部第一行是 `let slide = pres.addSlide();`。\n"
            "不要包含 require、new pptxgen()、pres.layout、writeFile。\n"
            "严格遵守视觉母题的配色和字体，保持与其他页一致。\n\n"
            "## 画布尺寸（必须遵守）\n"
            "画布为 LAYOUT_WIDE：13.33\" × 7.5\"（英寸）。\n"
            "- 背景色块/形状必须覆盖整个画布（x:0, y:0, w:13.33, h:7.5）\n"
            "- 内容区域使用 0.5\" 外边距（x 从 0.5 开始，最大宽度 12.33\"）\n"
            "- 不要把所有元素挤在画布中间一小块区域\n"
            "- 封面/结尾页的标题和装饰必须铺满全屏，不能只占中间一小块"
        )

        user = "\n".join(parts)

        last_raw = ""
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self.client.chat.completions.create(
                    model=config.PLANNER_MODEL,
                    max_tokens=4096,
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user", "content": user},
                    ],
                )
                last_raw = resp.choices[0].message.content
                code = self._extract_code(last_raw)
                code = code.strip()
                if not code.startswith("{"):
                    code = "{\n" + code + "\n}"
                print(f"[Planner] 第 {slide.slide_index} 页生成成功（{len(code)} 字符）")
                return code
            except Exception as e:
                err_msg = str(e)
                print(f"[Planner] 第 {slide.slide_index} 页第 {attempt} 次失败: {err_msg[:150]}")
                logger.warning(f"[Planner] 第{slide.slide_index}页第{attempt}次失败: {e}")

        raise RuntimeError(f"第{slide.slide_index}页生成失败，最后响应：{last_raw[:300]}")

    def plan_all_slides(
        self,
        outline: OutlinePlan,
        research_results: list[dict | None] | None,
        image_paths: list[str | None] | None,
        style: str = "auto",
        audience: str = "general",
        language: str = "中文",
    ) -> tuple[list[str], dict]:
        """
        逐页生成代码片段。
        返回 (slide_codes, theme)，slide_codes 与 outline.slides 一一对应。
        """
        research_results = research_results or [None] * len(outline.slides)
        image_paths = image_paths or [None] * len(outline.slides)

        print("[Planner] 确定视觉母题...")
        theme = self.decide_visual_theme(outline, style=style, audience=audience, language=language)

        slide_codes: list[str] = []
        prev_summary_lines: list[str] = []

        for i, slide in enumerate(outline.slides):
            print(f"[Planner] 生成第 {slide.slide_index} 页（{slide.layout.value}: {slide.topic}）...")
            research = research_results[i] if i < len(research_results) else None
            img = image_paths[i] if i < len(image_paths) else None
            prev_summary = "\n".join(prev_summary_lines[-5:])  # 最近5页

            code = self.plan_slide(slide, theme, research, img, prev_summary)
            slide_codes.append(code)
            prev_summary_lines.append(f"第{slide.slide_index}页 [{slide.layout.value}] {slide.topic}")

        print(f"[Planner] 逐页生成完成，共 {len(slide_codes)} 页")
        return slide_codes, theme

    def assemble_pptx(
        self,
        slide_codes: list[str],
        output_path: str,
        theme: dict,
    ) -> str:
        """
        把所有页代码片段组装成完整 JS，执行生成 .pptx。
        """
        output_path = os.path.abspath(output_path)
        safe_path = output_path.replace("\\", "/")

        header_font = theme.get("header_font", "Arial Black")
        body_font = theme.get("body_font", "Calibri")

        lines = [
            'const pptxgen = require("pptxgenjs");',
            "let pres = new pptxgen();",
            'pres.layout = "LAYOUT_WIDE";',
            f'pres.title = "Presentation";',
            f'// 视觉母题：{theme.get("motif_description", "")}',
            f'// 主色：#{theme.get("primary_color", "")}  辅色：#{theme.get("secondary_color", "")}  点缀：#{theme.get("accent_color", "")}',
            f'// 字体：{header_font} / {body_font}',
            "",
        ]

        for i, code in enumerate(slide_codes):
            lines.append(f"// ===== 第 {i} 页 =====")
            lines.append(code)
            lines.append("")

        lines.append(f'pres.writeFile({{ fileName: "{safe_path}" }});')

        full_code = "\n".join(lines)
        full_code = self._sanitize_generated_code(full_code)

        total_chars = sum(len(c) for c in slide_codes)
        print(f"[Planner] 组装完成（{len(slide_codes)} 页，{total_chars} 字符），执行生成 PPTX...")
        run_js(full_code, output_path)
        print(f"[Planner] PPTX 生成成功: {output_path}")
        return output_path

    def _build_slide_system_prompt(self) -> str:
        """单页生成的 system prompt：设计规范 + API 教程，要求只输出片段。"""
        skill = self._skill_md
        start_idx = skill.find("### Before Starting")
        end_idx = skill.find("## QA")
        if start_idx >= 0 and end_idx > start_idx:
            design_section = skill[start_idx:end_idx].strip()
        elif start_idx >= 0:
            design_section = skill[start_idx:].strip()
        else:
            idx = skill.find("## Design Ideas")
            design_section = skill[idx:].strip() if idx >= 0 else skill

        return f"""你是一位顶级 PPT 视觉设计工程师，使用 PptxGenJS（Node.js）生成单页幻灯片代码片段。

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

输出单页代码片段，用 <code> 标签包裹。要求：
1. 代码以 `{{` 开头，以 `}}` 结尾
2. 内部第一行必须是 `let slide = pres.addSlide();`
3. 不要包含 `require`、`new pptxgen()`、`pres.layout`、`writeFile`
4. 严格遵守用户提供的视觉母题配色和字体
5. 严格遵循上面所有设计规则和 PptxGenJS API 用法
"""

    def plan(self, topic: str, output_path: str = None,
             min_slides: int = 6, max_slides: int = 10,
             style: str = "auto", audience: str = "general",
             language: str = "中文", research_context: str = "",
             outline: OutlinePlan | None = None,
             research_results: list[dict | None] | None = None,
             image_paths: list[str | None] | None = None) -> tuple[str, list[str], dict]:
        """
        逐页生成 PptxGenJS 代码并执行，直接产出 .pptx 文件。

        Returns:
            (output_path, slide_codes, theme)
        """
        if output_path is None:
            os.makedirs(config.OUTPUT_DIR, exist_ok=True)
            output_path = os.path.join(config.OUTPUT_DIR, "output.pptx")

        output_path = os.path.abspath(output_path)

        if outline is None:
            raise ValueError("plan() 需要传入 outline，请先调用 plan_outline()")

        slide_codes, theme = self.plan_all_slides(
            outline,
            research_results=research_results,
            image_paths=image_paths,
            style=style,
            audience=audience,
            language=language,
        )
        result_path = self.assemble_pptx(slide_codes, output_path, theme)
        return result_path, slide_codes, theme

    def enrich_image_prompts(
        self,
        outline: OutlinePlan,
        research_results: list[dict | None],
    ) -> OutlinePlan:
        """
        用 research 结果重新生成每页的 image_prompt。
        单次批量 LLM 调用，返回更新后的 OutlinePlan。
        cover/toc/closing 页保持空。
        """
        SKIP = {SlideLayout.COVER, SlideLayout.TOC, SlideLayout.CLOSING}

        # 构造输入：只包含需要生成 image_prompt 的页
        items = []
        for slide, result in zip(outline.slides, research_results or []):
            if slide.layout in SKIP:
                continue
            summary = (result or {}).get("summary", "")
            bullets = (result or {}).get("bullet_points", [])
            items.append({
                "slide_index": slide.slide_index,
                "topic": slide.topic,
                "summary": summary,
                "bullets": bullets[:3],  # 只取前3条，控制 token
            })

        if not items:
            return outline

        print(f"[Planner] 基于 research 重写 image_prompt（{len(items)} 页）...")
        system = (
            "你是图片描述专家。根据每页的研究摘要，为每页输出一句15-40词的英文图片搜索描述，"
            "描述具体画面主体、场景、氛围，适合图片搜索或 AI 生图。\n"
            "只输出 JSON 数组，格式：[{\"slide_index\": 0, \"image_prompt\": \"...\"}]"
        )
        user = "以下是各页研究内容，请为每页生成 image_prompt：\n" + json.dumps(items, ensure_ascii=False)

        try:
            response = self.client.chat.completions.create(
                model=config.PLANNER_MODEL,
                max_tokens=1024,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
            )
            raw = response.choices[0].message.content
            enriched = self._extract_json(raw)
            if not isinstance(enriched, list):
                raise ValueError("期望 JSON 数组")

            prompt_map = {item["slide_index"]: item.get("image_prompt", "") for item in enriched}

            updated_slides = []
            for slide in outline.slides:
                if slide.slide_index in prompt_map and prompt_map[slide.slide_index]:
                    updated_slides.append(slide.model_copy(update={"image_prompt": prompt_map[slide.slide_index]}))
                else:
                    updated_slides.append(slide)

            print(f"[Planner] image_prompt 已基于 research 更新，覆盖 {len(prompt_map)} 页")
            return outline.model_copy(update={"slides": updated_slides})

        except Exception as e:
            import logging
            logging.getLogger(__name__).warning(f"[Planner] enrich_image_prompts 失败，保留原始 image_prompt: {e}")
            return outline

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

        lines = ["以下是逐页 ResearchAgent 结果，请充分吸收为对应页面的事实依据和表达素材："]
        for slide, result in zip(outline.slides, research_results):
            if not result:
                continue
            summary = result.get("summary") or slide.topic
            bullet_points = result.get("bullet_points") or []
            key_data = result.get("key_data") or []
            lines.append(f"- 第 {slide.slide_index} 页 {slide.topic}")
            lines.append(f"  - 摘要：{summary}")
            for point in bullet_points:
                lines.append(f"  - 要点：{point}")
            for kd in key_data:
                lines.append(f"  - 核心数据（适合大字展示）：{kd}")
        return "\n".join(lines)

    def _build_image_context(
        self,
        outline: OutlinePlan | None,
        image_paths: list[str | None] | None,
    ) -> str:
        if not outline or not image_paths:
            return ""

        lines = [
            "以下是 AssetAgent 为每页准备的本地图片路径，生成代码时用 addImage({ path: '...' }) 插入：",
            "图片建议位置：x=7.0, y=1.2, w=5.8, h=5.6（英寸），适合双栏右侧。",
            "如果某页布局不适合双栏，可自行调整坐标，但必须使用提供的图片路径。",
        ]
        has_image = False
        for slide, path in zip(outline.slides, image_paths):
            if path:
                lines.append(f"- 第 {slide.slide_index} 页（{slide.topic}）：{path}")
                has_image = True

        if not has_image:
            return ""
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
                    "image_prompt": slide.get("image_prompt") or None,
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

        # 替换中文引号为转义的 ASCII 引号，避免 JS 语法错误
        # 例如: "正式命名"深度学习"" → "正式命名\"深度学习\""
        code = code.replace("\u201c", '\\"')  # "
        code = code.replace("\u201d", '\\"')  # "
        code = code.replace("\u2018", "\\'")  # '
        code = code.replace("\u2019", "\\'")  # '

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
