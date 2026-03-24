import json
import re
from pathlib import Path
from openai import OpenAI
from models.schemas import PresentationPlan, SlideLayout
from tools.templates import apply_template
import config


class PlannerAgent:
    """
    调用 GLM API（OpenAI 兼容接口）生成 PPT 大纲和内容，
    然后将内容与固定模板坐标合并，生成完整的 PresentationPlan。
    """

    def __init__(self):
        self.client = OpenAI(api_key=config.GLM_API_KEY, base_url=config.GLM_BASE_URL)
        self.system_prompt = Path("prompts/planner_system.txt").read_text(encoding="utf-8")
        self.user_template = Path("prompts/planner_user.txt").read_text(encoding="utf-8")

    def plan(self, topic: str) -> PresentationPlan:
        """
        输入主题，输出完整的 PresentationPlan。
        """
        print(f"[Planner] 开始规划主题：{topic}")

        user_message = self.user_template.format(topic=topic)

        response = self.client.chat.completions.create(
            model=config.PLANNER_MODEL,
            max_tokens=config.MAX_TOKENS_PLANNER,
            messages=[
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": user_message}
            ]
        )

        raw_response = response.choices[0].message.content
        print(f"[Planner] API 调用完成，usage: {response.usage}")

        plan_data = self._parse_json(raw_response)
        presentation_plan = self._build_plan(plan_data)
        print(f"[Planner] 规划完成，共 {len(presentation_plan.slides)} 页")
        return presentation_plan

    def _parse_json(self, raw: str) -> dict:
        """
        解析 LLM 返回的 JSON。
        LLM 有时会在 JSON 外面包一层 ```json ... ```，需要去掉。
        """
        cleaned = re.sub(r"^```(?:json)?\s*", "", raw.strip())
        cleaned = re.sub(r"\s*```$", "", cleaned)

        try:
            return json.loads(cleaned)
        except json.JSONDecodeError as e:
            raise ValueError(
                f"LLM 返回的内容不是合法 JSON。\n"
                f"错误：{e}\n"
                f"原始内容（前 500 字）：{raw[:500]}"
            )

    def _build_plan(self, data: dict) -> PresentationPlan:
        """
        将 LLM 输出的 JSON 数据 + 固定模板坐标合并，
        构建完整的 PresentationPlan（含每个元素的 x/y/w/h）。
        """
        slides = []
        for slide_data in data.get("slides", []):
            layout_str = slide_data.get("layout", "content")

            try:
                layout = SlideLayout(layout_str)
            except ValueError:
                print(f"[Planner] 警告：未知布局类型 '{layout_str}'，使用 content 代替")
                layout = SlideLayout.CONTENT

            content = slide_data.get("content", {})
            slide_index = slide_data.get("slide_index", len(slides))

            slide_spec = apply_template(layout, content, slide_index)

            if "notes" in slide_data:
                slide_spec.speaker_notes = slide_data["notes"]

            slides.append(slide_spec)

        return PresentationPlan(
            title=data.get("title", "未命名演示"),
            topic=data.get("topic", ""),
            theme_color=data.get("theme_color", "#1F3864"),
            accent_color=data.get("accent_color", "#2E75B6"),
            slides=slides
        )
