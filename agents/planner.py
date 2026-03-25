import json
import re
import logging
from pathlib import Path
from openai import OpenAI
from models.schemas import PresentationPlan, LayoutValidationError
import config

logger = logging.getLogger(__name__)

MAX_RETRIES = 3


class PlannerAgent:
    """
    调用 GLM API（OpenAI 兼容接口）生成完整的 PPT 布局规划 JSON，
    包含每个元素的精确坐标和尺寸。
    """

    def __init__(self):
        self.client = OpenAI(api_key=config.GLM_API_KEY, base_url=config.GLM_BASE_URL)
        self._system_template = Path("prompts/planner_system.txt").read_text(encoding="utf-8")
        self._user_template = Path("prompts/planner_user.txt").read_text(encoding="utf-8")

    def _build_system_prompt(self) -> str:
        return self._system_template.format(
            slide_width=config.SLIDE_WIDTH_INCH,
            slide_height=config.SLIDE_HEIGHT_INCH,
        )

    def _build_user_prompt(self, topic: str, min_slides: int = 6, max_slides: int = 10) -> str:
        return self._user_template.format(
            topic=topic,
            slide_width=config.SLIDE_WIDTH_INCH,
            slide_height=config.SLIDE_HEIGHT_INCH,
            language="中文",
            min_slides=min_slides,
            max_slides=max_slides,
        )

    def plan(self, topic: str, min_slides: int = 6, max_slides: int = 10) -> PresentationPlan:
        """
        输入主题，输出完整的 PresentationPlan（含动态坐标）。
        若 JSON 解析或校验失败，自动重试最多 MAX_RETRIES 次。
        """
        print(f"[Planner] 开始规划主题：{topic}")

        system_prompt = self._build_system_prompt()
        user_prompt = self._build_user_prompt(topic, min_slides=min_slides, max_slides=max_slides)
        errors_so_far: list[str] = []
        last_raw = ""

        for attempt in range(1, MAX_RETRIES + 1):
            print(f"[Planner] 第 {attempt}/{MAX_RETRIES} 次尝试...")

            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ]

            # 如果之前有错误，注入到对话中让模型修正
            if errors_so_far:
                error_feedback = (
                    "你上一次的输出有以下问题，请修正后重新输出完整 JSON：\n"
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

            # 尝试解析和校验
            try:
                plan_data = self._parse_json(last_raw)
                plan = PresentationPlan.model_validate(plan_data)
                print(f"[Planner] 规划完成，共 {len(plan.slides)} 页")
                return plan
            except (ValueError, json.JSONDecodeError) as e:
                err_msg = f"JSON 解析失败: {e}"
                logger.warning(f"[Planner] 第 {attempt} 次尝试失败: {err_msg}")
                print(f"[Planner] 第 {attempt} 次尝试失败: {err_msg}")
                errors_so_far = [err_msg]
            except Exception as e:
                err_msg = f"校验失败: {e}"
                logger.warning(f"[Planner] 第 {attempt} 次尝试失败: {err_msg}")
                print(f"[Planner] 第 {attempt} 次尝试失败: {err_msg}")
                errors_so_far = [err_msg]

        # 所有重试都失败
        raise LayoutValidationError(
            errors=errors_so_far,
            raw_json=last_raw,
        )

    def _parse_json(self, raw: str) -> dict:
        """
        解析 LLM 返回的 JSON。
        兼容 ```json ... ``` 包裹格式。
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
