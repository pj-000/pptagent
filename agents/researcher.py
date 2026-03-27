import json
import re
import asyncio
import logging
from pathlib import Path
from openai import AsyncOpenAI
from tavily import AsyncTavilyClient
from models.schemas import SlideSpec, SlideLayout, TextElement
import config

logger = logging.getLogger(__name__)


class ResearchAgent:
    """用 Tavily 搜索每页主题，再用 GLM 提炼为 PPT 要点。"""

    SKIP_LAYOUTS = {SlideLayout.COVER, SlideLayout.CLOSING, SlideLayout.TOC}

    def __init__(self):
        self.tavily = AsyncTavilyClient(api_key=config.TAVILY_API_KEY)
        self.llm = AsyncOpenAI(api_key=config.GLM_API_KEY, base_url=config.GLM_BASE_URL)
        self.client = self.llm
        self._system_template = Path("prompts/researcher_system.txt").read_text(encoding="utf-8")

    async def research_topic(self, topic: str, language: str = "中文") -> dict:
        """
        对整份 PPT 主题做一次前置研究，供 Planner 生成时参考。
        当前主流程还没有“先产出 slides 再逐页 research”的中间态，
        所以这里用一个 synthetic content slide 复用现有 research_slide 逻辑。
        """
        slide = SlideSpec(
            slide_index=0,
            layout=SlideLayout.CONTENT,
            topic=topic,
            elements=[
                TextElement(
                    type="title",
                    content=topic,
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
        result = await self.research_slide(slide, language=language)
        return result or {
            "topic": topic,
            "summary": topic,
            "bullet_points": [],
        }

    async def research_slide(self, slide: SlideSpec, language: str = "中文") -> dict | None:
        """
        对单页做 Tavily 搜索 + GLM 提炼。
        cover / closing / toc 直接返回 None。
        失败时返回默认结构，不让全流程崩溃。
        """
        if slide.layout in self.SKIP_LAYOUTS:
            return None

        try:
            # Step 1: Tavily 搜索
            search_result = await self.tavily.search(
                query=slide.topic,
                max_results=3,
                search_depth="basic",
            )
            snippets = [r.get("content", "") for r in search_result.get("results", [])]
            context = "\n\n".join(snippets[:3])

            # Step 2: GLM 提炼为 PPT 要点
            system_prompt = self._system_template.replace("{language}", language)
            user_prompt = (
                f"页面主题：{slide.topic}\n\n"
                f"以下是搜索到的参考资料：\n{context[:2000]}\n\n"
                f"请根据以上资料，为该 PPT 页面生成精炼的内容。"
            )

            response = await self.client.chat.completions.create(
                model=config.RESEARCH_MODEL,
                max_tokens=config.MAX_TOKENS_RESEARCHER,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
            )
            raw = response.choices[0].message.content
            print(f"[Research] 第 {slide.slide_index} 页 LLM 原始响应前200字：{repr(raw[:200])}")
            data = self._parse_json(raw)

            if "bullet_points" not in data or not isinstance(data["bullet_points"], list):
                raise ValueError("缺少 bullet_points 字段")

            print(f"[Research] 第 {slide.slide_index} 页完成: {slide.topic}")
            return data

        except Exception as e:
            logger.warning(f"[Research] 第 {slide.slide_index} 页失败: {e}")
            print(f"[Research] 第 {slide.slide_index} 页失败，使用默认内容: {e}")
            return {
                "topic": slide.topic,
                "summary": slide.topic,
                "bullet_points": [],
            }

    async def research_all(
        self, slides: list[SlideSpec], language: str = "中文", concurrency: int = 3
    ) -> list[dict | None]:
        """并发研究所有页面，返回列表长度与 slides 一致。"""
        sem = asyncio.Semaphore(concurrency)

        async def _bounded(slide: SlideSpec):
            async with sem:
                return await self.research_slide(slide, language)

        results = await asyncio.gather(*[_bounded(s) for s in slides])
        return list(results)

    def _parse_json(self, raw: str) -> dict:
        cleaned = raw.strip()
        # 去掉 ```json ... ``` 围栏
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
        cleaned = cleaned.strip()

        # 尝试直接解析
        try:
            result = json.loads(cleaned)
            if isinstance(result, dict):
                return result
        except json.JSONDecodeError:
            pass

        # fallback：提取第一个 {...} 块
        match = re.search(r"\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", cleaned, re.DOTALL)
        if match:
            result = json.loads(match.group(0))
            if isinstance(result, dict):
                return result

        # 最后尝试：找最外层的 { 到最后一个 }
        first_brace = cleaned.find("{")
        last_brace = cleaned.rfind("}")
        if first_brace >= 0 and last_brace > first_brace:
            candidate = cleaned[first_brace:last_brace + 1]
            result = json.loads(candidate)
            if isinstance(result, dict):
                return result

        raise ValueError(f"无法从响应中提取 JSON dict，原始内容前200字：{cleaned[:200]}")
