import time
import os
import asyncio
from agents.planner import PlannerAgent
from agents.researcher import ResearchAgent
from tools.pptx_skill import read_pptx
import config


class OrchestratorAgent:
    """
    主控 Agent。
    Planner 生成 PptxGenJS 代码 → pptx_skill 执行 → 直接产出 .pptx
    """

    def __init__(
        self,
        debug_layout: bool = False,
        no_research: bool = False,
        no_images: bool = False,
    ):
        self.planner = PlannerAgent()
        self.researcher = ResearchAgent()
        self.debug_layout = debug_layout
        self.no_research = no_research

    def generate(
        self,
        topic: str,
        output_filename: str = "output.pptx",
        language: str = "中文",
        min_slides: int = 6,
        max_slides: int = 10,
        style: str = "auto",
        audience: str = "general",
    ) -> str:
        print(f"\n{'='*50}")
        print(f"开始生成 PPT：{topic}")
        print(f"{'='*50}\n")

        os.makedirs(config.OUTPUT_DIR, exist_ok=True)
        output_path = os.path.abspath(os.path.join(config.OUTPUT_DIR, output_filename))

        try:
            outline = self.planner.plan_outline(
                topic,
                min_slides=min_slides,
                max_slides=max_slides,
                style=style,
                audience=audience,
                language=language,
            )

            research_results = None
            if not self.no_research:
                research_results = self._research_outline(outline, language)
            t0 = time.time()
            result_path = self.planner.plan(
                topic,
                output_path=output_path,
                language=language,
                min_slides=min_slides,
                max_slides=max_slides,
                style=style,
                audience=audience,
                outline=outline,
                research_results=research_results,
            )
            elapsed = time.time() - t0
            print(f"[Orchestrator] 总耗时: {elapsed:.1f}s")

            if self.debug_layout:
                self._print_debug(result_path)

            print(f"\n{'='*50}")
            print(f"生成完成！文件路径：{result_path}")
            print(f"{'='*50}\n")

            return result_path

        except Exception as e:
            print(f"\n[Orchestrator] 生成失败: {e}")
            raise

    def _research_outline(self, outline, language: str) -> list[dict | None]:
        """对页级大纲逐页 research，返回与 slides 对齐的研究结果。"""
        print("[Orchestrator] ResearchAgent 逐页研究中...")
        try:
            slides = self.planner.outline_to_research_slides(outline)
            results = asyncio.run(self.researcher.research_all(slides, language=language))
        except Exception as e:
            print(f"[Orchestrator] ResearchAgent 跳过: {e}")
            return []

        researched_pages = sum(1 for result in results if result and result.get("bullet_points"))
        print(f"[Orchestrator] ResearchAgent 完成，{researched_pages} 页拿到研究要点")
        return results

    def _print_debug(self, pptx_path: str):
        """用 markitdown 提取内容做调试输出"""
        content = read_pptx(pptx_path)
        if content:
            print(f"\n{'─'*50}")
            print("[DEBUG] 提取的文本内容：")
            print(content[:2000])
            print(f"{'─'*50}\n")
        else:
            print("[DEBUG] markitdown 不可用，跳过内容提取")
