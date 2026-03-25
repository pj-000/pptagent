import time
import asyncio
from agents.planner import PlannerAgent
from agents.researcher import ResearchAgent
from agents.asset_agent import AssetAgent
from tools.pptx_renderer import PPTXRenderer
from models.schemas import LayoutValidationError


class OrchestratorAgent:
    """
    主控 Agent。
    Phase 3：Planner → 并行(Research + Asset) → 合并 → Renderer。
    """

    def __init__(
        self,
        debug_layout: bool = False,
        no_research: bool = False,
        no_images: bool = False,
    ):
        self.planner = PlannerAgent()
        self.renderer = PPTXRenderer()
        self.researcher = None if no_research else ResearchAgent()
        self.asset_agent = None if no_images else AssetAgent()
        self.debug_layout = debug_layout

    def generate(
        self,
        topic: str,
        output_filename: str = "output.pptx",
        language: str = "中文",
        min_slides: int = 6,
        max_slides: int = 10,
    ) -> str:
        print(f"\n{'='*50}")
        print(f"开始生成 PPT：{topic}")
        print(f"{'='*50}\n")

        try:
            # Step 1: Planner
            t0 = time.time()
            plan = self.planner.plan(topic, min_slides=min_slides, max_slides=max_slides)
            t_plan = time.time() - t0
            print(f"[Orchestrator] Planner 耗时: {t_plan:.1f}s")

            # Step 2: 并行 Research + Asset
            t1 = time.time()
            research_results, _ = asyncio.run(
                self._parallel_enrich(plan, language)
            )
            t_enrich = time.time() - t1
            print(f"[Orchestrator] Research+Asset 并行耗时: {t_enrich:.1f}s")

            # Step 3: 合并 research 结果到 plan
            if research_results:
                self._merge_research(plan, research_results)

            if self.debug_layout:
                self._print_layout_debug(plan)

            # Step 4: 渲染
            t2 = time.time()
            output_path = self.renderer.render(plan, output_filename)
            t_render = time.time() - t2
            print(f"[Orchestrator] Renderer 耗时: {t_render:.1f}s")

            print(f"\n{'='*50}")
            print(f"生成完成！文件路径：{output_path}")
            print(f"总耗时: {time.time() - t0:.1f}s")
            print(f"{'='*50}\n")

            return output_path

        except LayoutValidationError as e:
            print(f"\n[Orchestrator] 布局校验失败！")
            print(f"  错误数量：{len(e.errors)}")
            for err in e.errors:
                print(f"  - {err}")
            if e.raw_json:
                print(f"  原始 JSON（前 300 字）：{e.raw_json[:300]}")
            raise

    async def _parallel_enrich(self, plan, language):
        """并行执行 Research 和 Asset。"""
        tasks = []

        if self.researcher:
            tasks.append(self.researcher.research_all(plan.slides, language=language))
        else:
            tasks.append(self._noop())

        if self.asset_agent:
            tasks.append(self.asset_agent.fetch_all(plan.slides, job_id=plan.job_id))
        else:
            tasks.append(self._noop())

        results = await asyncio.gather(*tasks, return_exceptions=True)

        research_results = results[0] if not isinstance(results[0], Exception) else None
        asset_results = results[1] if not isinstance(results[1], Exception) else None

        if isinstance(results[0], Exception):
            print(f"[Orchestrator] Research 异常（已忽略）: {results[0]}")
        if isinstance(results[1], Exception):
            print(f"[Orchestrator] Asset 异常（已忽略）: {results[1]}")

        return research_results, asset_results

    @staticmethod
    async def _noop():
        return None

    def _merge_research(self, plan, research_results):
        """把 research 的 bullet_points 合并到对应 body 元素。"""
        for slide, research in zip(plan.slides, research_results):
            if research is None:
                continue
            bullet_points = research.get("bullet_points", [])
            if not bullet_points:
                continue

            # 找到第一个 body 元素，追加 bullet points
            for elem in slide.elements:
                if elem.type == "body":
                    new_points = "\n".join(f"• {bp}" for bp in bullet_points)
                    if elem.content:
                        elem.content = elem.content + "\n" + new_points
                    else:
                        elem.content = new_points
                    break

    def _print_layout_debug(self, plan):
        print(f"\n{'─'*50}")
        print(f"[DEBUG LAYOUT] 共 {len(plan.slides)} 页, job_id={plan.job_id}")
        print(f"{'─'*50}")
        for slide in plan.slides:
            print(f"\n  第 {slide.slide_index} 页 [{slide.layout.value}] - {slide.topic}")
            print(f"  背景色: {slide.background_color}")
            for i, elem in enumerate(slide.elements):
                right = elem.x + elem.width
                bottom = elem.y + elem.height
                line = (f"    元素 {i}: type={elem.type}, "
                        f"x={elem.x:.2f}, y={elem.y:.2f}, "
                        f"w={elem.width:.2f}, h={elem.height:.2f}, "
                        f"right={right:.3f}, bottom={bottom:.3f}, "
                        f"font_size={elem.font_size}pt")
                if elem.type == "image_placeholder":
                    line += f", query={elem.unsplash_query}"
                    if elem.local_image_path:
                        line += f", img={elem.local_image_path}"
                    else:
                        line += ", img=占位"
                print(line)
        print(f"\n{'─'*50}\n")
