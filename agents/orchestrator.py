from agents.planner import PlannerAgent
from tools.pptx_renderer import PPTXRenderer


class OrchestratorAgent:
    """
    主控 Agent。
    Phase 1：串行调用 Planner → Renderer。
    """

    def __init__(self):
        self.planner = PlannerAgent()
        self.renderer = PPTXRenderer()

    def generate(self, topic: str, output_filename: str = "output.pptx") -> str:
        """
        端到端生成 PPT。

        Args:
            topic: 用户输入的主题
            output_filename: 输出文件名

        Returns:
            生成的 .pptx 文件的绝对路径
        """
        print(f"\n{'='*50}")
        print(f"开始生成 PPT：{topic}")
        print(f"{'='*50}\n")

        # Step 1: 规划内容
        plan = self.planner.plan(topic)

        # Step 2: 渲染 PPTX
        output_path = self.renderer.render(plan, output_filename)

        print(f"\n{'='*50}")
        print(f"生成完成！文件路径：{output_path}")
        print(f"{'='*50}\n")

        return output_path
