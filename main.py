import sys
import os
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from agents.orchestrator import OrchestratorAgent


def slugify(text: str) -> str:
    """将主题文字转换为合法文件名"""
    # 保留中文、字母、数字，其余替换为下划线
    name = re.sub(r"[^\w\u4e00-\u9fff]", "_", text)
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:50] or "output"


def main():
    topic = input("请输入 PPT 主题：").strip()
    if not topic:
        print("主题不能为空")
        sys.exit(1)

    filename = f"{slugify(topic)}.pptx"

    orchestrator = OrchestratorAgent()
    output_path = orchestrator.generate(topic=topic, output_filename=filename)
    print(f"完成！请用 PowerPoint 或 WPS 打开：{output_path}")


if __name__ == "__main__":
    main()
