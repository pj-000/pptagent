import sys
import os
import re
import argparse
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from agents.orchestrator import OrchestratorAgent


def slugify(text: str) -> str:
    name = re.sub(r"[^\w\u4e00-\u9fff]", "_", text)
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:50] or "output"


def ask(prompt: str, default: str = "") -> str:
    hint = f" [{default}]" if default else ""
    val = input(f"{prompt}{hint}: ").strip()
    return val if val else default


def main():
    parser = argparse.ArgumentParser(description="PPT Agent - 自动生成 PPT")
    parser.add_argument("--topic", type=str, help="PPT 主题")
    parser.add_argument("--language", type=str, help="输出语言（默认中文）")
    parser.add_argument("--debug-layout", action="store_true", help="打印布局调试信息")
    parser.add_argument("--no-research", action="store_true", help="跳过 Tavily 搜索增强")
    parser.add_argument("--no-images", action="store_true", help="跳过图片下载，使用灰色占位")
    parser.add_argument("--slides", type=str, help="页数，如 8 或 6-10")
    args = parser.parse_args()

    # 交互式收集参数（未通过 flag 提供时）
    topic = args.topic or ask("请输入 PPT 主题")
    if not topic:
        print("主题不能为空")
        sys.exit(1)

    language = args.language or ask("输出语言", default="中文")

    # 页数：优先用 --slides flag，否则交互询问
    slides_raw = args.slides or ask("PPT 页数（如 8，或范围如 6-10）", default="6-10")
    try:
        if "-" in slides_raw:
            parts = slides_raw.split("-")
            min_slides, max_slides = int(parts[0].strip()), int(parts[1].strip())
        else:
            n = int(slides_raw)
            min_slides, max_slides = n, n
    except ValueError:
        min_slides, max_slides = 6, 10

    # 仅在完全交互模式下（没有任何 flag）才询问高级选项
    no_research = args.no_research
    no_images = args.no_images
    if not any([args.topic, args.no_research, args.no_images]):
        ans = ask("是否启用 Tavily 搜索增强内容？(y/n)", default="y").lower()
        no_research = ans != "y"
        ans2 = ask("是否下载配图？(y/n)", default="n").lower()
        no_images = ans2 != "y"

    filename = f"{slugify(topic)}.pptx"

    try:
        orchestrator = OrchestratorAgent(
            debug_layout=args.debug_layout,
            no_research=no_research,
            no_images=no_images,
        )
        output_path = orchestrator.generate(
            topic=topic,
            output_filename=filename,
            language=language,
            min_slides=min_slides,
            max_slides=max_slides,
        )
        print(f"\n完成！请用 PowerPoint 或 WPS 打开：{output_path}")
    except Exception as e:
        print(f"\n生成失败：{e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
