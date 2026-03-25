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
    parser = argparse.ArgumentParser(description="PPT Agent - 自动生成精美 PPT")
    parser.add_argument("--topic", type=str, help="PPT 主题")
    parser.add_argument("--language", type=str, help="输出语言（默认中文）")
    parser.add_argument("--debug-layout", action="store_true", help="打印内容调试信息")
    parser.add_argument("--style", type=str, help="风格：auto/executive/ocean/minimal/coral/terracotta/teal/forest/berry/cherry")
    args = parser.parse_args()

    topic = args.topic or ask("请输入 PPT 主题")
    if not topic:
        print("主题不能为空")
        sys.exit(1)

    language = args.language or ask("输出语言", default="中文")

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

    style = args.style or ask(
        "PPT 风格（auto/executive/ocean/minimal/coral/terracotta/teal/forest/berry/cherry）",
        default="auto"
    )

    filename = f"{slugify(topic)}.pptx"

    try:
        orchestrator = OrchestratorAgent(debug_layout=args.debug_layout)
        output_path = orchestrator.generate(
            topic=topic,
            output_filename=filename,
            language=language,
            min_slides=min_slides,
            max_slides=max_slides,
            style=style,
        )
        print(f"\n完成！请用 PowerPoint 或 WPS 打开：{output_path}")
    except Exception as e:
        print(f"\n生成失败：{e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
