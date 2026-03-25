import sys
import os
import re
import argparse
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from agents.orchestrator import OrchestratorAgent
from agents.planner import SUPPORTED_AUDIENCES, SUPPORTED_STYLES



def slugify(text: str) -> str:
    name = re.sub(r"[^\w\u4e00-\u9fff]", "_", text)
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:50] or "output"


def ask(prompt: str, default: str = "") -> str:
    hint = f" [{default}]" if default else ""
    val = input(f"{prompt}{hint}: ").strip()
    return val if val else default


def ask_yes_no(prompt: str, default: bool = True) -> bool:
    suffix = "Y/n" if default else "y/N"
    raw = input(f"{prompt} [{suffix}]: ").strip().lower()
    if not raw:
        return default
    return raw in {"y", "yes", "1", "true", "是", "开启", "开"}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="PPT Agent - 自动生成精美 PPT")
    parser.add_argument("--topic", type=str, help="PPT 主题")
    parser.add_argument("--language", type=str, help="输出语言（默认中文）")
    parser.add_argument("--slides", type=str, help="PPT 页数，如 8 或 6-10")
    parser.add_argument(
        "--style",
        type=str,
        help=f"PPT 风格，可填自由描述；常见示例：{'/'.join(SUPPORTED_STYLES)}",
    )
    parser.add_argument("--debug-layout", action="store_true", help="打印内容调试信息")
    parser.add_argument(
        "--audience",
        type=str,
        help=f"目标受众，可填自由描述；常见示例：{'/'.join(SUPPORTED_AUDIENCES)}",
    )
    research_group = parser.add_mutually_exclusive_group()
    research_group.add_argument("--research", dest="research", action="store_true", help="启用联网 Research（Tavily）")
    research_group.add_argument("--no-research", dest="research", action="store_false", help="跳过联网 Research（Tavily）")
    parser.set_defaults(research=None)
    return parser


def parse_slide_range(slides_raw: str, default: tuple[int, int] = (6, 10)) -> tuple[int, int]:
    text = (slides_raw or "").strip()
    if not text:
        return default

    range_match = re.match(r"^(\d+)\s*[-~～]\s*(\d+)$", text)
    if range_match:
        min_slides, max_slides = int(range_match.group(1)), int(range_match.group(2))
        return tuple(sorted((min_slides, max_slides)))

    if text.isdigit():
        value = int(text)
        return value, value

    return default


def main():
    parser = build_parser()
    args = parser.parse_args()

    topic = args.topic or ask("请输入 PPT 主题")
    if not topic:
        print("主题不能为空")
        sys.exit(1)

    language = args.language or ask("输出语言", default="中文")

    slides_raw = args.slides or ask("PPT 页数（如 8，或范围如 6-10）", default="6-10")
    min_slides, max_slides = parse_slide_range(slides_raw)

    style = args.style or ask(
        "PPT 风格（auto/executive/ocean/minimal/coral/terracotta/teal/forest/berry/cherry）",
        default="auto"
    )

    audience_raw = args.audience or ask(
        "目标受众（可自由输入，如 大学生 / 投资人 / 企业老板 / 技术团队）",
        default="general"
    )
    audience = audience_raw.strip() or "general"

    if args.research is None:
        research_enabled = ask_yes_no("启用联网 Research（Tavily，用于补充内容）", default=True)
    else:
        research_enabled = args.research

    filename = f"{slugify(topic)}.pptx"

    try:
        orchestrator = OrchestratorAgent(
            debug_layout=args.debug_layout,
            no_research=not research_enabled,
        )
        output_path = orchestrator.generate(
            topic=topic,
            output_filename=filename,
            language=language,
            min_slides=min_slides,
            max_slides=max_slides,
            style=style,
            audience=audience,
        )
        print(f"\n完成！请用 PowerPoint 或 WPS 打开：{output_path}")
    except Exception as e:
        print(f"\n生成失败：{e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
