"""
tools/pptx_skill.py

本地封装 Anthropic 官方 PPTX skill（vendor/anthropic_pptx_skill）。

能力：
1. run_js(code, output_path) — 执行 PptxGenJS 代码生成 .pptx
2. read_pptx(path) — 用 vendored SKILL 推荐的 markitdown 提取文本
3. pptx_to_images(path) — 调用 vendored scripts/office/soffice.py + pdftoppm 转图
4. skill_paths() — 返回本地 skill 路径，便于 planner 读取本地文档
"""
import os
import subprocess
import tempfile
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SKILL_ROOT = PROJECT_ROOT / "vendor" / "anthropic_pptx_skill"
SCRIPTS_ROOT = SKILL_ROOT / "scripts"
OFFICE_ROOT = SCRIPTS_ROOT / "office"

# 全局 node_modules 路径
_NPM_PREFIX = subprocess.run(
    ["npm", "config", "get", "prefix"],
    capture_output=True, text=True, timeout=10
).stdout.strip()
NODE_PATH = os.path.join(_NPM_PREFIX, "lib", "node_modules")


def skill_paths() -> dict:
    """返回本地 vendored skill 的关键路径。"""
    return {
        "root": str(SKILL_ROOT),
        "skill_md": str(SKILL_ROOT / "SKILL.md"),
        "pptxgenjs_md": str(SKILL_ROOT / "pptxgenjs.md"),
        "editing_md": str(SKILL_ROOT / "editing.md"),
        "thumbnail_py": str(SCRIPTS_ROOT / "thumbnail.py"),
        "soffice_py": str(OFFICE_ROOT / "soffice.py"),
        "unpack_py": str(OFFICE_ROOT / "unpack.py"),
        "pack_py": str(OFFICE_ROOT / "pack.py"),
        "validate_py": str(OFFICE_ROOT / "validate.py"),
    }


def assert_skill_present():
    """确保本地 vendored skill 存在。"""
    required = [
        SKILL_ROOT / "SKILL.md",
        SKILL_ROOT / "pptxgenjs.md",
        SCRIPTS_ROOT / "thumbnail.py",
        OFFICE_ROOT / "soffice.py",
    ]
    missing = [str(p) for p in required if not p.exists()]
    if missing:
        raise FileNotFoundError(
            "本地 Anthropic PPTX skill 不完整，缺少: " + ", ".join(missing)
        )


def run_js(code: str, output_path: str, timeout: int = 60) -> str:
    """
    执行一段 PptxGenJS JavaScript 代码，生成 .pptx 文件。
    使用本地 node + pptxgenjs。
    """
    assert_skill_present()
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".js", delete=False, encoding="utf-8"
    ) as f:
        f.write(code)
        tmp_js = f.name

    try:
        env = os.environ.copy()
        env["NODE_PATH"] = NODE_PATH

        result = subprocess.run(
            ["node", tmp_js],
            capture_output=True, text=True,
            timeout=timeout, env=env,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"PptxGenJS 代码执行失败 (exit {result.returncode}):\n"
                f"stderr: {result.stderr[:800]}\n"
                f"stdout: {result.stdout[:200]}"
            )

        if not os.path.isfile(output_path):
            raise RuntimeError(
                f"代码执行成功但未生成文件: {output_path}\n"
                f"stdout: {result.stdout[:300]}"
            )

        size = os.path.getsize(output_path)
        print(f"[PptxSkill] 生成成功: {output_path} ({size:,} bytes)")
        return output_path

    finally:
        os.unlink(tmp_js)


def read_pptx(path: str) -> str:
    """用 markitdown 提取 .pptx 文本。"""
    try:
        result = subprocess.run(
            ["python", "-m", "markitdown", path],
            capture_output=True, text=True, timeout=30,
        )
        return result.stdout if result.returncode == 0 else ""
    except Exception as e:
        logger.warning(f"[PptxSkill] markitdown 失败: {e}")
        return ""


def pptx_to_images(pptx_path: str, output_dir: str = None) -> list[str]:
    """
    调用本地 vendored skill 的 soffice.py 将 .pptx 转 PDF，再用 pdftoppm 转图片。
    """
    assert_skill_present()

    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(pptx_path), "slides_preview")
    os.makedirs(output_dir, exist_ok=True)

    pdf_path = os.path.join(output_dir, "temp.pdf")

    try:
        # 使用 vendored soffice.py
        result = subprocess.run(
            [
                "python", str(OFFICE_ROOT / "soffice.py"),
                "--headless", "--convert-to", "pdf", pptx_path,
            ],
            cwd=output_dir,
            capture_output=True, text=True, timeout=120,
        )

        # soffice.py 会在 cwd 产出同名 pdf
        base = os.path.splitext(os.path.basename(pptx_path))[0]
        generated_pdf = os.path.join(output_dir, f"{base}.pdf")
        if os.path.isfile(generated_pdf):
            os.rename(generated_pdf, pdf_path)
    except Exception as e:
        logger.warning(f"[PptxSkill] PPTX→PDF 失败: {e}")
        return []

    if not os.path.isfile(pdf_path):
        return []

    try:
        prefix = os.path.join(output_dir, "slide")
        subprocess.run(
            ["pdftoppm", "-jpeg", "-r", "150", pdf_path, prefix],
            capture_output=True, timeout=60,
        )
    except Exception as e:
        logger.warning(f"[PptxSkill] PDF→图片 失败: {e}")
        return []
    finally:
        if os.path.isfile(pdf_path):
            os.unlink(pdf_path)

    images = sorted([
        os.path.join(output_dir, f)
        for f in os.listdir(output_dir)
        if f.startswith("slide") and f.endswith(".jpg")
    ])
    print(f"[PptxSkill] 生成 {len(images)} 张预览图")
    return images
