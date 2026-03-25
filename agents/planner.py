import re
import os
import logging
from pathlib import Path
from openai import OpenAI
from tools.pptx_skill import run_js, skill_paths, assert_skill_present
import config

logger = logging.getLogger(__name__)

MAX_RETRIES = 3

# 本地 vendored skill 路径
_SKILL = skill_paths()


class PlannerAgent:
    """
    读取本地 Anthropic PPTX skill 文档作为 system prompt，
    让 GLM 生成 PptxGenJS 代码，通过 pptx_skill.run_js() 执行。
    """

    def __init__(self):
        assert_skill_present()
        self.client = OpenAI(api_key=config.GLM_API_KEY, base_url=config.GLM_BASE_URL)

        # 直接读取本地 vendored skill 文档
        self._skill_md = Path(_SKILL["skill_md"]).read_text(encoding="utf-8")
        self._pptxgenjs_md = Path(_SKILL["pptxgenjs_md"]).read_text(encoding="utf-8")
        self._user_template = Path("prompts/planner_user.txt").read_text(encoding="utf-8")

    def _build_system_prompt(self) -> str:
        """
        system prompt = 角色指令 + 官方 SKILL.md 的 Design Ideas 部分 + 完整 pptxgenjs.md
        """
        # 从 SKILL.md 中提取 Design Ideas 到末尾的内容
        design_section = self._skill_md
        marker = "## Design Ideas"
        idx = design_section.find(marker)
        if idx >= 0:
            design_section = design_section[idx:]

        return f"""你是一位顶级 PPT 视觉设计工程师。你使用 PptxGenJS（Node.js）生成精美的演示文稿。

以下是你必须严格遵守的设计规范和 API 参考，来自 Anthropic 官方 PPTX Design Skill：

---
{design_section}
---

以下是 PptxGenJS 的完整 API 教程，你生成的代码必须严格遵循这些用法和注意事项：

---
{self._pptxgenjs_md}
---

## 你的输出格式

输出一段完整的 Node.js 代码，用 <code> 标签包裹。代码要求：

1. `const pptxgen = require("pptxgenjs");` 开头
2. 使用 `pres.layout = "LAYOUT_WIDE";`（13.33" × 7.5"）
3. 最后调用 `pres.writeFile({{ fileName: "OUTPUT_PATH" }});`
4. 只使用 pptxgenjs，不要其他 npm 包
5. 严格遵循上面 Design Ideas 中的所有设计规则
6. 严格遵循上面 PptxGenJS Tutorial 中的所有 API 用法和 Common Pitfalls
"""

    def _build_user_prompt(self, topic: str, min_slides: int = 6, max_slides: int = 10,
                           style: str = "auto") -> str:
        return (self._user_template
                .replace("{topic}", topic)
                .replace("{slide_width}", str(config.SLIDE_WIDTH_INCH))
                .replace("{slide_height}", str(config.SLIDE_HEIGHT_INCH))
                .replace("{language}", "中文")
                .replace("{style}", style)
                .replace("{min_slides}", str(min_slides))
                .replace("{max_slides}", str(max_slides)))

    def plan(self, topic: str, output_path: str = None,
             min_slides: int = 6, max_slides: int = 10,
             style: str = "auto") -> str:
        """
        生成 PptxGenJS 代码并执行，直接产出 .pptx 文件。

        Returns:
            生成的 .pptx 文件绝对路径
        """
        if output_path is None:
            os.makedirs(config.OUTPUT_DIR, exist_ok=True)
            output_path = os.path.join(config.OUTPUT_DIR, "output.pptx")

        output_path = os.path.abspath(output_path)
        print(f"[Planner] 开始规划主题：{topic}")

        system_prompt = self._build_system_prompt()
        user_prompt = self._build_user_prompt(topic, min_slides, max_slides, style=style)
        errors_so_far: list[str] = []
        last_raw = ""

        for attempt in range(1, MAX_RETRIES + 1):
            print(f"[Planner] 第 {attempt}/{MAX_RETRIES} 次尝试...")

            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ]

            if errors_so_far:
                error_feedback = (
                    "你上一次输出的代码执行失败，请修正后重新输出完整的 <code>...</code>：\n"
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

            try:
                code = self._extract_code(last_raw)
                print(f"[Planner] 提取到代码（{len(code)} 字符）")

                full_code = self._inject_output_path(code, output_path)
                run_js(full_code, output_path)
                print(f"[Planner] 生成完成: {output_path}")
                return output_path

            except Exception as e:
                err_msg = str(e)
                logger.warning(f"[Planner] 第 {attempt} 次失败: {err_msg[:200]}")
                print(f"[Planner] 第 {attempt} 次失败: {err_msg[:200]}")
                errors_so_far = [err_msg[:500]]

        raise RuntimeError(
            f"连续 {MAX_RETRIES} 次生成失败。\n"
            f"最后错误: {errors_so_far}\n"
            f"最后响应（前500字）: {last_raw[:500]}"
        )

    def _extract_code(self, raw: str) -> str:
        """从 LLM 响应中提取 <code>...</code> 或 ```javascript...``` 中的代码。"""
        m = re.search(r"<code>(.*?)</code>", raw, re.DOTALL)
        if m:
            return m.group(1).strip()

        m = re.search(r"```(?:javascript|js)\s*(.*?)```", raw, re.DOTALL)
        if m:
            return m.group(1).strip()

        m = re.search(r"```\s*(.*?)```", raw, re.DOTALL)
        if m:
            return m.group(1).strip()

        raise ValueError("LLM 响应中未找到 <code> 或 ```javascript 代码块")

    def _inject_output_path(self, code: str, output_path: str) -> str:
        """确保代码中的 writeFile 使用正确的 output_path。"""
        safe_path = output_path.replace("\\", "/")

        if "writeFile" in code or "writeToFile" in code:
            code = re.sub(
                r'fileName\s*:\s*["\'][^"\']*["\']',
                f'fileName: "{safe_path}"',
                code
            )
            return code

        code += f'\npres.writeFile({{ fileName: "{safe_path}" }});\n'
        return code
