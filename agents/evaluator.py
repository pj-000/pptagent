"""
agents/evaluator.py

用 qwen-vl 对生成的 PPT 每页做视觉评分。
图片 base64 编码后发给多模态模型，返回 SlideEvalResult 列表。
如果 QWEN_API_KEY 未配置，evaluate_all() 直接返回空列表，不影响主流程。
"""
import base64
import json
import logging
import re
from pathlib import Path

from openai import OpenAI

import config
from models.schemas import OutlinePlan, SlideEvalResult

logger = logging.getLogger(__name__)


class EvaluatorAgent:
    def __init__(self):
        self.enabled = bool(config.QWEN_API_KEY and config.QWEN_BASE_URL)
        if self.enabled:
            self.client = OpenAI(
                api_key=config.QWEN_API_KEY,
                base_url=config.QWEN_BASE_URL,
            )
        else:
            self.client = None
            print("[Evaluator] QWEN_API_KEY 未配置，视觉 QA 已禁用")

    def evaluate_all(
        self,
        image_paths: list[str],
        outline: OutlinePlan,
    ) -> list[SlideEvalResult]:
        """
        逐页视觉评分。image_paths 与 outline.slides 一一对应（长度可能不等，取 min）。
        返回 SlideEvalResult 列表，仅包含成功评分的页。
        """
        if not self.enabled:
            return []

        results = []
        n = min(len(image_paths), len(outline.slides))
        for i in range(n):
            img_path = image_paths[i]
            slide = outline.slides[i]
            if not img_path or not Path(img_path).exists():
                continue
            try:
                result = self._evaluate_slide(img_path, slide.slide_index, slide.topic, slide.layout.value)
                results.append(result)
                print(
                    f"[Evaluator] 第 {slide.slide_index} 页评分: "
                    f"layout={result.layout_score:.1f} content={result.content_score:.1f} "
                    f"design={result.design_score:.1f} overall={result.overall:.1f}"
                )
            except Exception as e:
                logger.warning(f"[Evaluator] 第 {slide.slide_index} 页评分失败: {e}")

        if results:
            avg = sum(r.overall for r in results) / len(results)
            low = [r for r in results if r.overall < config.EVAL_SCORE_THRESHOLD]
            print(f"[Evaluator] 评分完成，平均分 {avg:.2f}，{len(low)} 页低于阈值 {config.EVAL_SCORE_THRESHOLD}")

        return results

    def _evaluate_slide(
        self,
        image_path: str,
        slide_index: int,
        topic: str,
        layout: str,
    ) -> SlideEvalResult:
        """单页评分：base64 编码图片，发给 qwen-vl，解析 JSON 结果。"""
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")

        ext = Path(image_path).suffix.lower().lstrip(".")
        mime = "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"

        system = (
            "你是 PPT 视觉质量评审专家。假设幻灯片有问题，你的任务是找到它们。\n"
            "对幻灯片图片进行评分，只输出 JSON，不要解释。\n"
            "格式：\n"
            "{\n"
            '  "layout_score": 4.0,\n'
            '  "content_score": 3.5,\n'
            '  "design_score": 4.0,\n'
            '  "issues": ["问题1", "问题2"],\n'
            '  "suggestions": ["建议1", "建议2"]\n'
            "}\n\n"
            "评分标准（1-5分）：\n"
            "layout_score 重点检查：\n"
            "- 元素重叠（文字穿过形状、线条穿过文字、元素堆叠）\n"
            "- 文字溢出或被裁切（超出文本框/幻灯片边界）\n"
            "- 装饰线为单行文字设计但标题换行了\n"
            "- 来源引用或页脚与上方内容碰撞\n"
            "- 元素间距过小（< 0.3\"）或卡片/区块几乎贴在一起\n"
            "- 间距不均匀（一处大片空白，另一处拥挤）\n"
            "- 幻灯片边缘边距不足（< 0.5\"）\n"
            "- 列或相似元素未对齐\n"
            "- 文本框过窄导致过度换行\n\n"
            "content_score 重点检查：\n"
            "- 标题是否清晰可读\n"
            "- 要点是否充实（不是只有几个词）\n"
            "- 是否有占位符残留文字\n"
            "- 信息密度是否适当\n\n"
            "design_score 重点检查：\n"
            "- 低对比度文字（如浅灰文字在浅色背景上）\n"
            "- 低对比度图标（如深色图标在深色背景上没有对比圆圈）\n"
            "- 是否有视觉重心（不是纯文字堆砌）\n"
            "- 配色是否协调\n"
            "- 是否有标题下方的装饰横线（AI 生成的典型反模式）\n\n"
            "issues 必须列出所有发现的问题，即使是小问题。suggestions 给出具体修复建议。"
        )

        user_text = (
            f"这是第 {slide_index} 页幻灯片，主题：{topic}，布局类型：{layout}。\n"
            "请对这张幻灯片图片进行视觉质量评分。"
        )

        response = self.client.chat.completions.create(
            model=config.QWEN_VL_MODEL,
            max_tokens=512,
            messages=[
                {"role": "system", "content": system},
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:{mime};base64,{b64}"},
                        },
                        {"type": "text", "text": user_text},
                    ],
                },
            ],
        )

        raw = response.choices[0].message.content
        data = self._parse_json(raw)

        layout_score = float(data.get("layout_score", 3.0))
        content_score = float(data.get("content_score", 3.0))
        design_score = float(data.get("design_score", 3.0))
        overall = layout_score * 0.4 + content_score * 0.3 + design_score * 0.3

        return SlideEvalResult(
            slide_index=slide_index,
            layout_score=layout_score,
            content_score=content_score,
            design_score=design_score,
            overall=round(overall, 2),
            issues=data.get("issues", []),
            suggestions=data.get("suggestions", []),
        )

    @staticmethod
    def _parse_json(raw: str) -> dict:
        cleaned = raw.strip()
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            match = re.search(r"(\{.*\})", cleaned, re.DOTALL)
            if match:
                return json.loads(match.group(1))
            raise
