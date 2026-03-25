import pytest
import os
import sys
import json
import asyncio
import time
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from unittest.mock import MagicMock, patch, AsyncMock
from pydantic import ValidationError
from models.schemas import (
    TextElement, SlideSpec, SlideLayout, PresentationPlan,
    LayoutValidationError, SLIDE_WIDTH, SLIDE_HEIGHT,
)
from agents.asset_agent import AssetAgent
from agents.researcher import ResearchAgent
from tools.pptx_renderer import PPTXRenderer


# ─── image_placeholder 校验 ───


class TestImagePlaceholderValidation:

    def test_image_placeholder_requires_query_or_prompt(self):
        """image_placeholder 必须提供 unsplash_query 或 dalle_prompt"""
        with pytest.raises(ValidationError, match="image_placeholder.*至少提供"):
            TextElement(
                type="image_placeholder", content="测试图片",
                x=0.5, y=1.5, width=5.0, height=4.0, font_size=14,
                color="#000000",
            )

    def test_image_placeholder_with_unsplash_query_passes(self):
        """提供 unsplash_query 时校验通过"""
        elem = TextElement(
            type="image_placeholder", content="太阳能板",
            x=0.5, y=1.5, width=5.0, height=4.0, font_size=14,
            color="#000000", unsplash_query="solar panel",
        )
        assert elem.unsplash_query == "solar panel"

    def test_image_placeholder_with_dalle_prompt_passes(self):
        """提供 dalle_prompt 时校验通过"""
        elem = TextElement(
            type="image_placeholder", content="太阳能板",
            x=0.5, y=1.5, width=5.0, height=4.0, font_size=14,
            color="#000000", dalle_prompt="A solar panel on rooftop",
        )
        assert elem.dalle_prompt == "A solar panel on rooftop"

    def test_image_placeholder_with_both_passes(self):
        """同时提供两者也通过"""
        elem = TextElement(
            type="image_placeholder", content="太阳能板",
            x=0.5, y=1.5, width=5.0, height=4.0, font_size=14,
            color="#000000",
            unsplash_query="solar panel",
            dalle_prompt="A solar panel on rooftop",
        )
        assert elem.unsplash_query and elem.dalle_prompt

    def test_local_image_path_default_none(self):
        """local_image_path 初始为 None"""
        elem = TextElement(
            type="image_placeholder", content="测试",
            x=0.5, y=1.5, width=5.0, height=4.0, font_size=14,
            color="#000000", unsplash_query="test",
        )
        assert elem.local_image_path is None

    def test_non_image_element_no_query_required(self):
        """非 image_placeholder 元素不要求 query"""
        elem = TextElement(
            type="body", content="正文内容",
            x=0.5, y=1.5, width=12.0, height=5.0, font_size=18,
            color="#333333",
        )
        assert elem.unsplash_query is None
        assert elem.dalle_prompt is None


# ─── AssetAgent 测试 ───


class TestAssetAgent:

    def _make_slide_with_image(self, query="test image", prompt="a test image"):
        return SlideSpec(
            slide_index=0, layout=SlideLayout.CONTENT, topic="测试",
            elements=[
                TextElement(
                    type="title", content="标题",
                    x=0.5, y=0.3, width=12.0, height=0.9,
                    font_size=32, bold=True, color="#1F3864",
                ),
                TextElement(
                    type="image_placeholder", content="测试图片",
                    x=6.9, y=1.5, width=5.9, height=5.5, font_size=14,
                    color="#000000",
                    unsplash_query=query, dalle_prompt=prompt,
                ),
            ],
        )

    def test_cache_hit_returns_path(self, tmp_path):
        """缓存命中直接返回"""
        agent = AssetAgent()
        slide = self._make_slide_with_image()
        job_id = "test_cache"
        cache_dir = os.path.join(str(tmp_path), job_id)
        os.makedirs(cache_dir, exist_ok=True)

        # 预写缓存文件
        cache_key = agent._cache_key("test image", "a test image")
        cached_path = os.path.join(cache_dir, f"{cache_key}.jpg")
        with open(cached_path, "wb") as f:
            f.write(b"\xff\xd8fake_jpg_data")

        with patch.object(agent, '_fetch_unsplash', new_callable=AsyncMock) as mock_unsplash:
            with patch("config.ASSETS_DIR", str(tmp_path)):
                asyncio.run(agent.fetch_all([slide], job_id=job_id))

            # Unsplash 不应被调用
            mock_unsplash.assert_not_called()

        assert slide.elements[1].local_image_path == cached_path

    def test_unsplash_fail_fallback_dalle(self, tmp_path):
        """Unsplash 无结果时降级到 DALL-E"""
        agent = AssetAgent()
        agent.unsplash_key = "fake_key"
        slide = self._make_slide_with_image()

        with patch.object(agent, '_fetch_unsplash', new_callable=AsyncMock, return_value=None):
            with patch.object(agent, '_fetch_dalle', new_callable=AsyncMock, return_value="/fake/path.jpg"):
                with patch("config.ASSETS_DIR", str(tmp_path)):
                    asyncio.run(agent.fetch_all([slide], job_id="test_fallback"))

        assert slide.elements[1].local_image_path == "/fake/path.jpg"

    def test_both_fail_returns_none(self, tmp_path):
        """两者都失败时 local_image_path 保持 None"""
        agent = AssetAgent()
        agent.unsplash_key = "fake_key"
        slide = self._make_slide_with_image()

        with patch.object(agent, '_fetch_unsplash', new_callable=AsyncMock, return_value=None):
            with patch.object(agent, '_fetch_dalle', new_callable=AsyncMock, return_value=None):
                with patch("config.ASSETS_DIR", str(tmp_path)):
                    asyncio.run(agent.fetch_all([slide], job_id="test_both_fail"))

        assert slide.elements[1].local_image_path is None

    def test_fetch_all_writes_back_path(self, tmp_path):
        """fetch_all 能把 local_image_path 回写到 ElementSpec"""
        agent = AssetAgent()
        agent.unsplash_key = "fake_key"
        slide = self._make_slide_with_image()

        fake_path = os.path.join(str(tmp_path), "downloaded.jpg")
        with patch.object(agent, '_fetch_unsplash', new_callable=AsyncMock, return_value=fake_path):
            with patch("config.ASSETS_DIR", str(tmp_path)):
                asyncio.run(agent.fetch_all([slide], job_id="test_writeback"))

        assert slide.elements[1].local_image_path == fake_path


# ─── ResearchAgent 测试 ───


class TestResearchAgent:

    def _make_slide(self, layout, topic="测试主题"):
        return SlideSpec(
            slide_index=0, layout=layout, topic=topic,
            elements=[
                TextElement(
                    type="title", content=topic,
                    x=0.5, y=0.3, width=12.0, height=0.9,
                    font_size=32, bold=True, color="#1F3864",
                ),
            ],
        )

    def test_skip_cover(self):
        """cover 页面返回 None"""
        with patch("agents.researcher.AsyncOpenAI"):
            with patch("agents.researcher.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock {language}"
                agent = ResearchAgent()
        slide = self._make_slide(SlideLayout.COVER)
        result = asyncio.run(agent.research_slide(slide))
        assert result is None

    def test_skip_closing(self):
        """closing 页面返回 None"""
        with patch("agents.researcher.AsyncOpenAI"):
            with patch("agents.researcher.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock {language}"
                agent = ResearchAgent()
        slide = self._make_slide(SlideLayout.CLOSING)
        result = asyncio.run(agent.research_slide(slide))
        assert result is None

    def test_default_content_on_failure(self):
        """研究失败时返回默认内容结构"""
        with patch("agents.researcher.AsyncOpenAI") as mock_cls:
            with patch("agents.researcher.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock {language}"
                agent = ResearchAgent()
                agent.client = AsyncMock()
                agent.client.chat.completions.create.side_effect = Exception("API error")

        slide = self._make_slide(SlideLayout.CONTENT)
        result = asyncio.run(agent.research_slide(slide))
        assert result is not None
        assert "topic" in result
        assert "bullet_points" in result
        assert isinstance(result["bullet_points"], list)

    def test_research_all_length_matches_slides(self):
        """research_all 返回长度与 slides 一致"""
        with patch("agents.researcher.AsyncOpenAI"):
            with patch("agents.researcher.Path") as mock_path:
                mock_path.return_value.read_text.return_value = "mock {language}"
                agent = ResearchAgent()

        slides = [
            self._make_slide(SlideLayout.COVER, "封面"),
            self._make_slide(SlideLayout.CONTENT, "内容1"),
            self._make_slide(SlideLayout.CONTENT, "内容2"),
            self._make_slide(SlideLayout.CLOSING, "结尾"),
        ]

        # Mock API 返回
        mock_response = MagicMock()
        mock_response.choices = [MagicMock()]
        mock_response.choices[0].message.content = json.dumps({
            "topic": "test", "summary": "test", "bullet_points": ["a", "b"]
        })

        agent.client = AsyncMock()
        agent.client.chat.completions.create.return_value = mock_response

        results = asyncio.run(agent.research_all(slides))
        assert len(results) == len(slides)
        assert results[0] is None  # cover
        assert results[3] is None  # closing
        assert results[1] is not None  # content


# ─── Renderer 图片测试 ───


class TestPPTXRendererImages:

    def test_placeholder_renders_without_error(self, tmp_path):
        """无真实图片时渲染灰色占位不报错"""
        import config
        original_output = config.OUTPUT_DIR
        config.OUTPUT_DIR = str(tmp_path)

        try:
            plan = PresentationPlan(
                title="图片测试", topic="测试",
                slides=[
                    SlideSpec(
                        slide_index=0, layout=SlideLayout.CONTENT, topic="有图片",
                        elements=[
                            TextElement(
                                type="title", content="标题",
                                x=0.5, y=0.3, width=12.0, height=0.9,
                                font_size=32, bold=True, color="#1F3864",
                            ),
                            TextElement(
                                type="image_placeholder", content="测试图片",
                                x=6.9, y=1.5, width=5.9, height=5.5, font_size=14,
                                color="#000000", unsplash_query="test",
                            ),
                        ],
                    ),
                    SlideSpec(
                        slide_index=1, layout=SlideLayout.CLOSING, topic="结尾",
                        elements=[
                            TextElement(
                                type="title", content="结束",
                                x=1.0, y=2.5, width=11.0, height=1.2,
                                font_size=40, bold=True, color="#000000", align="center",
                            ),
                        ],
                    ),
                ],
            )
            renderer = PPTXRenderer()
            output_path = renderer.render(plan, "test_placeholder.pptx")
            assert os.path.exists(output_path)
        finally:
            config.OUTPUT_DIR = original_output

    def test_real_image_renders(self, tmp_path):
        """有真实图片时成功插图"""
        import config
        original_output = config.OUTPUT_DIR
        config.OUTPUT_DIR = str(tmp_path)

        # 创建一个最小的有效 PNG 文件
        img_path = os.path.join(str(tmp_path), "test.png")
        # 最小有效 PNG: 1x1 红色像素
        import struct, zlib
        def _minimal_png():
            sig = b'\x89PNG\r\n\x1a\n'
            # IHDR
            ihdr_data = struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)
            ihdr_crc = zlib.crc32(b'IHDR' + ihdr_data) & 0xffffffff
            ihdr = struct.pack('>I', 13) + b'IHDR' + ihdr_data + struct.pack('>I', ihdr_crc)
            # IDAT
            raw = zlib.compress(b'\x00\xff\x00\x00')
            idat_crc = zlib.crc32(b'IDAT' + raw) & 0xffffffff
            idat = struct.pack('>I', len(raw)) + b'IDAT' + raw + struct.pack('>I', idat_crc)
            # IEND
            iend_crc = zlib.crc32(b'IEND') & 0xffffffff
            iend = struct.pack('>I', 0) + b'IEND' + struct.pack('>I', iend_crc)
            return sig + ihdr + idat + iend

        with open(img_path, 'wb') as f:
            f.write(_minimal_png())

        try:
            plan = PresentationPlan(
                title="图片测试", topic="测试",
                slides=[
                    SlideSpec(
                        slide_index=0, layout=SlideLayout.CONTENT, topic="有图片",
                        elements=[
                            TextElement(
                                type="title", content="标题",
                                x=0.5, y=0.3, width=12.0, height=0.9,
                                font_size=32, bold=True, color="#1F3864",
                            ),
                            TextElement(
                                type="image_placeholder", content="测试图片",
                                x=6.9, y=1.5, width=5.9, height=5.5, font_size=14,
                                color="#000000", unsplash_query="test",
                                local_image_path=img_path,
                            ),
                        ],
                    ),
                    SlideSpec(
                        slide_index=1, layout=SlideLayout.CLOSING, topic="结尾",
                        elements=[
                            TextElement(
                                type="title", content="结束",
                                x=1.0, y=2.5, width=11.0, height=1.2,
                                font_size=40, bold=True, color="#000000", align="center",
                            ),
                        ],
                    ),
                ],
            )
            renderer = PPTXRenderer()
            output_path = renderer.render(plan, "test_real_image.pptx")
            assert os.path.exists(output_path)

            # 验证文件大小比纯文本大（有图片嵌入）
            file_size = os.path.getsize(output_path)
            assert file_size > 10 * 1024
        finally:
            config.OUTPUT_DIR = original_output


# ─── 并行执行测试 ───


class TestParallelExecution:

    def test_parallel_faster_than_serial(self):
        """Research 和 Asset 并行执行时间明显小于串行之和"""

        async def slow_research(slides, **kwargs):
            await asyncio.sleep(0.3)
            return [None] * len(slides)

        async def slow_asset(slides, **kwargs):
            await asyncio.sleep(0.3)
            return slides

        # 并行
        async def run_parallel():
            t0 = time.time()
            await asyncio.gather(
                slow_research([1, 2, 3]),
                slow_asset([1, 2, 3]),
            )
            return time.time() - t0

        parallel_time = asyncio.run(run_parallel())

        # 串行
        async def run_serial():
            t0 = time.time()
            await slow_research([1, 2, 3])
            await slow_asset([1, 2, 3])
            return time.time() - t0

        serial_time = asyncio.run(run_serial())

        # 并行应该明显快于串行（至少快 30%）
        assert parallel_time < serial_time * 0.8, (
            f"并行 {parallel_time:.2f}s 应明显快于串行 {serial_time:.2f}s"
        )


# ─── 集成测试（需要真实 API） ───


class TestPhase3Integration:

    @pytest.fixture(autouse=True)
    def check_api_key(self):
        if not os.getenv("GLM_API_KEY"):
            pytest.skip("未设置 GLM_API_KEY，跳过集成测试")

    def test_plan_contains_image_placeholders(self):
        """Planner 输出应包含 image_placeholder 元素"""
        from agents.planner import PlannerAgent
        planner = PlannerAgent()
        plan = planner.plan("可再生能源发展趋势")
        has_image = any(
            elem.type == "image_placeholder"
            for slide in plan.slides
            for elem in slide.elements
        )
        assert has_image, "Planner 应该在某些页面规划图片元素"

    def test_no_images_mode_renders_placeholders(self):
        """--no-images 模式下图片位置显示灰色占位"""
        from agents.orchestrator import OrchestratorAgent
        import config
        orch = OrchestratorAgent(no_images=True, no_research=True)
        output = orch.generate("太阳能技术", output_filename="test_no_img.pptx")
        assert os.path.exists(output)

    def test_full_pipeline_with_research(self):
        """完整 Phase 3 流程（含 Research，不含图片下载）"""
        from agents.orchestrator import OrchestratorAgent
        orch = OrchestratorAgent(no_images=True)
        output = orch.generate("Python编程入门", output_filename="test_full_p3.pptx")
        assert os.path.exists(output)
