import os
import hashlib
import asyncio
import logging
import httpx
from models.schemas import SlideSpec, TextElement
import config

logger = logging.getLogger(__name__)


class AssetAgent:
    """为 image_placeholder 元素获取真实图片。优先 Unsplash，降级 DALL-E。"""

    def __init__(self):
        self.unsplash_key = config.UNSPLASH_ACCESS_KEY
        self.glm_api_key = config.GLM_API_KEY
        self.glm_base_url = config.GLM_BASE_URL

    async def fetch_all(
        self, slides: list[SlideSpec], job_id: str, concurrency: int = 3
    ) -> list[SlideSpec]:
        """
        扫描所有 image_placeholder 元素，下载图片并回写 local_image_path。
        返回修改后的 slides（原地修改）。
        """
        cache_dir = os.path.join(config.ASSETS_DIR, job_id)
        os.makedirs(cache_dir, exist_ok=True)

        sem = asyncio.Semaphore(concurrency)
        tasks = []

        for slide in slides:
            for elem in slide.elements:
                if elem.type == "image_placeholder":
                    tasks.append(self._fetch_one(elem, cache_dir, sem))

        if tasks:
            await asyncio.gather(*tasks)

        return slides

    async def _fetch_one(self, elem: TextElement, cache_dir: str, sem: asyncio.Semaphore):
        """获取单张图片，失败时 local_image_path 保持 None。"""
        async with sem:
            # 生成缓存 key
            cache_key = self._cache_key(elem.unsplash_query, elem.dalle_prompt)
            cached_path = os.path.join(cache_dir, f"{cache_key}.jpg")

            # 1. 缓存命中
            if os.path.exists(cached_path) and os.path.getsize(cached_path) > 0:
                elem.local_image_path = cached_path
                print(f"[Asset] 缓存命中: {cached_path}")
                return

            # 2. Unsplash
            if elem.unsplash_query and self.unsplash_key:
                path = await self._fetch_unsplash(elem.unsplash_query, cached_path)
                if path:
                    elem.local_image_path = path
                    return

            # 3. DALL-E 降级
            if elem.dalle_prompt and self.glm_api_key:
                path = await self._fetch_dalle(elem.dalle_prompt, cached_path)
                if path:
                    elem.local_image_path = path
                    return

            # 两者都失败
            logger.warning(f"[Asset] 图片获取失败: query={elem.unsplash_query}")
            print(f"[Asset] 图片获取失败: {elem.unsplash_query or elem.dalle_prompt}")

    async def _fetch_unsplash(self, query: str, save_path: str) -> str | None:
        """从 Unsplash 搜索并下载第一张图片。"""
        try:
            async with httpx.AsyncClient(timeout=15) as client:
                resp = await client.get(
                    f"{config.UNSPLASH_BASE_URL}/search/photos",
                    params={"query": query, "per_page": 1, "orientation": "landscape"},
                    headers={"Authorization": f"Client-ID {self.unsplash_key}"},
                )
                if resp.status_code != 200:
                    logger.warning(f"[Asset] Unsplash API 返回 {resp.status_code}")
                    return None

                data = resp.json()
                results = data.get("results", [])
                if not results:
                    logger.info(f"[Asset] Unsplash 无结果: {query}")
                    return None

                image_url = results[0]["urls"]["regular"]
                img_resp = await client.get(image_url, timeout=30)
                if img_resp.status_code == 200:
                    with open(save_path, "wb") as f:
                        f.write(img_resp.content)
                    print(f"[Asset] Unsplash 下载成功: {query} -> {save_path}")
                    return save_path
        except Exception as e:
            logger.warning(f"[Asset] Unsplash 失败: {e}")
        return None

    async def _fetch_dalle(self, prompt: str, save_path: str) -> str | None:
        """通过 OpenAI 兼容接口调用图片生成。"""
        try:
            from openai import AsyncOpenAI

            client = AsyncOpenAI(api_key=self.glm_api_key, base_url=self.glm_base_url)
            response = await client.images.generate(
                model=config.DALLE_MODEL,
                prompt=prompt,
                size=config.DALLE_IMAGE_SIZE,
                quality=config.DALLE_IMAGE_QUALITY,
                n=1,
            )
            image_url = response.data[0].url
            if not image_url:
                return None

            async with httpx.AsyncClient(timeout=30) as http_client:
                img_resp = await http_client.get(image_url)
                if img_resp.status_code == 200:
                    with open(save_path, "wb") as f:
                        f.write(img_resp.content)
                    print(f"[Asset] DALL-E 生成成功: {prompt[:40]}... -> {save_path}")
                    return save_path
        except Exception as e:
            logger.warning(f"[Asset] DALL-E 失败: {e}")
        return None

    @staticmethod
    def _cache_key(query: str | None, prompt: str | None) -> str:
        """基于 query/prompt 生成稳定的缓存文件名。"""
        raw = f"{query or ''}|{prompt or ''}"
        return hashlib.md5(raw.encode()).hexdigest()[:12]
