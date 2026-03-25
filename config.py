import os
from dotenv import load_dotenv

load_dotenv()

GLM_API_KEY = os.getenv("GLM_API_KEY", "")
GLM_BASE_URL = os.getenv("GLM_BASE_URL", "")
PLANNER_MODEL = os.getenv("PLANNER_MODEL", "glm-5")
MAX_TOKENS_PLANNER = 4096

# Research Agent (Tavily)
TAVILY_API_KEY = os.getenv("TAVILY_API_KEY", "")
TAVILY_BASE_URL = "https://api.tavily.com"
RESEARCH_MODEL = os.getenv("RESEARCH_MODEL", PLANNER_MODEL)
MAX_TOKENS_RESEARCHER = 2048

# 幻灯片尺寸（英寸，16:9）
SLIDE_WIDTH_INCH = 13.333
SLIDE_HEIGHT_INCH = 7.5

OUTPUT_DIR = "outputs"
ASSETS_DIR = "assets"

# Unsplash
UNSPLASH_ACCESS_KEY = os.getenv("UNSPLASH_ACCESS_KEY", "")
UNSPLASH_BASE_URL = "https://api.unsplash.com"

# DALL-E (通过 OpenAI 兼容接口)
DALLE_MODEL = os.getenv("DALLE_MODEL", "dall-e-3")
DALLE_IMAGE_SIZE = os.getenv("DALLE_IMAGE_SIZE", "1024x1024")
DALLE_IMAGE_QUALITY = os.getenv("DALLE_IMAGE_QUALITY", "standard")
