import os
from dotenv import load_dotenv

load_dotenv()

GLM_API_KEY = os.getenv("GLM_API_KEY", "")
GLM_BASE_URL = os.getenv("GLM_BASE_URL", "")
PLANNER_MODEL = os.getenv("PLANNER_MODEL", "glm-5")
MAX_TOKENS_PLANNER = 16384

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

# 豆包图片生成
ARK_API_KEY = os.getenv("ARK_API_KEY", "")
ARK_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DOUBAO_IMAGE_MODEL = os.getenv("DOUBAO_IMAGE_MODEL", "doubao-seedream-4-5-251128")
DOUBAO_IMAGE_SIZE = os.getenv("DOUBAO_IMAGE_SIZE", "2K")

# Qwen-VL 视觉评估
QWEN_API_KEY = os.getenv("QWEN_API_KEY", "")
QWEN_BASE_URL = os.getenv("QWEN_BASE_URL", "")
QWEN_VL_MODEL = os.getenv("QWEN_VL_MODEL", "qwen-vl-max")
EVAL_SCORE_THRESHOLD = float(os.getenv("EVAL_SCORE_THRESHOLD", "3.0"))
EVAL_MAX_ROUNDS = int(os.getenv("EVAL_MAX_ROUNDS", "2"))
