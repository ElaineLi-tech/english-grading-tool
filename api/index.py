import sys
import os

# 将项目根目录加入 Python 路径
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

# 直接导入 Flask 应用，Vercel 会自动识别为 WSGI app
from app import app  # noqa: F401
