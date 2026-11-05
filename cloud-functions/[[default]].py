"""
EdgeOne Makers Cloud Functions 入口（catch-all）
WSGI 模式 - Flask 应用
文件路径: cloud-functions/[[default]].py
自动映射为: /* （处理所有路由，包括 / 和 /api/*）
"""

import os
import sys

# 将项目根目录加入 Python 路径
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# 设置环境为云模式
os.environ.setdefault('STORAGE_MODE', 'cloud')

from flask import Flask
from lib.app_core import create_app

# EdgeOne Makers 入口识别：需要 Flask 实例标识
app = Flask(__name__)
app = create_app()

# 此云函数处理所有路由（/*）：
# - GET /          -> Flask 返回 index.html（首页）
# - /api/*         -> Flask API 路由（API_PREFIX = '/api'）
# - 静态资源等      -> 由 Flask 或 EdgeOne 静态托管处理
