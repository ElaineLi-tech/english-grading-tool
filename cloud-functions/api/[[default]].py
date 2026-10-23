"""
EdgeOne Makers Cloud Functions 入口
WSGI 模式 - Flask 应用
文件路径: cloud-functions/api/[[default]].py
自动映射为: /api/*
"""

import os
import sys

# 将项目根目录加入 Python 路径
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# 设置环境为云模式
os.environ.setdefault('STORAGE_MODE', 'cloud')

# app_core.py 使用了相对导入（from .templates_data import ...）
# 必须以 lib 包的方式导入
from flask import Flask
from lib.app_core import create_app

# EdgeOne Makers 入口文件识别规则：
# 只有包含 `app = Flask(...)` 等入口标识的 .py 文件才会被注册为路由。
# 因此先创建一个 Flask 实例用于入口识别，再用 create_app() 返回的
# 完整配置应用实例覆盖它。
app = Flask(__name__)
app = create_app()

# EdgeOne Makers Python WSGI 模式
# 当文件导出一个 WSGI 可调用对象（如 Flask app）时，自动识别为 WSGI 模式
# 路由映射：cloud-functions/api/[[default]].py -> /api/*
#
# 注意：EdgeOne 会自动剥离文件系统路由前缀 /api
# 请求 /api/grade 经剥离后 Flask 收到的路径为 /grade
# 因此 app_core.py 中路由定义为 API_PREFIX + '/grade'
# 在 cloud 模式下 API_PREFIX = ''，所以路由实际为 /grade
