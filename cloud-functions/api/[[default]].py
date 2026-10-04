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

# 将 lib 目录加入 Python 路径
LIB_DIR = os.path.join(PROJECT_ROOT, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# 设置环境为云模式
os.environ.setdefault('STORAGE_MODE', 'cloud')

from app_core import create_app

# 创建 Flask 应用
app = create_app()

# EdgeOne Makers Python WSGI 模式
# 当文件导出一个 WSGI 可调用对象（如 Flask app）时，自动识别为 WSGI 模式
# 路由映射：cloud-functions/api/[[default]].py -> /api/*
# Flask 内部路由定义为 /api/xxx，因此需要做路径调整

# 使用 werkzeug DispatcherMiddleware 处理路径前缀
# Makers 中请求路径是 /api/xxx，Flask 内部路由也是 /api/xxx
# 所以直接暴露 app 即可（WSGI 环境中的 PATH_INFO 已经是去掉 /api 前缀的？还是完整的？）
# 根据 Makers 文档：WSGI 模式下，传入的 environ 中的 PATH_INFO 是完整的请求路径
# 因此我们的 Flask app 路由定义为 /api/xxx 是正确的

# 直接导出 app 作为 WSGI 应用
# EdgeOne Makers 会自动检测到这是一个 WSGI app 并使用 WSGI 模式
