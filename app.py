"""
英语作业自动批改工具 - Flask Web应用
支持：作文批改、概要写作批改、翻译批改

本地开发入口文件
"""

from lib.app_core import create_app

app = create_app()

if __name__ == '__main__':
    print("=" * 50)
    print("  英语作业自动批改工具")
    print("  访问地址：http://localhost:5005")
    print("=" * 50)
    app.run(host='0.0.0.0', debug=True, port=5005)
