import sys
import os
import io

# 将项目根目录加入 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app
from werkzeug.test import EnvironBuilder
from werkzeug.wrappers import Response


def handler(request):
    """Vercel Python Serverless Function 入口
    将 Vercel 的 request 对象转换为 Flask WSGI environ，调用 Flask 应用，再转换回 Vercel response
    """
    # 构建 WSGI environ
    builder = EnvironBuilder(
        method=request.method,
        path=request.path,
        query_string=request.query,
        headers=dict(request.headers),
        data=request.body if request.body else None,
    )
    environ = builder.get_environ()

    # 调用 Flask 应用
    status = '200 OK'
    response_headers = []

    def start_response(s, h):
        nonlocal status, response_headers
        status = s
        response_headers = h

    body_iter = app.wsgi_app(environ, start_response)
    body_bytes = b''.join(body_iter)

    # 转换 headers
    headers_dict = {}
    for key, value in response_headers:
        headers_dict[key] = value

    status_code = int(status.split()[0])

    # 判断是否为二进制内容
    content_type = headers_dict.get('Content-Type', '')
    is_binary = not content_type.startswith('text/') and 'json' not in content_type and 'javascript' not in content_type and 'xml' not in content_type

    import base64
    if is_binary:
        return {
            'statusCode': status_code,
            'headers': headers_dict,
            'body': base64.b64encode(body_bytes).decode('utf-8'),
            'encoding': 'base64',
        }
    else:
        return {
            'statusCode': status_code,
            'headers': headers_dict,
            'body': body_bytes.decode('utf-8', errors='replace'),
        }
