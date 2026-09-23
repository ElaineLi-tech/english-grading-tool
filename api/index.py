import sys
import os
import base64
from io import BytesIO

# 将项目根目录加入 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app


def handler(event, context):
    """Vercel Python Serverless Function 入口
    将 API Gateway 格式的 event 转换为 WSGI environ，调用 Flask，再转换回 API Gateway 响应
    """
    # 解析 event 中的 HTTP 信息
    http_method = event.get('httpMethod', event.get('method', 'GET'))
    path = event.get('path', '/')
    query_string_params = event.get('queryStringParameters', {}) or {}
    headers = event.get('headers', {}) or {}
    body = event.get('body', '')

    # 处理 base64 编码的 body
    if event.get('isBase64Encoded', False) and body:
        body = base64.b64decode(body)
    else:
        body = body.encode('utf-8') if body else b''

    # 构建 query string
    query_string = '&'.join([f'{k}={v}' for k, v in query_string_params.items()])

    # 构建 WSGI environ
    environ = {
        'REQUEST_METHOD': http_method,
        'SCRIPT_NAME': '',
        'PATH_INFO': path,
        'QUERY_STRING': query_string,
        'CONTENT_TYPE': headers.get('content-type', headers.get('Content-Type', '')),
        'CONTENT_LENGTH': str(len(body)),
        'SERVER_NAME': 'vercel',
        'SERVER_PORT': '443',
        'SERVER_PROTOCOL': 'HTTP/1.1',
        'wsgi.version': (1, 0),
        'wsgi.url_scheme': 'https',
        'wsgi.input': BytesIO(body),
        'wsgi.errors': sys.stderr,
        'wsgi.multithread': False,
        'wsgi.multiprocess': False,
        'wsgi.run_once': False,
    }

    # 添加 HTTP 头
    for key, value in headers.items():
        key_upper = key.upper().replace('-', '_')
        if key_upper == 'CONTENT_TYPE':
            continue
        if key_upper == 'CONTENT_LENGTH':
            continue
        environ[f'HTTP_{key_upper}'] = value

    # 调用 Flask 应用
    status = '200 OK'
    response_headers = []

    def start_response(s, h):
        nonlocal status, response_headers
        status = s
        response_headers = h

    body_iter = app.wsgi_app(environ, start_response)
    body_bytes = b''.join(body_iter)

    # 转换为 API Gateway 响应格式
    status_code = int(status.split()[0])

    headers_dict = {}
    is_binary = True
    for key, value in response_headers:
        headers_dict[key] = value
        if key.lower() == 'content-type':
            ct = value.lower()
            if ct.startswith('text/') or ct in ('application/json', 'application/javascript', 'application/xml', 'image/svg+xml'):
                is_binary = False

    if is_binary:
        return {
            'statusCode': status_code,
            'headers': headers_dict,
            'body': base64.b64encode(body_bytes).decode('utf-8'),
            'isBase64Encoded': True,
        }
    else:
        return {
            'statusCode': status_code,
            'headers': headers_dict,
            'body': body_bytes.decode('utf-8', errors='replace'),
            'isBase64Encoded': False,
        }
