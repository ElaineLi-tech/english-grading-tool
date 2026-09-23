import sys
import os
import base64

# 将项目根目录加入 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app


def handler(event, context):
    """Vercel Python Serverless Function - 使用 Flask test client 处理请求"""
    method = event.get('httpMethod', 'GET')
    path = event.get('path', '/')
    query_params = event.get('queryStringParameters', {}) or {}
    headers = event.get('headers', {}) or {}
    body = event.get('body', '')
    is_base64 = event.get('isBase64Encoded', False)

    # 解码 body
    if is_base64 and body:
        data = base64.b64decode(body)
    else:
        data = body.encode('utf-8') if body else None

    # 使用 Flask test client 发起请求
    with app.test_client() as client:
        resp = client.open(
            path,
            method=method,
            query_string=query_params,
            headers=headers,
            data=data,
            content_type=headers.get('content-type', ''),
        )

    # 构建响应
    resp_headers = dict(resp.headers)
    content_type = resp_headers.get('Content-Type', '')

    # 判断是否为二进制内容
    is_binary = not (
        content_type.startswith('text/') or
        content_type == 'application/json' or
        content_type == 'application/javascript' or
        content_type == 'image/svg+xml' or
        'xml' in content_type
    )

    body_bytes = resp.get_data()

    if is_binary:
        return {
            'statusCode': resp.status_code,
            'headers': resp_headers,
            'body': base64.b64encode(body_bytes).decode('utf-8'),
            'isBase64Encoded': True,
        }
    else:
        return {
            'statusCode': resp.status_code,
            'headers': resp_headers,
            'body': body_bytes.decode('utf-8', errors='replace'),
            'isBase64Encoded': False,
        }
