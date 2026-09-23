import sys
import os
import base64
import traceback

# 将项目根目录加入 Python 路径
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

# 调试信息：列出项目根目录的文件
_debug_info = {
    'project_root': PROJECT_ROOT,
    'cwd': os.getcwd(),
    'files': [],
}

try:
    _debug_info['files'] = os.listdir(PROJECT_ROOT)
except Exception as e:
    _debug_info['list_error'] = str(e)

# 尝试导入 Flask 应用
_app = None
_import_error = None

try:
    from app import app as _app
except Exception:
    _import_error = traceback.format_exc()


def handler(event, context):
    """Vercel Python Serverless Function handler"""
    
    # 如果 app 加载失败，返回详细的错误信息
    if _app is None:
        return {
            'statusCode': 500,
            'headers': {
                'Content-Type': 'text/plain; charset=utf-8',
                'X-Debug': 'app-load-failed',
            },
            'body': (
                'App failed to load!\n\n'
                f'Import error:\n{_import_error}\n\n'
                f'Project root: {PROJECT_ROOT}\n'
                f'Current dir: {os.getcwd()}\n'
                f'Python path: {sys.path}\n\n'
                f'Files in project root:\n'
                + '\n'.join(f'  - {f}' for f in _debug_info['files'])
            ),
            'isBase64Encoded': False,
        }
    
    try:
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
        with _app.test_client() as client:
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
        content_type = resp_headers.get('Content-Type', 'text/plain')
        
        # 判断是否为二进制内容
        is_binary = not (
            content_type.startswith('text/') or
            content_type.startswith('application/json') or
            content_type.startswith('application/javascript') or
            content_type == 'image/svg+xml' or
            'xml' in content_type.lower()
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
    
    except Exception:
        error_detail = traceback.format_exc()
        return {
            'statusCode': 500,
            'headers': {
                'Content-Type': 'text/plain; charset=utf-8',
                'X-Debug': 'handler-error',
            },
            'body': f'Handler error:\n{error_detail}',
            'isBase64Encoded': False,
        }
