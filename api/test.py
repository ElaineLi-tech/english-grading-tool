def handler(event, context):
    """简单测试函数，验证Vercel Python运行时"""
    import os
    return {
        'statusCode': 200,
        'headers': {'Content-Type': 'text/plain; charset=utf-8'},
        'body': (
            'Vercel Python Function OK!\n'
            f'Python path: {os.getcwd()}\n'
            f'Files: {os.listdir(".")}\n'
        ),
        'isBase64Encoded': False,
    }
