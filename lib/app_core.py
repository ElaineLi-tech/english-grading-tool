"""
Flask 应用核心模块
包含：应用创建、路由注册、业务逻辑
"""

import os
import uuid
import time
import json
import base64
import datetime

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

from .templates_data import BUILTIN_TEMPLATES, BUILTIN_GRADING_STANDARDS
from .storage import get_storage
from .ai_service import (
    call_doubao_chat,
    extract_text_from_docx,
    extract_text_from_file,
    get_template_file_path,
    grade_with_doubao,
    grade_objective_with_doubao,
    grade_objective_text_with_doubao,
    generate_summary_handout_with_doubao,
    extract_text_with_doubao,
)
from .doc_generator import (
    generate_essay_doc,
    generate_objective_doc,
    generate_summary_handout_doc,
)


# ==================== 应用创建 ====================

def create_app():
    """创建并配置 Flask 应用"""
    # 项目根目录
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'))

    # Vercel 环境检测
    IS_VERCEL = os.environ.get('VERCEL', '') == '1' or os.environ.get('VERCEL_ENV') is not None

    if IS_VERCEL:
        app.config['UPLOAD_FOLDER'] = '/tmp/uploads'
        app.config['OUTPUT_FOLDER'] = '/tmp/outputs'
        app.config['TEMPLATE_FOLDER'] = '/tmp/templates_user'
        app.config['STATIC_TEMPLATE_FOLDER'] = os.path.join(BASE_DIR, 'templates_user')
        app.config['BUILTIN_TEMPLATE_FILES_DIR'] = os.path.join(BASE_DIR, 'builtin_templates')
    else:
        app.config['UPLOAD_FOLDER'] = 'uploads'
        app.config['OUTPUT_FOLDER'] = 'outputs'
        app.config['TEMPLATE_FOLDER'] = 'templates_user'
        app.config['STATIC_TEMPLATE_FOLDER'] = 'templates_user'
        app.config['BUILTIN_TEMPLATE_FILES_DIR'] = os.path.join(BASE_DIR, 'builtin_templates')

    app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    os.makedirs(app.config['TEMPLATE_FOLDER'], exist_ok=True)

    # Vercel 环境：将静态模板目录的文件复制到 /tmp 可写目录
    if IS_VERCEL:
        import shutil
        static_dir = app.config['STATIC_TEMPLATE_FOLDER']
        temp_dir = app.config['TEMPLATE_FOLDER']
        if os.path.exists(static_dir):
            for fname in os.listdir(static_dir):
                src = os.path.join(static_dir, fname)
                dst = os.path.join(temp_dir, fname)
                if os.path.isfile(src) and not os.path.exists(dst):
                    shutil.copy2(src, dst)

    # 初始化 storage
    storage = get_storage(mode='local', base_dir=app.config['TEMPLATE_FOLDER'])
    app.storage = storage

    # ==================== 辅助函数（使用 storage 抽象层） ====================

    def load_builtin_overrides():
        """加载用户对内置模板的覆盖配置"""
        data = storage.get_json('builtin_overrides')
        return data if data is not None else {}

    def save_builtin_overrides(overrides):
        """保存用户对内置模板的覆盖配置"""
        storage.put_json('builtin_overrides', overrides)

    def get_effective_template(grade_type):
        """获取实际生效的模板（优先使用用户覆盖版本）"""
        builtin = BUILTIN_TEMPLATES.get(grade_type)
        if not builtin:
            return None
        overrides = load_builtin_overrides()
        if grade_type in overrides:
            effective = builtin.copy()
            effective.update(overrides[grade_type])
            effective['is_overridden'] = True
            return effective
        result = builtin.copy()
        result['is_overridden'] = False
        return result

    def get_builtin_template_file_path(template_file):
        """获取内置模板文件的完整路径"""
        from .ai_service import get_builtin_template_file_path as _get_builtin_path
        return _get_builtin_path(template_file, app.config.get('BUILTIN_TEMPLATE_FILES_DIR', ''))

    def get_template_file_path_local(template_file, is_overridden=False):
        """获取模板文件的完整路径（优先用户覆盖版本，其次内置版本）"""
        return get_template_file_path(
            template_file,
            app.config['TEMPLATE_FOLDER'],
            app.config.get('BUILTIN_TEMPLATE_FILES_DIR', ''),
            is_overridden
        )

    def load_rules():
        """加载所有自定义规则"""
        data = storage.get_json('rules')
        return data if data is not None else []

    def save_rules(rules):
        """保存自定义规则"""
        storage.put_json('rules', rules)

    def get_enabled_rules_for_type(grade_type):
        """获取指定批改类型下启用的规则"""
        rules = load_rules()
        enabled = []
        for r in rules:
            if r.get('enabled', True):
                applies = r.get('applies_to', 'all')
                if applies == 'all' or applies == grade_type:
                    enabled.append(r)
        return enabled

    def load_grading_standards():
        """加载用户上传的评分标准配置"""
        data = storage.get_json('grading_standards')
        return data if data is not None else {}

    def save_grading_standards(standards):
        """保存评分标准配置"""
        storage.put_json('grading_standards', standards)

    def get_grading_standard_info(grade_type):
        """获取指定题型的评分标准信息（是否有自定义、文件名等）"""
        builtin = BUILTIN_GRADING_STANDARDS.get(grade_type)
        if not builtin:
            return None

        standards = load_grading_standards()
        has_custom = grade_type in standards and standards[grade_type].get('standard_file')

        return {
            'grade_type': grade_type,
            'name': builtin['name'],
            'has_custom': has_custom,
            'standard_file': standards.get(grade_type, {}).get('standard_file') if has_custom else builtin['default_file'],
            'is_default': not has_custom
        }

    def get_grading_standard_content(grade_type):
        """获取指定题型的评分标准文本内容（用于注入AI prompt）"""
        info = get_grading_standard_info(grade_type)
        if not info:
            return ''

        standard_file = info['standard_file']
        if not standard_file:
            return ''

        file_path = os.path.join(app.config['TEMPLATE_FOLDER'], standard_file)
        if not os.path.exists(file_path):
            return ''

        ext = os.path.splitext(file_path)[1].lower()
        if ext in {'.docx', '.doc'}:
            return extract_text_from_docx(file_path) or ''
        elif ext == '.txt':
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read()
            except:
                return ''
        return ''

    def load_history():
        """加载批改历史记录"""
        data = storage.get_json('history')
        return data if data is not None else []

    def save_history(history):
        """保存批改历史记录"""
        storage.put_json('history', history)

    def add_history_record(record):
        """添加一条历史记录"""
        history = load_history()
        history.insert(0, record)
        if len(history) > 100:
            history = history[:100]
        save_history(history)

    def load_config():
        """加载系统配置（API Key等）"""
        data = storage.get_json('config')
        if data is not None:
            return data
        return {
            'doubao_api_key': '',
            'doubao_model': 'doubao-seed-2-0-mini-260428',
            'ai_enabled': False
        }

    def save_config(config):
        """保存系统配置"""
        storage.put_json('config', config)

    def load_templates():
        """加载所有自定义模板"""
        data = storage.get_json('templates')
        return data if data is not None else []

    def save_templates(templates):
        """保存自定义模板"""
        storage.put_json('templates', templates)

    # ==================== 路由 ====================

    @app.route('/')
    def index():
        return render_template('index.html')

    @app.route('/download/<filename>')
    def download(filename):
        """下载批改结果（已弃用：生成的文件直接流式返回，不存储在服务端）"""
        return jsonify({
            'success': False,
            'error': '该功能已调整：批改结果文件在生成时直接下载，不存储在服务器。请重新批改获取文件。'
        }), 410

    # ==================== AI配置API ====================

    @app.route('/api/config', methods=['GET'])
    def api_config_get():
        """获取系统配置（不含API Key明文）"""
        config = load_config()
        masked_config = {
            'ai_enabled': config.get('ai_enabled', False),
            'doubao_api_key_set': bool(config.get('doubao_api_key', '')),
            'doubao_model': config.get('doubao_model', 'doubao-seed-2-0-mini-260428'),
        }
        return jsonify({'success': True, 'config': masked_config})

    @app.route('/api/config', methods=['POST'])
    def api_config_save():
        """保存系统配置"""
        try:
            data = request.get_json() or {}
            config = load_config()

            if 'doubao_api_key' in data:
                config['doubao_api_key'] = data['doubao_api_key'].strip()
            if 'doubao_model' in data:
                config['doubao_model'] = data['doubao_model'].strip()

            config['ai_enabled'] = bool(config.get('doubao_api_key'))

            save_config(config)

            return jsonify({
                'success': True,
                'ai_enabled': config['ai_enabled'],
                'message': '配置已保存'
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/config/test', methods=['POST'])
    def api_config_test():
        """测试AI配置是否有效"""
        try:
            config = load_config()
            api_key = config.get('doubao_api_key', '')
            model_name = config.get('doubao_model', 'doubao-seed-2-0-mini-260428')

            if not api_key:
                return jsonify({'success': False, 'error': '请先配置API Key'})

            # 发送一条简单测试消息（Responses API格式）
            input_list = [{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": "请回复'OK'两个字"}
                ]
            }]
            result = call_doubao_chat(input_list, api_key, model_name, temperature=0)

            if result['success']:
                return jsonify({
                    'success': True,
                    'message': '连接成功！AI模型响应正常',
                    'reply': result['content'][:100]
                })
            else:
                return jsonify({'success': False, 'error': result['error']})

        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    # ==================== 自定义模板API ====================

    @app.route('/api/templates', methods=['GET'])
    def api_templates_list():
        """获取所有自定义模板列表"""
        templates = load_templates()
        return jsonify({
            'success': True,
            'templates': templates
        })

    @app.route('/api/templates', methods=['POST'])
    def api_template_create():
        """创建新的自定义模板"""
        try:
            name = request.form.get('name', '').strip()
            description = request.form.get('description', '').strip()
            full_score = int(request.form.get('full_score', 20))
            template_file = request.files.get('template_file')
            sample_text = request.form.get('sample_text', '').strip()
            style = request.form.get('style', 'default').strip()
            scoring_json = request.form.get('scoring_dimensions', '').strip()

            if not name:
                return jsonify({'success': False, 'error': '请输入模板名称'})

            template_id = str(uuid.uuid4())[:8]
            created_at = int(time.time())

            # 保存模板文件
            template_file_path = ''
            if template_file and template_file.filename:
                ext = os.path.splitext(template_file.filename)[1]
                template_file_path = f"template_{template_id}{ext}"
                save_path = os.path.join(app.config['TEMPLATE_FOLDER'], template_file_path)
                template_file.save(save_path)

            # 解析评分维度
            scoring_dimensions = []
            if scoring_json:
                try:
                    scoring_dimensions = json.loads(scoring_json)
                except:
                    pass

            # 保存模板信息
            templates = load_templates()
            new_template = {
                'id': template_id,
                'name': name,
                'description': description,
                'full_score': full_score,
                'template_file': template_file_path,
                'sample_text': sample_text,
                'style': style,
                'scoring_dimensions': scoring_dimensions,
                'created_at': created_at,
                'usage_count': 0
            }
            templates.append(new_template)
            save_templates(templates)

            return jsonify({
                'success': True,
                'template': new_template
            })

        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/templates/<template_id>', methods=['DELETE'])
    def api_template_delete(template_id):
        """删除自定义模板"""
        try:
            templates = load_templates()
            templates = [t for t in templates if t['id'] != template_id]
            save_templates(templates)
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    # ==================== 自定义规则API ====================

    @app.route('/api/rules', methods=['GET'])
    def api_rules_list():
        """获取所有自定义规则"""
        try:
            rules = load_rules()
            return jsonify({
                'success': True,
                'rules': rules
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/rules', methods=['POST'])
    def api_rule_create():
        """创建新的自定义规则"""
        try:
            data = request.get_json() or {}
            name = data.get('name', '').strip()
            content = data.get('content', '').strip()
            applies_to = data.get('applies_to', 'all').strip()

            if not name:
                return jsonify({'success': False, 'error': '请输入规则名称'})
            if not content:
                return jsonify({'success': False, 'error': '请输入规则内容'})

            rule_id = str(uuid.uuid4())[:8]
            new_rule = {
                'id': rule_id,
                'name': name,
                'content': content,
                'applies_to': applies_to,
                'enabled': True,
                'created_at': int(time.time())
            }

            rules = load_rules()
            rules.append(new_rule)
            save_rules(rules)

            return jsonify({
                'success': True,
                'rule': new_rule
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/rules/<rule_id>', methods=['PUT'])
    def api_rule_update(rule_id):
        """更新自定义规则"""
        try:
            data = request.get_json() or {}
            rules = load_rules()
            rule = next((r for r in rules if r['id'] == rule_id), None)

            if not rule:
                return jsonify({'success': False, 'error': '规则不存在'})

            if 'name' in data:
                rule['name'] = data['name'].strip()
            if 'content' in data:
                rule['content'] = data['content'].strip()
            if 'applies_to' in data:
                rule['applies_to'] = data['applies_to'].strip()
            if 'enabled' in data:
                rule['enabled'] = bool(data['enabled'])

            rule['updated_at'] = int(time.time())
            save_rules(rules)

            return jsonify({
                'success': True,
                'rule': rule
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/rules/<rule_id>', methods=['DELETE'])
    def api_rule_delete(rule_id):
        """删除自定义规则"""
        try:
            rules = load_rules()
            rules = [r for r in rules if r['id'] != rule_id]
            save_rules(rules)
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/rules/<rule_id>/toggle', methods=['POST'])
    def api_rule_toggle(rule_id):
        """切换规则启用/禁用状态"""
        try:
            rules = load_rules()
            rule = next((r for r in rules if r['id'] == rule_id), None)

            if not rule:
                return jsonify({'success': False, 'error': '规则不存在'})

            rule['enabled'] = not rule.get('enabled', True)
            save_rules(rules)

            return jsonify({
                'success': True,
                'enabled': rule['enabled']
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    # ==================== 历史记录API ====================

    @app.route('/api/history', methods=['GET'])
    def api_history():
        """获取批改历史记录（仅保存摘要，不存储生成的文件）"""
        try:
            history = load_history()
            return jsonify({'success': True, 'history': history})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/history/<record_id>', methods=['DELETE'])
    def api_history_delete(record_id):
        """删除历史记录（仅删除摘要记录，生成的文件未存储）"""
        try:
            history = load_history()
            history = [r for r in history if r['id'] != record_id]
            save_history(history)
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    # ==================== 模板管理API ====================

    @app.route('/api/templates/all', methods=['GET'])
    def api_templates_all():
        """获取所有模板（内置+自定义）"""
        try:
            overrides = load_builtin_overrides()
            # 内置模板
            builtin = []
            for key, t in BUILTIN_TEMPLATES.items():
                is_overridden = key in overrides
                # 如果被覆盖，使用覆盖后的版本显示
                display_t = t.copy()
                if is_overridden:
                    display_t.update(overrides[key])
                builtin.append({
                    'key': key,
                    'id': t['id'],
                    'name': display_t.get('name', t['name']),
                    'description': display_t.get('description', t['description']),
                    'full_score': display_t.get('full_score', t['full_score']),
                    'scoring_dimensions': display_t.get('scoring_dimensions', t['scoring_dimensions']),
                    'type': 'builtin',
                    'style': display_t.get('style', t.get('style', 'default')),
                    'is_overridden': is_overridden,
                    'has_template_file': (is_overridden and overrides[key].get('template_file') is not None) or (t.get('template_file') is not None),
                })
            # 自定义模板
            custom = load_templates()
            for t in custom:
                t['type'] = 'custom'
            return jsonify({
                'success': True,
                'builtin': builtin,
                'custom': custom,
                'total': len(builtin) + len(custom)
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/templates/detail/<template_id>', methods=['GET'])
    def api_template_detail(template_id):
        """获取单个模板详情"""
        try:
            overrides = load_builtin_overrides()
            # 先查找内置模板
            for key, t in BUILTIN_TEMPLATES.items():
                if t['id'] == template_id:
                    is_overridden = key in overrides
                    display_t = t.copy()
                    if is_overridden:
                        display_t.update(overrides[key])
                    return jsonify({
                        'success': True,
                        'template': {
                            'key': key,
                            'id': t['id'],
                            'name': display_t.get('name', t['name']),
                            'description': display_t.get('description', t['description']),
                            'full_score': display_t.get('full_score', t['full_score']),
                            'scoring_dimensions': display_t.get('scoring_dimensions', t['scoring_dimensions']),
                            'type': 'builtin',
                            'style': display_t.get('style', t.get('style', 'default')),
                            'is_overridden': is_overridden,
                        }
                    })
            # 再查找自定义模板
            templates = load_templates()
            template = next((t for t in templates if t['id'] == template_id), None)
            if template:
                template['type'] = 'custom'
                return jsonify({'success': True, 'template': template})
            return jsonify({'success': False, 'error': '模板不存在'})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/templates/<template_id>/content', methods=['GET'])
    def api_template_content(template_id):
        """获取模板文件的内容（用于预览模板格式）"""
        try:
            overrides = load_builtin_overrides()
            template_file = None
            template_name = ''
            is_overridden = False
            builtin_template_file = None

            # 先查找内置模板（可能有用户上传的覆盖文件）
            for key, t in BUILTIN_TEMPLATES.items():
                if t['id'] == template_id:
                    template_name = t['name']
                    builtin_template_file = t.get('template_file')
                    if key in overrides:
                        override_file = overrides[key].get('template_file')
                        if override_file:
                            template_file = override_file
                            is_overridden = True
                    break

            # 如果用户没有覆盖，使用内置模板文件
            if not template_file and builtin_template_file:
                template_file = builtin_template_file

            # 再查找自定义模板
            if not template_file:
                templates = load_templates()
                template = next((t for t in templates if t['id'] == template_id), None)
                if template:
                    template_file = template.get('template_file')
                    template_name = template.get('name', '')

            if not template_file:
                return jsonify({
                    'success': True,
                    'has_file': False,
                    'template_name': template_name,
                    'content': '',
                    'message': '该模板暂无上传的格式文件'
                })

            # 查找模板文件路径（优先用户覆盖，其次内置）
            template_path = get_template_file_path_local(template_file, is_overridden)
            if not template_path or not os.path.exists(template_path):
                return jsonify({
                    'success': True,
                    'has_file': False,
                    'template_name': template_name,
                    'content': '',
                    'message': '模板文件不存在'
                })

            ext = os.path.splitext(template_path)[1].lower()
            # 如果没有扩展名，尝试检测是否为Word文档
            if not ext:
                try:
                    import zipfile
                    with zipfile.ZipFile(template_path) as z:
                        if 'word/document.xml' in z.namelist():
                            ext = '.docx'
                except:
                    pass

            content = ''
            file_type = 'text'

            if ext in {'.docx', '.doc'}:
                content = extract_text_from_docx(template_path) or ''
                file_type = 'docx'
            elif ext in {'.txt'}:
                with open(template_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                file_type = 'txt'
            elif ext in {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'}:
                # 图片类型，返回base64
                import base64
                with open(template_path, 'rb') as f:
                    img_data = f.read()
                content = base64.b64encode(img_data).decode('utf-8')
                file_type = 'image'
            elif ext in {'.pdf'}:
                # PDF类型，将所有页转为图片预览（使用 PyMuPDF，纯Python实现）
                try:
                    import base64
                    import io
                    import fitz  # PyMuPDF
                    pdf_doc = fitz.open(template_path)
                    if pdf_doc.page_count > 0:
                        pages = []
                        for page_num in range(pdf_doc.page_count):
                            page = pdf_doc[page_num]
                            # 渲染为图片，150 DPI
                            pix = page.get_pixmap(dpi=150)
                            img_bytes = pix.tobytes("jpeg")
                            page_data = base64.b64encode(img_bytes).decode('utf-8')
                            pages.append(page_data)
                        pdf_doc.close()
                        # 如果只有一页，保持向后兼容
                        if len(pages) == 1:
                            content = pages[0]
                        else:
                            # 多页情况，用特殊标记分隔
                            content = '|||'.join(pages)
                        file_type = 'pdf_image'
                    else:
                        pdf_doc.close()
                        content = ''
                except Exception as e:
                    print(f"PDF预览转换失败: {e}")
                    content = ''
            else:
                # 尝试以文本读取
                try:
                    with open(template_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    file_type = 'text'
                except:
                    content = ''

            return jsonify({
                'success': True,
                'has_file': True,
                'template_name': template_name,
                'content': content,
                'file_type': file_type,
                'file_name': template_file
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/templates/builtin/<template_key>/override', methods=['POST'])
    def api_builtin_template_override(template_key):
        """上传用户自定义模板来覆盖内置模板"""
        try:
            if template_key not in BUILTIN_TEMPLATES:
                return jsonify({'success': False, 'error': '未知的模板类型'})

            name = request.form.get('name', '').strip()
            description = request.form.get('description', '').strip()
            full_score = request.form.get('full_score', type=int)
            scoring_dimensions_str = request.form.get('scoring_dimensions', '[]')
            template_file = request.files.get('template_file')

            if not name:
                return jsonify({'success': False, 'error': '请输入模板名称'})
            if not full_score:
                return jsonify({'success': False, 'error': '请输入满分'})

            try:
                scoring_dimensions = json.loads(scoring_dimensions_str)
            except:
                scoring_dimensions = []

            # 保存模板文件
            template_filename = None
            if template_file and template_file.filename:
                filename = secure_filename(template_file.filename)
                ext = os.path.splitext(filename)[1]
                template_filename = f'override_{template_key}_{int(time.time())}{ext}'
                template_file.save(os.path.join(app.config['TEMPLATE_FOLDER'], template_filename))

            # 构建覆盖数据
            override_data = {
                'name': name,
                'description': description,
                'full_score': full_score,
                'scoring_dimensions': scoring_dimensions,
                'overridden_at': int(time.time()),
            }
            if template_filename:
                override_data['template_file'] = template_filename

            overrides = load_builtin_overrides()
            # 如果之前有覆盖，先删除旧的模板文件
            if template_key in overrides and template_filename:
                old_file = overrides[template_key].get('template_file')
                if old_file:
                    old_path = os.path.join(app.config['TEMPLATE_FOLDER'], old_file)
                    if os.path.exists(old_path):
                        os.remove(old_path)

            overrides[template_key] = override_data
            save_builtin_overrides(overrides)

            return jsonify({'success': True, 'message': '模板更新成功'})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/templates/builtin/<template_key>/reset', methods=['POST'])
    def api_builtin_template_reset(template_key):
        """恢复内置模板为默认（删除用户覆盖）"""
        try:
            if template_key not in BUILTIN_TEMPLATES:
                return jsonify({'success': False, 'error': '未知的模板类型'})

            overrides = load_builtin_overrides()
            if template_key in overrides:
                # 删除模板文件
                template_file = overrides[template_key].get('template_file')
                if template_file:
                    file_path = os.path.join(app.config['TEMPLATE_FOLDER'], template_file)
                    if os.path.exists(file_path):
                        os.remove(file_path)
                del overrides[template_key]
                save_builtin_overrides(overrides)

            return jsonify({'success': True, 'message': '已恢复为内置模板'})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    # ==================== 评分标准API ====================

    @app.route('/api/grading-standards/<grade_type>', methods=['GET'])
    def api_grading_standard_info(grade_type):
        """获取指定题型的评分标准信息"""
        try:
            info = get_grading_standard_info(grade_type)
            if not info:
                return jsonify({'success': False, 'error': '该题型暂无评分标准配置'})
            return jsonify({
                'success': True,
                'info': info
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/grading-standards/<grade_type>/content', methods=['GET'])
    def api_grading_standard_content(grade_type):
        """获取评分标准文件内容（用于预览）"""
        try:
            info = get_grading_standard_info(grade_type)
            if not info:
                return jsonify({'success': False, 'error': '该题型暂无评分标准配置'})

            standard_file = info['standard_file']
            if not standard_file:
                return jsonify({
                    'success': True,
                    'has_file': False,
                    'message': '暂无评分标准文件'
                })

            file_path = os.path.join(app.config['TEMPLATE_FOLDER'], standard_file)
            if not os.path.exists(file_path):
                return jsonify({
                    'success': True,
                    'has_file': False,
                    'message': '评分标准文件不存在'
                })

            ext = os.path.splitext(file_path)[1].lower()
            content = ''
            file_type = 'text'

            if ext in {'.docx', '.doc'}:
                content = extract_text_from_docx(file_path) or ''
                file_type = 'docx'
            elif ext in {'.txt'}:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                file_type = 'txt'
            elif ext in {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'}:
                import base64
                with open(file_path, 'rb') as f:
                    img_data = f.read()
                content = base64.b64encode(img_data).decode('utf-8')
                file_type = 'image'
            elif ext in {'.pdf'}:
                try:
                    import base64
                    import fitz  # PyMuPDF
                    pdf_doc = fitz.open(file_path)
                    if pdf_doc.page_count > 0:
                        pages = []
                        for page_num in range(pdf_doc.page_count):
                            page = pdf_doc[page_num]
                            pix = page.get_pixmap(dpi=150)
                            img_bytes = pix.tobytes("jpeg")
                            page_data = base64.b64encode(img_bytes).decode('utf-8')
                            pages.append(page_data)
                        pdf_doc.close()
                        if len(pages) == 1:
                            content = pages[0]
                        else:
                            content = '|||'.join(pages)
                        file_type = 'pdf_image'
                    else:
                        pdf_doc.close()
                        content = ''
                except Exception as e:
                    print(f"PDF预览转换失败: {e}")
                    content = ''
            else:
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    file_type = 'text'
                except:
                    content = ''

            return jsonify({
                'success': True,
                'has_file': True,
                'name': info['name'],
                'content': content,
                'file_type': file_type,
                'file_name': standard_file,
                'is_default': info['is_default']
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/grading-standards/<grade_type>/upload', methods=['POST'])
    def api_grading_standard_upload(grade_type):
        """上传用户自定义评分标准"""
        try:
            if grade_type not in BUILTIN_GRADING_STANDARDS:
                return jsonify({'success': False, 'error': '未知的题型'})

            standard_file = request.files.get('standard_file')
            if not standard_file or not standard_file.filename:
                return jsonify({'success': False, 'error': '请上传评分标准文件'})

            ext = os.path.splitext(standard_file.filename)[1].lower()
            if ext not in {'.docx', '.doc', '.pdf', '.txt'}:
                return jsonify({'success': False, 'error': '仅支持 Word、PDF、TXT 格式文件'})

            # 保存文件
            timestamp = int(time.time())
            filename = f"standard_{grade_type}_{timestamp}{ext}"
            save_path = os.path.join(app.config['TEMPLATE_FOLDER'], filename)
            standard_file.save(save_path)

            # 更新配置
            standards = load_grading_standards()
            # 先删除旧的自定义文件
            if grade_type in standards:
                old_file = standards[grade_type].get('standard_file')
                if old_file and old_file != BUILTIN_GRADING_STANDARDS[grade_type]['default_file']:
                    old_path = os.path.join(app.config['TEMPLATE_FOLDER'], old_file)
                    if os.path.exists(old_path):
                        os.remove(old_path)

            standards[grade_type] = {
                'standard_file': filename,
                'updated_at': timestamp
            }
            save_grading_standards(standards)

            info = get_grading_standard_info(grade_type)
            return jsonify({
                'success': True,
                'message': '评分标准已更新',
                'info': info
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/api/grading-standards/<grade_type>/reset', methods=['POST'])
    def api_grading_standard_reset(grade_type):
        """重置评分标准为默认"""
        try:
            if grade_type not in BUILTIN_GRADING_STANDARDS:
                return jsonify({'success': False, 'error': '未知的题型'})

            standards = load_grading_standards()
            if grade_type in standards:
                # 删除用户上传的文件
                standard_file = standards[grade_type].get('standard_file')
                if standard_file and standard_file != BUILTIN_GRADING_STANDARDS[grade_type]['default_file']:
                    file_path = os.path.join(app.config['TEMPLATE_FOLDER'], standard_file)
                    if os.path.exists(file_path):
                        os.remove(file_path)
                del standards[grade_type]
                save_grading_standards(standards)

            info = get_grading_standard_info(grade_type)
            return jsonify({
                'success': True,
                'message': '已恢复默认评分标准',
                'info': info
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    # ==================== 文字提取API ====================

    @app.route('/api/extract-text', methods=['POST'])
    def api_extract_text():
        """从图片中提取文字（OCR），供用户审核后再批改
        适用于作文和翻译类型
        """
        try:
            grade_type = request.form.get('type', 'essay')
            files = request.files.getlist('files')

            if not files:
                return jsonify({'success': False, 'error': '请上传图片文件'})

            # 检查AI配置
            config = load_config()
            api_key = config.get('doubao_api_key', '')
            model_name = config.get('doubao_model', 'doubao-seed-2-0-mini-260428')
            ai_enabled = config.get('ai_enabled', False)

            if not ai_enabled or not api_key:
                return jsonify({'success': False, 'error': '请先在设置中配置并启用AI模型'})

            # 生成唯一ID
            task_id = str(uuid.uuid4())[:8]

            # 保存上传的文件
            uploaded_files = []
            image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff'}
            for f in files:
                if f.filename:
                    filename = secure_filename(f.filename)
                    save_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_extract_{filename}")
                    f.save(save_path)
                    uploaded_files.append(save_path)

            # 分离图片文件
            image_files = [
                f for f in uploaded_files
                if os.path.splitext(f)[1].lower() in image_extensions
            ]

            # 如果有PDF，转为图片
            pdf_files = [
                f for f in uploaded_files
                if os.path.splitext(f)[1].lower() == '.pdf'
            ]
            if pdf_files:
                try:
                    import fitz  # PyMuPDF
                    for pdf_file in pdf_files:
                        pdf_doc = fitz.open(pdf_file)
                        for page_num in range(pdf_doc.page_count):
                            page = pdf_doc[page_num]
                            pix = page.get_pixmap(dpi=150)
                            img_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_extract_pdf_{page_num}.jpg")
                            pix.save(img_path)
                            image_files.append(img_path)
                        pdf_doc.close()
                except Exception as e:
                    return jsonify({'success': False, 'error': f'PDF转换失败：{str(e)}'})

            if not image_files:
                # 清理临时文件
                for f in uploaded_files:
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except Exception:
                        pass
                return jsonify({'success': False, 'error': '未找到图片文件，请上传图片或PDF'})

            # 调用AI提取文字
            extract_result = extract_text_with_doubao(
                image_files, api_key=api_key, model_name=model_name,
                grade_type=grade_type
            )

            # 清理临时文件
            for f in uploaded_files:
                try:
                    if os.path.exists(f):
                        os.remove(f)
                except Exception:
                    pass
            # 清理PDF转换的临时图片
            for f in image_files:
                if '_extract_pdf_' in os.path.basename(f):
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except Exception:
                        pass

            if extract_result.get('success'):
                return jsonify({
                    'success': True,
                    'text': extract_result['text']
                })
            else:
                return jsonify({
                    'success': False,
                    'error': extract_result.get('error', '文字提取失败')
                })

        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    # ==================== 批改API ====================

    @app.route('/api/grade', methods=['POST'])
    def api_grade():
        """批改API - 全部使用AI智能批改"""
        try:
            grade_type = request.form.get('type', 'essay')
            student_input = request.form.get('text', '')
            template_id = request.form.get('template_id', '')

            # 生成唯一ID
            task_id = str(uuid.uuid4())[:8]
            timestamp = int(time.time())

            # 检查AI配置
            config = load_config()
            api_key = config.get('doubao_api_key', '')
            model_name = config.get('doubao_model', 'doubao-seed-2-0-mini-260428')
            ai_enabled = config.get('ai_enabled', False)

            if not ai_enabled or not api_key:
                return jsonify({'success': False, 'error': '请先在设置中配置并启用AI模型'})

            # 获取适用的自定义规则
            custom_rules = get_enabled_rules_for_type(grade_type)

            # ========== 客观题批改 ==========
            if grade_type == 'objective':
                student_files = request.files.getlist('student_files')
                answer_files = request.files.getlist('answer_file')

                if not student_files:
                    return jsonify({'success': False, 'error': '请上传学生作业'})
                if not answer_files:
                    return jsonify({'success': False, 'error': '请上传题目解析'})

                # 保存学生作业文件
                uploaded_student_files = []
                for f in student_files:
                    if f.filename:
                        filename = secure_filename(f.filename)
                        save_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_student_{filename}")
                        f.save(save_path)
                        uploaded_student_files.append(save_path)

                # 保存题目解析文件
                uploaded_answer_files = []
                for f in answer_files:
                    if f.filename:
                        filename = secure_filename(f.filename)
                        save_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_answer_{filename}")
                        f.save(save_path)
                        uploaded_answer_files.append(save_path)

                # 学生作业转图片（如果是PDF）
                image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff'}
                student_images = [
                    f for f in uploaded_student_files
                    if os.path.splitext(f)[1].lower() in image_extensions
                ]

                # 如果有PDF，需要转图片（使用 PyMuPDF，纯Python实现）
                pdf_files = [
                    f for f in uploaded_student_files
                    if os.path.splitext(f)[1].lower() == '.pdf'
                ]
                if pdf_files:
                    try:
                        import fitz  # PyMuPDF
                        for pdf_file in pdf_files:
                            pdf_doc = fitz.open(pdf_file)
                            for page_num in range(pdf_doc.page_count):
                                page = pdf_doc[page_num]
                                pix = page.get_pixmap(dpi=150)
                                img_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_pdf_{page_num}.jpg")
                                pix.save(img_path)
                                student_images.append(img_path)
                            pdf_doc.close()
                    except Exception as e:
                        return jsonify({'success': False, 'error': f'PDF转换失败：{str(e)}'})

                if not student_images:
                    return jsonify({'success': False, 'error': '学生作业文件格式不支持，请上传图片或PDF'})

                # 读取题目解析（Word文档）
                answer_key_text = ''
                for ans_file in uploaded_answer_files:
                    ext = os.path.splitext(ans_file)[1].lower()
                    if ext in {'.docx', '.doc'}:
                        text = extract_text_from_docx(ans_file)
                        if text:
                            answer_key_text += text + '\n'

                if not answer_key_text:
                    return jsonify({'success': False, 'error': '未能读取题目解析内容，请检查文件格式'})

                # 获取模板
                template = get_effective_template('objective')
                title = '客观题批改'

                # 调用AI批改
                ai_result = grade_objective_with_doubao(
                    student_images, answer_key_text, template,
                    api_key=api_key,
                    model_name=model_name,
                    custom_rules=custom_rules,
                    template_folder=app.config['TEMPLATE_FOLDER']
                )

                if not ai_result or not ai_result.get('success'):
                    error_msg = ai_result.get('error', 'AI返回为空') if ai_result else 'AI调用失败'
                    return jsonify({'success': False, 'error': f'AI批改失败：{error_msg}'})

                result = ai_result
                result['ai_used'] = True

                # 生成Word文档（流式返回，不保存文件）
                today = datetime.date.today()
                date_str = f"{today.month}.{today.day}"
                project_name = '客观题'
                base_filename = f"{date_str}-{project_name}批改"
                output_filename = f"{base_filename}.docx"

                doc_bytesio = generate_objective_doc(result, None, title, student_images)
                file_data = base64.b64encode(doc_bytesio.getvalue()).decode('utf-8')

                # 保存历史记录
                history_record = {
                    'id': task_id,
                    'filename': output_filename,
                    'title': title,
                    'score': result.get('score', 0),
                    'full_score': result.get('full_score', 100),
                    'create_time': int(time.time()),
                    'grade_type': grade_type,
                }
                add_history_record(history_record)

                # 自动删除上传的源文件
                all_uploaded = uploaded_student_files + uploaded_answer_files
                for f in all_uploaded:
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except Exception:
                        pass
                # 也删除PDF转的临时图片
                for f in student_images:
                    if '_pdf_' in os.path.basename(f):
                        try:
                            if os.path.exists(f):
                                os.remove(f)
                        except Exception:
                            pass

                return jsonify({
                    'success': True,
                    'task_id': task_id,
                    'filename': output_filename,
                    'file_data': file_data,
                    'score': result.get('score', 0),
                    'full_score': result.get('full_score', 100),
                    'ai_used': True,
                    'rules_applied': len(custom_rules),
                    'rule_names': [r['name'] for r in custom_rules],
                })

            # ========== 客观题子题型（文字输入模式）==========
            objective_subtypes = ['objective_reading', 'objective_six_choose_four', 'objective_cloze', 'objective_grammar', 'objective_cat_fishing']
            if grade_type in objective_subtypes:
                # 获取表单数据
                wrong_questions_text = request.form.get('wrong_questions', '').strip()
                answer_files = request.files.getlist('answer_file')

                print(f"[DEBUG] 题型: {grade_type}")
                print(f"[DEBUG] 错题文本长度: {len(wrong_questions_text)}")
                print(f"[DEBUG] 答案文件数量: {len(answer_files)}")
                for i, f in enumerate(answer_files):
                    print(f"[DEBUG] 文件{i}: filename={f.filename}, content_type={f.content_type}")

                if not wrong_questions_text:
                    return jsonify({'success': False, 'error': '请输入错题信息'})
                if not answer_files:
                    return jsonify({'success': False, 'error': '请上传题目解析文档'})

                # 保存题目解析文件
                uploaded_answer_files = []
                for f in answer_files:
                    if f.filename:
                        filename = secure_filename(f.filename)
                        save_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_answer_{filename}")
                        f.save(save_path)
                        file_size = os.path.getsize(save_path)
                        print(f"[DEBUG] 文件已保存: {save_path}, 大小: {file_size} bytes")
                        uploaded_answer_files.append(save_path)

                # 读取题目解析（Word文档）
                answer_key_text = ''
                for ans_file in uploaded_answer_files:
                    ext = os.path.splitext(ans_file)[1].lower()
                    print(f"[DEBUG] 读取文件: {ans_file}, 扩展名: {ext}")
                    if ext in {'.docx', '.doc'}:
                        text = extract_text_from_docx(ans_file)
                        print(f"[DEBUG] 提取文本长度: {len(text) if text else 0}")
                        if text:
                            answer_key_text += text + '\n'

                if not answer_key_text:
                    return jsonify({'success': False, 'error': '未能读取题目解析内容，请检查文件格式'})

                # 获取模板
                template = get_effective_template(grade_type)
                title = template.get('name', '客观题批改')

                # 调用AI批改
                ai_result = grade_objective_text_with_doubao(
                    wrong_questions_text, answer_key_text, template,
                    api_key=api_key,
                    model_name=model_name,
                    custom_rules=custom_rules,
                    template_folder=app.config['TEMPLATE_FOLDER'],
                    builtin_template_files_dir=app.config.get('BUILTIN_TEMPLATE_FILES_DIR', '')
                )

                if not ai_result or not ai_result.get('success'):
                    error_msg = ai_result.get('error', 'AI返回为空') if ai_result else 'AI调用失败'
                    return jsonify({'success': False, 'error': f'AI批改失败：{error_msg}'})

                result = ai_result
                result['ai_used'] = True
                result['student_text'] = wrong_questions_text

                # 生成Word文档（流式返回，不保存文件）
                today = datetime.date.today()
                date_str = f"{today.month}.{today.day}"
                project_name = title
                base_filename = f"{date_str}-{project_name}批改"
                output_filename = f"{base_filename}.docx"

                doc_bytesio = generate_objective_doc(result, None, title, None, template=template)
                file_data = base64.b64encode(doc_bytesio.getvalue()).decode('utf-8')

                # 保存历史记录
                history_record = {
                    'id': task_id,
                    'filename': output_filename,
                    'title': title,
                    'score': result.get('score', 0),
                    'full_score': result.get('full_score', 100),
                    'create_time': int(time.time()),
                    'grade_type': grade_type,
                }
                add_history_record(history_record)

                # 自动删除上传的源文件
                for f in uploaded_answer_files:
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except:
                        pass

                return jsonify({
                    'success': True,
                    'result': {
                        'score': result.get('score', 0),
                        'full_score': result.get('full_score', 100),
                        'total_questions': result.get('total_questions', 0),
                        'correct_count': result.get('correct_count', 0),
                        'wrong_count': result.get('wrong_count', 0),
                        'filename': output_filename,
                        'file_data': file_data,
                    }
                })

            # ========== 其他题型批改 ==========
            files = request.files.getlist('files')

            # 保存上传的文件
            uploaded_files = []
            for f in files:
                if f.filename:
                    filename = secure_filename(f.filename)
                    save_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_{filename}")
                    f.save(save_path)
                    uploaded_files.append(save_path)

            # 筛选出图片文件
            image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff'}
            uploaded_images = [
                f for f in uploaded_files
                if os.path.splitext(f)[1].lower() in image_extensions
            ]

            # 获取对应的批改模板
            template = None
            title = ''

            if grade_type == 'custom':
                # 自定义模板
                templates = load_templates()
                template = next((t for t in templates if t['id'] == template_id), None)
                if not template:
                    return jsonify({'success': False, 'error': '模板不存在'})
                title = f'{template["name"]}批改'
                # 更新使用次数
                for t in templates:
                    if t['id'] == template_id:
                        t['usage_count'] = t.get('usage_count', 0) + 1
                save_templates(templates)
            else:
                # 内置模板（优先使用用户覆盖版本）
                template = get_effective_template(grade_type)
                if not template:
                    return jsonify({'success': False, 'error': f'未知的批改类型：{grade_type}'})
                # 设置标题
                title_map = {
                    'essay_zhongkao': '中考作文批改',
                    'essay_gaokao': '高考作文批改',
                    'essay_ielts': '雅思作文批改',
                    'summary': '概要写作批改',
                    'translation': '翻译批改',
                }
                title = title_map.get(grade_type, '作业批改')

            # ===== 概要写作讲义模式 =====
            if grade_type == 'summary':
                handout_mode = request.form.get('mode') == 'handout'
                if handout_mode:
                    # 讲义模式：图片=原文，docx/pdf=格式参考
                    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff'}
                    doc_extensions = {'.docx', '.doc', '.pdf', '.txt'}

                    # 分离图片文件和文档文件
                    handout_images = [f for f in uploaded_files if os.path.splitext(f)[1].lower() in image_extensions]
                    handout_docs = [f for f in uploaded_files if os.path.splitext(f)[1].lower() in doc_extensions]

                    # 提取格式参考文档的文本内容
                    format_reference_text = ''
                    for doc_file in handout_docs:
                        doc_text = extract_text_from_file(doc_file)
                        if doc_text:
                            format_reference_text += doc_text + '\n'

                    # 原文内容：如果有图片，用图片识别；否则用文本输入
                    passage_text = student_input

                    # 如果有图片且没有文本输入，提示AI识别图片中的原文
                    if handout_images and not passage_text.strip():
                        passage_text = '（请识别以上图片中的原文内容，基于该原文生成讲义）'

                    # 获取评分标准内容
                    scoring_standard_text = ''
                    if grade_type in BUILTIN_GRADING_STANDARDS:
                        scoring_standard_text = get_grading_standard_content(grade_type)

                    handout_result = generate_summary_handout_with_doubao(
                        passage_text, template,
                        scoring_standard_text=scoring_standard_text,
                        api_key=api_key,
                        model_name=model_name,
                        custom_rules=custom_rules,
                        template_folder=app.config['TEMPLATE_FOLDER'],
                        builtin_template_files_dir=app.config.get('BUILTIN_TEMPLATE_FILES_DIR', ''),
                        image_paths=handout_images if handout_images else None,
                        format_reference_text=format_reference_text
                    )

                    if not handout_result or not handout_result.get('success'):
                        error_msg = handout_result.get('error', 'AI返回为空') if handout_result else 'AI调用失败'
                        return jsonify({'success': False, 'error': f'AI生成失败：{error_msg}'})

                    result = handout_result
                    result['ai_used'] = True
                    result['is_handout'] = True

                    # 生成讲义Word文档（流式返回，不保存文件）
                    today = datetime.date.today()
                    date_str = f"{today.month}.{today.day}"
                    project_name = '概要写作讲义'
                    base_filename = f"{date_str}-{project_name}"
                    output_filename = f"{base_filename}.docx"

                    doc_bytesio = generate_summary_handout_doc(result, None, project_name)
                    file_data = base64.b64encode(doc_bytesio.getvalue()).decode('utf-8')

                    # 保存历史记录
                    history_record = {
                        'id': task_id,
                        'filename': output_filename,
                        'title': project_name,
                        'score': 0,
                        'full_score': template.get('full_score', 10),
                        'create_time': int(time.time()),
                        'grade_type': 'summary_handout',
                    }
                    add_history_record(history_record)

                    return jsonify({
                        'success': True,
                        'task_id': task_id,
                        'filename': output_filename,
                        'file_data': file_data,
                        'score': 0,
                        'full_score': template.get('full_score', 10),
                        'ai_used': True,
                        'rules_applied': len(custom_rules),
                        'rule_names': [r['name'] for r in custom_rules],
                        'result': {
                            'is_handout': True,
                            'title': result.get('title', ''),
                            'pattern': result.get('pattern_name', ''),
                        }
                    })

            # 获取评分标准内容
            grading_standard = None
            if grade_type and grade_type in BUILTIN_GRADING_STANDARDS:
                grading_standard = get_grading_standard_content(grade_type)

            # 提取上传的参考文档内容（docx/pdf/txt，仅作参考，不是学生答案）
            reference_documents = ''
            doc_extensions = {'.docx', '.doc', '.pdf', '.txt'}
            ref_files = [f for f in uploaded_files if os.path.splitext(f)[1].lower() in doc_extensions]
            for ref_file in ref_files:
                ref_text = extract_text_from_file(ref_file)
                if ref_text:
                    reference_documents += ref_text + '\n'

            # 使用AI智能批改
            ai_result = grade_with_doubao(
                student_input, template,
                image_paths=uploaded_images,
                api_key=api_key,
                model_name=model_name,
                custom_rules=custom_rules,
                grading_standard=grading_standard,
                template_folder=app.config['TEMPLATE_FOLDER'],
                builtin_template_files_dir=app.config.get('BUILTIN_TEMPLATE_FILES_DIR', ''),
                reference_documents=reference_documents
            )

            if not ai_result or not ai_result.get('success'):
                error_msg = ai_result.get('error', 'AI返回为空') if ai_result else 'AI调用失败'
                return jsonify({'success': False, 'error': f'AI批改失败：{error_msg}'})

            result = ai_result
            result['ai_used'] = True

            # 设置学生原文（优先使用AI识别的结果）
            if not result.get('student_text'):
                result['student_text'] = student_input or '（已上传图片，文字已由AI识别提取）'

            # 生成Word文档（流式返回，不保存文件）
            # 文件名格式：日期-批改项目，如 "9.9-中考作文批改.docx"
            today = datetime.date.today()
            date_str = f"{today.month}.{today.day}"

            # 简化标题（去掉"批改"和"模板"字样）
            project_name = title.replace('批改', '').replace('模板', '').strip()
            if not project_name:
                project_name = '作业'

            # 基础文件名
            base_filename = f"{date_str}-{project_name}批改"
            output_filename = f"{base_filename}.docx"

            doc_bytesio = generate_essay_doc(result, None, title, uploaded_images, template=template)
            file_data = base64.b64encode(doc_bytesio.getvalue()).decode('utf-8')

            # 保存历史记录
            history_record = {
                'id': task_id,
                'filename': output_filename,
                'title': title,
                'score': result['score'],
                'full_score': result['full_score'],
                'create_time': int(time.time()),
                'grade_type': grade_type,
            }
            add_history_record(history_record)

            # 自动删除上传的源文件
            for f in uploaded_files:
                try:
                    if os.path.exists(f):
                        os.remove(f)
                except Exception:
                    pass

            return jsonify({
                'success': True,
                'task_id': task_id,
                'filename': output_filename,
                'file_data': file_data,
                'score': result['score'],
                'full_score': result['full_score'],
                'ai_used': True,
                'rules_applied': len(custom_rules),
                'rule_names': [r['name'] for r in custom_rules],
            })

        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    return app
