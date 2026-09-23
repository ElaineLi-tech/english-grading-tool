"""
英语作业自动批改工具 - Flask Web应用
支持：作文批改、概要写作批改、翻译批改
"""

from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import uuid
import time
import json
import re
import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['TEMPLATE_FOLDER'] = 'templates_user'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMPLATE_FOLDER'], exist_ok=True)

# 模板存储文件
TEMPLATES_FILE = os.path.join(app.config['TEMPLATE_FOLDER'], 'templates.json')

# 内置模板覆盖文件（用户修改后的内置模板）
BUILTIN_OVERRIDES_FILE = os.path.join(app.config['TEMPLATE_FOLDER'], 'builtin_overrides.json')


def load_builtin_overrides():
    """加载用户对内置模板的覆盖配置"""
    if os.path.exists(BUILTIN_OVERRIDES_FILE):
        with open(BUILTIN_OVERRIDES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_builtin_overrides(overrides):
    """保存用户对内置模板的覆盖配置"""
    with open(BUILTIN_OVERRIDES_FILE, 'w', encoding='utf-8') as f:
        json.dump(overrides, f, ensure_ascii=False, indent=2)


def get_effective_template(grade_type):
    """获取实际生效的模板（优先使用用户覆盖版本）"""
    builtin = BUILTIN_TEMPLATES.get(grade_type)
    if not builtin:
        return None
    overrides = load_builtin_overrides()
    if grade_type in overrides:
        # 合并内置模板的基础信息和用户覆盖的信息
        effective = builtin.copy()
        effective.update(overrides[grade_type])
        effective['is_overridden'] = True
        return effective
    result = builtin.copy()
    result['is_overridden'] = False
    return result

# ==================== 自定义规则系统 ====================

RULES_FILE = os.path.join(app.config['TEMPLATE_FOLDER'], 'rules.json')


def load_rules():
    """加载所有自定义规则"""
    if os.path.exists(RULES_FILE):
        with open(RULES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_rules(rules):
    """保存自定义规则"""
    with open(RULES_FILE, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)


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


# ==================== 评分标准系统 ====================

GRADING_STANDARDS_FILE = os.path.join(app.config['TEMPLATE_FOLDER'], 'grading_standards.json')

# 内置默认评分标准（题型 -> 默认文件名）
BUILTIN_GRADING_STANDARDS = {
    'summary': {
        'default_file': 'standard_summary_default.docx',
        'name': '概要写作评分标准（3+2得分点）'
    }
}


def load_grading_standards():
    """加载用户上传的评分标准配置"""
    if os.path.exists(GRADING_STANDARDS_FILE):
        with open(GRADING_STANDARDS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_grading_standards(standards):
    """保存评分标准配置"""
    with open(GRADING_STANDARDS_FILE, 'w', encoding='utf-8') as f:
        json.dump(standards, f, ensure_ascii=False, indent=2)


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


# 历史记录存储文件
HISTORY_FILE = os.path.join(app.config['TEMPLATE_FOLDER'], 'history.json')


def load_history():
    """加载批改历史记录"""
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_history(history):
    """保存批改历史记录"""
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def add_history_record(record):
    """添加一条历史记录"""
    history = load_history()
    history.insert(0, record)  # 最新的放在最前面
    # 最多保留100条记录
    if len(history) > 100:
        history = history[:100]
    save_history(history)


# ==================== AI配置管理 ====================

CONFIG_FILE = os.path.join(app.config['TEMPLATE_FOLDER'], 'config.json')


def load_config():
    """加载系统配置（API Key等）"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {
        'doubao_api_key': '',
        'doubao_model': 'doubao-seed-2-0-mini-260428',
        'ai_enabled': False
    }


def save_config(config):
    """保存系统配置"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ==================== 豆包API调用 ====================

import base64
import requests


def encode_image_to_base64(image_path):
    """将图片文件转为base64编码"""
    with open(image_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')


def call_doubao_chat(input_list, api_key, model_name='doubao-seed-2-0-mini-260428', temperature=0.3):
    """调用豆包大模型Responses API
    input_list: [{"role": "user", "content": [{"type": "input_text", "text": "..."}, {"type": "input_image", "image_url": "..."}]}]
    """
    url = "https://ark.cn-beijing.volces.com/api/v3/responses"
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    
    payload = {
        "model": model_name,
        "input": input_list,
        "temperature": temperature,
        "max_output_tokens": 8000
    }
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        result = response.json()
        
        # 从Responses格式中提取输出文本
        content = ''
        if 'output' in result:
            for item in result['output']:
                if item.get('type') == 'message' and item.get('role') == 'assistant':
                    for c in item.get('content', []):
                        if c.get('type') == 'output_text':
                            content += c.get('text', '')
        
        return {
            'success': True,
            'content': content,
            'usage': result.get('usage', {})
        }
    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }


def build_grading_prompt(student_text, template, has_image=False, custom_rules=None, grading_standard=None):
    """构建批改prompt，让AI按指定模板格式返回JSON
    custom_rules: 用户自定义规则列表，每条规则包含name和content
    grading_standard: 评分标准文本内容，用于指导AI评分
    """
    
    scoring_dims = template.get('scoring_dimensions', [
        {'name': '内容', 'full_score': 8},
        {'name': '语言', 'full_score': 8},
        {'name': '组织结构', 'full_score': 4}
    ])
    
    dims_desc = '\n'.join([f"  - {d['name']}（满分{d['full_score']}分）" for d in scoring_dims])
    
    # 构建自定义规则部分
    rules_section = ''
    if custom_rules and len(custom_rules) > 0:
        rules_lines = []
        for i, rule in enumerate(custom_rules, 1):
            rule_name = rule.get('name', f'规则{i}')
            rule_content = rule.get('content', '')
            rules_lines.append(f"{i}. {rule_name}：{rule_content}")
        rules_section = "\n【自定义批改规则】\n请严格遵守以下用户自定义的批改规则：\n" + "\n".join(rules_lines) + "\n"
    
    # 构建评分标准部分
    standard_section = ''
    if grading_standard and grading_standard.strip():
        # 限制长度，避免prompt过长（取前3000字符）
        standard_text = grading_standard.strip()
        if len(standard_text) > 3000:
            standard_text = standard_text[:3000] + '\n...（内容过长，已截取核心部分）'
        standard_section = f"\n【评分标准参考】\n请严格按照以下评分标准进行评分和点评，评分标准是你打分的核心依据：\n{standard_text}\n"
    
    prompt = f"""你是一位专业的英语作文批改老师，请按照以下要求批改学生作文。

【批改模板】
模板名称：{template['name']}
模板说明：{template.get('description', '英语作文批改')}
满分：{template.get('full_score', 20)}分
评分维度：
{dims_desc}
{rules_section}
{standard_section}
【学生作文】
{student_text}

【批改要求】
请仔细阅读学生作文，按照模板的风格进行详细批改，返回严格的JSON格式，不要有其他文字。JSON结构如下：
{{
  "student_text": "识别/提取的学生作文原文（如果是图片请准确OCR识别文字）",
  "error_categories": [
    {{
      "category": "错误类型（如：大写错误、第三人称单数错误、拼写错误、语法错误、用词错误、介词错误、内容不完整等）",
      "errors": [
        {{
          "error_text": "学生原文中的错误表达",
          "analysis": "详细分析：为什么错了，语法/词汇规则是什么",
          "suggestion": "修改建议：正确的表达是什么，类似用法拓展"
        }}
      ]
    }}
  ],
  "praise": "先肯定学生的优点，鼓励的话",
  "scoring_dimensions": [
    {{
      "name": "维度名称",
      "full_score": 满分,
      "score": 得分,
      "evaluation": "该维度的详细评价"
    }}
  ],
  "model_essay": "参考范文（在学生原文基础上改写提升，所有修改/提升的部分请用**两个星号**包裹标记，例如：**improved** 表示此处有修改。结构更完整、用词更丰富、句式更多样）\n\n【好词好句学习】\n1. 好词/好句型 - 解释\n2. ...",
  "good_sentences": [
    "好词1 - 释义",
    "好词2 - 释义"
  ]
}}

注意：
1. 错误要分类归纳，同一类的错误放在一个category下
2. 每个错误要有具体的分析和建议，不能太笼统
3. 评分要客观，参考中考/高考评分标准
4. 范文要基于原文改写，不要脱离题目太远
5. 范文中所有修改/提升的部分（词汇变化、时态变化、新增的词、句式调整等）必须用**两个星号**包裹标记，例如：**However**、**went**。原文中没有的词或形式不同的词都要标记。
6. error_text必须是学生原文中出现的准确文本，方便在原文中定位标注
7. 返回纯JSON，不要有markdown格式或其他说明文字
8. 如果有图片输入，请仔细OCR识别所有图片中的文字，确保准确完整；如果有多张图片，将所有图片的文字内容合并后作为学生原文进行批改
"""
    return prompt


def extract_text_from_docx(file_path):
    """从Word文档中提取纯文本内容"""
    try:
        doc = Document(file_path)
        paragraphs = []
        for para in doc.paragraphs:
            if para.text.strip():
                paragraphs.append(para.text)
        # 也提取表格中的内容
        for table in doc.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    if cell.text.strip():
                        row_text.append(cell.text.strip())
                if row_text:
                    paragraphs.append(' | '.join(row_text))
        return '\n'.join(paragraphs)
    except Exception as e:
        print(f"读取Word文档失败: {e}")
        return ''


def build_objective_grading_prompt(answer_key_text, template, custom_rules=None, template_folder=None):
    """构建客观题批改的prompt
    template_folder: 模板文件夹路径，用于读取用户上传的模板文件
    """
    template_name = template.get('name', '客观题批改')
    
    # 读取用户上传的错题分析模板文件内容
    error_analysis_template = ''
    template_file = template.get('template_file')
    if template_file and template_folder:
        template_path = os.path.join(template_folder, template_file)
        if os.path.exists(template_path):
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
            if ext in {'.docx', '.doc'}:
                template_content = extract_text_from_docx(template_path)
                if template_content:
                    error_analysis_template = template_content
    
    custom_rules_text = ''
    if custom_rules and len(custom_rules) > 0:
        custom_rules_text = '\n【自定义批改规则】\n'
        for i, rule in enumerate(custom_rules, 1):
            custom_rules_text += f"{i}. {rule.get('content', '')}\n"
    
    error_analysis_section = ''
    if error_analysis_template:
        error_analysis_section = f"""
【错题分析格式模板】
以下是错题分析的模板范例，请严格按照这个模板的格式、结构、用语习惯来生成每道错题的分析：
{error_analysis_template}

重要：每道错题的完整分析内容（包括题号、正确答案、学生答案、解析等全部部分）都要放在 error_analysis 字段中，严格参照模板的格式来写。
"""
    
    prompt = f"""你是一个专业的英语老师，负责批改学生的客观题作业（选择题、填空题等）。

【批改模板】
{template_name}
{error_analysis_section}
【题目解析（参考答案+解析）】
{answer_key_text}

【批改要求】
1. 仔细识别学生作业图片中的所有题目和学生答案
2. 对照题目解析中的正确答案，找出所有做错的题目
3. 如果提供了【错题分析格式模板】，每道错题的 error_analysis 字段必须严格按照模板的格式来写，包括题号的写法、正确答案的格式、学生答案的格式、解析的格式等，完全参照模板范例
4. 错题分析要详细，讲清楚知识点和正确思路
5. 统计总题数、做对题数、做错题数、得分率
6. 返回严格的JSON格式，结构如下：
{{
  "total_questions": 总题数,
  "correct_count": 做对题数,
  "wrong_count": 做错题数,
  "score": 得分（百分制）,
  "full_score": 100,
  "wrong_questions": [
    {{
      "question_number": "题号（如：第11空、第3题）",
      "question_text": "题目内容（可留空）",
      "student_answer": "学生答案",
      "correct_answer": "正确答案",
      "error_analysis": "完整的错题分析，严格按照错题分析格式模板的格式来写，包含题号、正确答案、学生答案、解析等全部内容",
      "knowledge_point": "知识点"
    }}
  ],
  "student_text": "学生作业的完整文字内容（OCR识别结果）",
  "overall_comment": "整体评价和建议"
}}
7. 返回纯JSON，不要有markdown格式或其他说明文字
{custom_rules_text}
"""
    return prompt


def grade_objective_with_doubao(student_images, answer_key_text, template, api_key='', model_name='doubao-seed-2-0-mini-260428', custom_rules=None, template_folder=None):
    """使用豆包AI进行客观题批改
    student_images: 学生作业图片路径列表
    answer_key_text: 题目解析文本内容
    template: 批改模板
    template_folder: 模板文件夹路径
    """
    if not api_key:
        return None
    
    import json
    
    # 构建 input 列表（Responses API格式）
    content = []
    
    # 添加所有学生作业图片
    if student_images and len(student_images) > 0:
        for img_path in student_images:
            img_base64 = encode_image_to_base64(img_path)
            content.append({
                "type": "input_image",
                "image_url": f"data:image/jpeg;base64,{img_base64}"
            })
    
    # 构建prompt
    prompt = build_objective_grading_prompt(answer_key_text, template, custom_rules=custom_rules, template_folder=template_folder)
    
    content.append({
        "type": "input_text",
        "text": prompt
    })
    
    input_list = [{"role": "user", "content": content}]
    
    # 调用API
    result = call_doubao_chat(input_list, api_key, model_name, temperature=0.3)
    
    if not result['success']:
        return {
            'success': False,
            'error': f'AI调用失败：{result["error"]}'
        }
    
    output_text = result.get('content', '')
    
    if not output_text:
        return {'success': False, 'error': 'AI返回为空'}
    
    # 尝试提取JSON
    try:
        # 先尝试直接解析
        result_data = json.loads(output_text)
    except:
        # 尝试从markdown代码块中提取
        import re
        json_match = re.search(r'\{[\s\S]*\}', output_text)
        if json_match:
            result_data = json.loads(json_match.group(0))
        else:
            return {'success': False, 'error': f'AI返回格式错误：{output_text[:200]}'}
    
    result_data['success'] = True
    return result_data


def grade_with_doubao(student_text, template, image_paths=None, api_key='', model_name='doubao-seed-2-0-mini-260428', custom_rules=None, grading_standard=None):
    """使用豆包AI进行智能批改
    custom_rules: 用户自定义规则列表
    grading_standard: 评分标准文本内容
    """
    
    if not api_key:
        return None
    
    # 构建 input 列表（Responses API格式）
    content = []
    
    # 如果有图片，添加所有图片
    if image_paths and len(image_paths) > 0:
        for img_path in image_paths:
            img_base64 = encode_image_to_base64(img_path)
            content.append({
                "type": "input_image",
                "image_url": f"data:image/jpeg;base64,{img_base64}"
            })
    
    # prompt文字
    image_count = len(image_paths) if image_paths else 0
    if image_count > 1 and not student_text:
        prompt_text = f"（共{image_count}张图片，请识别所有图片中的文字内容并合并后进行批改）"
    else:
        prompt_text = student_text or "（请识别图片中的学生作文并批改）"
    prompt = build_grading_prompt(prompt_text, template, has_image=bool(image_paths), custom_rules=custom_rules, grading_standard=grading_standard)
    content.append({
        "type": "input_text",
        "text": prompt
    })
    
    input_list = [{"role": "user", "content": content}]
    
    # 调用API
    result = call_doubao_chat(input_list, api_key, model_name, temperature=0.3)
    
    if not result['success']:
        return {
            'success': False,
            'error': f'AI调用失败：{result["error"]}'
        }
    
    # 解析返回的JSON
    try:
        content = result['content']
        # 尝试提取JSON（去掉可能的markdown代码块标记）
        content = content.strip()
        if content.startswith('```json'):
            content = content[7:]
        if content.startswith('```'):
            content = content[3:]
        if content.endswith('```'):
            content = content[:-3]
        content = content.strip()
        
        ai_result = json.loads(content)
        
        # 计算总分
        total_score = 0
        for dim in ai_result.get('scoring_dimensions', []):
            total_score += dim.get('score', 0)
        
        ai_result['score'] = total_score
        ai_result['full_score'] = template.get('full_score', 20)
        ai_result['success'] = True
        ai_result['template_name'] = template['name']
        ai_result['usage'] = result.get('usage', {})
        
        return ai_result
        
    except Exception as e:
        return {
            'success': False,
            'error': f'AI返回格式解析失败：{str(e)}\n返回内容：{result.get("content", "")[:500]}'
        }


# ==================== 内置批改模板 ====================

BUILTIN_TEMPLATES = {
    'essay_zhongkao': {
        'id': 'builtin_zhongkao',
        'name': '中考英语作文',
        'description': '上海中考英语作文20分制批改模板（内容8+语言8+组织结构4），按错误类型分类批改，含范文改写和好词好句学习。',
        'full_score': 20,
        'scoring_dimensions': [
            {'name': '内容', 'full_score': 8},
            {'name': '语言', 'full_score': 8},
            {'name': '组织结构', 'full_score': 4}
        ],
        'error_types': ['大写错误', '第三人称单数错误', '拼写错误', '表达错误', '介词错误', '内容不完整'],
        'has_model_essay': True,
        'has_good_sentences': True,
        'style': 'chu_zhongkao'
    },
    'essay_gaokao': {
        'id': 'builtin_gaokao',
        'name': '高考英语作文',
        'description': '高考英语作文批改标准，满分25分，从内容、语言、组织结构三个维度评分。',
        'full_score': 25,
        'scoring_dimensions': [
            {'name': '内容', 'full_score': 10},
            {'name': '语言', 'full_score': 10},
            {'name': '组织结构', 'full_score': 5}
        ],
        'style': 'gaokao'
    },
    'essay_ielts': {
        'id': 'builtin_ielts',
        'name': '雅思作文',
        'description': '雅思写作批改标准，满分9分，从任务完成度、连贯与衔接、词汇资源、语法范围与准确性四个维度评分。',
        'full_score': 9,
        'scoring_dimensions': [
            {'name': '任务完成度', 'full_score': 9},
            {'name': '连贯与衔接', 'full_score': 9},
            {'name': '词汇资源', 'full_score': 9},
            {'name': '语法范围与准确性', 'full_score': 9}
        ],
        'style': 'ielts'
    },
    'summary': {
        'id': 'builtin_summary',
        'name': '概要写作',
        'description': '概要写作批改标准，满分10分，从内容要点、语言表达、组织结构三个维度评分。',
        'full_score': 10,
        'scoring_dimensions': [
            {'name': '内容要点', 'full_score': 5},
            {'name': '语言表达', 'full_score': 3},
            {'name': '组织结构', 'full_score': 2}
        ],
        'style': 'summary'
    },
    'translation': {
        'id': 'builtin_translation',
        'name': '翻译',
        'description': '英语翻译批改标准，满分15分，从准确性、流畅性、用词地道三个维度评分。',
        'full_score': 15,
        'scoring_dimensions': [
            {'name': '准确性', 'full_score': 7},
            {'name': '流畅性', 'full_score': 5},
            {'name': '用词地道', 'full_score': 3}
        ],
        'style': 'translation'
    },
    'objective': {
        'id': 'builtin_objective',
        'name': '客观题批改',
        'description': '客观题（选择题、填空题等）自动批改，对照学生作业和题目解析，自动找出错题并生成详细的错题分析。',
        'full_score': 100,
        'scoring_dimensions': [
            {'name': '正确率', 'full_score': 100}
        ],
        'style': 'objective',
        'has_error_analysis': True
    }
}


def load_templates():
    """加载所有自定义模板"""
    if os.path.exists(TEMPLATES_FILE):
        with open(TEMPLATES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_templates(templates):
    """保存自定义模板"""
    with open(TEMPLATES_FILE, 'w', encoding='utf-8') as f:
        json.dump(templates, f, ensure_ascii=False, indent=2)

# 颜色常量
BLACK = RGBColor(0, 0, 0)
BLUE = RGBColor(0x00, 0x70, 0xC0)
RED = RGBColor(0xFF, 0x00, 0x00)
GRAY = RGBColor(0x80, 0x80, 0x80)
GREEN = RGBColor(0x00, 0x80, 0x00)
ORANGE = RGBColor(0xFF, 0x8C, 0x00)
FONT_CN = '微软雅黑'
FONT_EN = 'Times New Roman'


def sf(run, size=12, bold=False, color=None, italic=False, underline=False):
    """设置字体格式"""
    run.font.name = FONT_EN
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    if color:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), FONT_EN)
    rFonts.set(qn('w:hAnsi'), FONT_EN)
    rFonts.set(qn('w:eastAsia'), FONT_CN)


def add_para(doc, runs_data, size=12, align=None, space_before=0, space_after=4,
             line_spacing=1.5, indent=None):
    """添加带多种格式的段落"""
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if indent:
        p.paragraph_format.first_line_indent = Pt(indent)
    for data in runs_data:
        if isinstance(data, str):
            text, bold, color = data, False, BLACK
        else:
            text = data[0]
            bold = data[1] if len(data) > 1 else False
            color = data[2] if len(data) > 2 else BLACK
            italic = data[3] if len(data) > 3 else False
            underline = data[4] if len(data) > 4 else False
        run = p.add_run(text)
        sf(run, size, bold, color, italic, underline)
    return p


def create_doc_with_title(title, subtitle):
    """创建带标题的Word文档"""
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(2.54)
        sec.bottom_margin = Cm(2.54)
        sec.left_margin = Cm(2.54)
        sec.right_margin = Cm(2.54)
    style = doc.styles['Normal']
    style.font.name = FONT_EN
    style.font.size = Pt(12)
    rPr = style.element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), FONT_EN)
    rFonts.set(qn('w:hAnsi'), FONT_EN)
    rFonts.set(qn('w:eastAsia'), FONT_CN)
    rPr.insert(0, rFonts)
    add_para(doc, [(title, True, BLUE)], size=16, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_para(doc, [(subtitle, False, GRAY)], size=10, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    return doc


def extract_english_correction(suggestion):
    """从AI的suggestion中提取纯英文的正确答案
    去掉中文解释、"改为："等前缀，只保留英文内容
    """
    if not suggestion:
        return ''
    
    text = suggestion.strip()
    
    # 去掉开头的"改为："、"答案："等中文前缀
    import re
    text = re.sub(r'^(改为|改成|正确(答案|表达|形式)?|应该是|应为|建议[用为]?)[：: ]*', '', text)
    
    # 提取英文部分（取第一个中文之前的所有英文内容）
    # 匹配英文单词、标点、数字等，直到遇到中文字符
    match = re.match(r'^[a-zA-Z0-9\s\.,\?!\-\'\"\(\)]+', text)
    if match:
        english = match.group(0).strip()
        if english:
            return english
    
    # 如果没找到纯英文，尝试提取第一个英文单词或短语
    words = re.findall(r'[a-zA-Z]+[\sa-zA-Z]*', text)
    if words:
        return words[0].strip()
    
    # 实在不行返回原文
    return text


def build_annotated_student_text_runs(student_text, result):
    """构建带标注的学生原文runs列表
    将错误处用红色粗体标出，并在后面加上"改为：正确答案"
    返回格式: [(text, bold, color), ...]
    """
    # 收集所有错误（去重，同一错误文本只保留一个）
    errors = []
    seen_errors = set()
    if 'error_categories' in result and result['error_categories']:
        for cat in result['error_categories']:
            for err in cat.get('errors', []):
                err_text = err.get('error_text', '').strip()
                if err_text and err_text.lower() not in seen_errors:
                    seen_errors.add(err_text.lower())
                    errors.append({
                        'error_text': err_text,
                        'suggestion': extract_english_correction(err.get('suggestion', ''))
                    })
    elif 'errors' in result:
        for err in result['errors']:
            err_text = err.get('error', '').strip()
            if err_text and err_text.lower() not in seen_errors:
                seen_errors.add(err_text.lower())
                errors.append({
                    'error_text': err_text,
                    'suggestion': extract_english_correction(err.get('suggestion', ''))
                })
    
    if not errors or not student_text:
        return [(student_text, False, BLACK)]
    
    # 构建替换映射（按错误文本长度从长到短排序，避免部分匹配问题）
    replacements = {}
    for err in errors:
        err_text = err['error_text']
        suggestion = err['suggestion']
        if err_text and suggestion and err_text.lower() not in {k.lower() for k in replacements.keys()}:
            replacements[err_text] = suggestion
    
    if not replacements:
        return [(student_text, False, BLACK)]
    
    # 按长度从长到短排序，避免短的先替换掉长的一部分
    sorted_errors = sorted(replacements.keys(), key=len, reverse=True)
    
    # 用统一结构存储segments: {'text': str, 'is_error': bool, 'suggestion': str}
    segments = [{'text': student_text, 'is_error': False, 'suggestion': ''}]
    
    for err_text in sorted_errors:
        new_segments = []
        suggestion = replacements[err_text]
        for seg in segments:
            if seg['is_error']:
                # 已经是错误段，不再处理
                new_segments.append(seg)
                continue
            
            seg_text = seg['text']
            # 查找所有出现位置（不区分大小写）
            lower_text = seg_text.lower()
            lower_err = err_text.lower()
            pos = 0
            while True:
                idx = lower_text.find(lower_err, pos)
                if idx == -1:
                    if pos < len(seg_text):
                        new_segments.append({
                            'text': seg_text[pos:],
                            'is_error': False,
                            'suggestion': ''
                        })
                    break
                
                # 添加前面的正常文本
                if idx > pos:
                    new_segments.append({
                        'text': seg_text[pos:idx],
                        'is_error': False,
                        'suggestion': ''
                    })
                
                # 添加错误文本段
                actual_err = seg_text[idx:idx + len(err_text)]
                new_segments.append({
                    'text': actual_err,
                    'is_error': True,
                    'suggestion': suggestion
                })
                pos = idx + len(err_text)
        
        segments = new_segments
    
    # 将segments转换为runs格式
    runs = []
    for seg in segments:
        if not seg['is_error']:
            # 正常文本
            runs.append((seg['text'], False, BLACK))
        else:
            # 错误文本（红色粗体） + 改为： + 正确内容（绿色粗体）
            runs.append((seg['text'], True, RED))
            runs.append(('改为：', False, GREEN))
            runs.append((seg['suggestion'], True, GREEN))
    
    return runs


def parse_bold_markdown(text):
    """解析markdown格式的粗体标记（**text**）
    返回格式: [(text, is_bold), ...]
    """
    if not text:
        return []
    
    parts = []
    import re
    # 分割 **bold** 标记
    pattern = re.compile(r'\*\*(.+?)\*\*')
    last_end = 0
    
    for match in pattern.finditer(text):
        start = match.start()
        end = match.end()
        bold_text = match.group(1)
        
        if start > last_end:
            parts.append((text[last_end:start], False))
        parts.append((bold_text, True))
        last_end = end
    
    if last_end < len(text):
        parts.append((text[last_end:], False))
    
    if not parts:
        parts.append((text, False))
    
    return parts


def diff_and_bold_model_essay(student_text, model_essay):
    """对比学生原文和范文，自动找出修改的单词并用**加粗标记**
    只在范文中没有**标记时使用
    """
    import re
    
    # 如果范文中已经有**标记，直接返回
    if '**' in model_essay:
        return model_essay
    
    # 提取原文的词（小写），用于对比（完全匹配）
    student_words_lower = set(w.lower() for w in re.findall(r"[a-zA-Z']+", student_text))
    
    def is_original_word(word):
        """判断一个单词是否在原文中出现过（完全相同，不考虑变形）
        词形变化也属于修改，应该加粗标记
        """
        w = word.lower()
        if w in student_words_lower:
            return True
        return False
    
    # 逐词处理范文
    tokens = re.findall(r"[a-zA-Z']+|[^a-zA-Z']+", model_essay)
    
    result_parts = []
    bold_buffer = []  # 累积要加粗的连续token
    
    def flush_bold():
        """将累积的加粗buffer写入结果"""
        nonlocal bold_buffer
        if bold_buffer:
            # 去掉开头和结尾的空格
            text = ''.join(bold_buffer)
            stripped = text.strip()
            if stripped:
                # 计算前后空格
                leading = text[:len(text) - len(text.lstrip())]
                trailing = text[len(text.rstrip()):]
                result_parts.append(leading)
                result_parts.append('**' + stripped + '**')
                result_parts.append(trailing)
            else:
                result_parts.append(text)
            bold_buffer = []
    
    for token in tokens:
        if re.match(r"[a-zA-Z']+", token):
            # 是单词
            if is_original_word(token):
                # 原文中有的词，结束加粗
                flush_bold()
                result_parts.append(token)
            else:
                # 新单词/修改过的词，加入加粗buffer
                bold_buffer.append(token)
        else:
            # 非字母token
            if bold_buffer:
                # 如果在加粗状态，空格和标点也加入加粗
                # 但如果是句号或换行，就结束加粗
                if token.strip() == '' or (len(token.strip()) == 1 and token.strip() in ",;"):
                    bold_buffer.append(token)
                else:
                    # 句号等结束标点，先flush再加标点
                    flush_bold()
                    result_parts.append(token)
            else:
                result_parts.append(token)
    
    flush_bold()
    return ''.join(result_parts)


def generate_essay_doc(result, output_path, title, uploaded_images=None):
    """生成作文批改Word文档
    uploaded_images: 上传的学生作业图片路径列表，用于在原文部分展示手写原图
    """
    doc = create_doc_with_title(title, f"满分：{result['full_score']}分  得分：{result['score']}分")

    # 原文
    add_para(doc, [('原文', True, BLUE)], size=14, space_before=6, space_after=4)
    
    # 在"原文"标题下方、提取文字之前插入学生手写原图
    if uploaded_images and len(uploaded_images) > 0:
        for img_path in uploaded_images:
            try:
                if os.path.exists(img_path):
                    # 添加图片，宽度设为15厘米（约为页面宽度）
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    # 计算合适的宽度，最大15cm
                    from PIL import Image as PILImage
                    with PILImage.open(img_path) as img:
                        w, h = img.size
                        max_width_cm = 15
                        ratio = h / w
                        width_cm = max_width_cm
                        height_cm = width_cm * ratio
                        # 如果高度太大，限制高度为20cm
                        if height_cm > 20:
                            height_cm = 20
                            width_cm = height_cm / ratio
                    run.add_picture(img_path, width=Cm(width_cm))
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(6)
            except Exception as e:
                # 图片插入失败不影响整体
                pass
    
    # 提取的文字内容（在图片之后）— 带错误标注
    student_text = result.get('student_text', '（待提取）')
    annotated_runs = build_annotated_student_text_runs(student_text, result)
    # 转为add_para的格式 [(text, bold, color), ...]
    para_runs = [(text, bold, color) for text, bold, color in annotated_runs]
    add_para(doc, para_runs, size=12, indent=24)

    # 批改
    add_para(doc, [('批改', True, BLUE)], size=14, space_before=10, space_after=4)
    
    # 支持两种格式：旧格式（errors列表）和新格式（error_categories按类型分类）
    if 'error_categories' in result and result['error_categories']:
        # 新格式：按类型分类
        for cat in result['error_categories']:
            category_name = cat.get('category', '错误')
            errors = cat.get('errors', [])
            for err in errors:
                add_para(doc, [
                    (f"【{category_name}】", True, RED),
                    (f"错误：{err.get('error_text', '')}", False, BLACK),
                ], size=12, space_before=4, space_after=2)
                add_para(doc, [('分析：' + err.get('analysis', ''), False, BLACK)], size=11, space_after=2, indent=24)
                add_para(doc, [('建议：' + err.get('suggestion', ''), False, GREEN)], size=11, space_after=4, indent=24)
    else:
        # 旧格式：平铺的错误列表
        for err in result.get('errors', []):
            add_para(doc, [
                (f"【{err['type']}】", True, RED),
                (f"错误：{err['error']}", False, BLACK),
            ], size=12, space_before=4, space_after=2)
            add_para(doc, [('分析：' + err['analysis'], False, BLACK)], size=11, space_after=2, indent=24)
            add_para(doc, [('建议：' + err['suggestion'], False, GREEN)], size=11, space_after=4, indent=24)

    # 总评
    add_para(doc, [('总评', True, BLUE)], size=14, space_before=10, space_after=4)
    add_para(doc, [('表扬：', False, GREEN), (result['praise'], False, BLACK)], size=12, space_after=4)
    
    # 支持两种评分格式：多维度评分或三要素评分
    if 'scoring_dimensions' in result and result['scoring_dimensions']:
        # 新格式：多维度评分
        for i, dim in enumerate(result['scoring_dimensions'], 1):
            dim_name = dim.get('name', '')
            dim_score = dim.get('score', 0)
            dim_full = dim.get('full_score', 0)
            dim_eval = dim.get('evaluation', '')
            score_text = f"（{dim_score}/{dim_full}）" if dim_full else ''
            eval_text = dim_eval or result.get(f'{dim_name.lower()}_eval', '')
            if not eval_text:
                if dim_name == '内容':
                    eval_text = result.get('content_eval', '')
                elif dim_name == '语言':
                    eval_text = result.get('language_eval', '')
                elif '组织' in dim_name or '结构' in dim_name:
                    eval_text = result.get('org_eval', '')
            add_para(doc, [(f"{i}. {dim_name}{score_text}：{eval_text}", False, BLACK)], size=12, space_after=4, indent=24)
    else:
        # 旧格式：三要素
        add_para(doc, [('1. 内容：' + result['content_eval'], False, BLACK)], size=12, space_after=4, indent=24)
        add_para(doc, [('2. 语言：' + result['language_eval'], False, BLACK)], size=12, space_after=4, indent=24)
        add_para(doc, [('3. 组织结构：' + result['org_eval'], False, BLACK)], size=12, space_after=4, indent=24)

    # 范文 — 修改部分用黑色粗体标出
    add_para(doc, [('范文参考', True, BLUE)], size=14, space_before=10, space_after=4)
    model_essay = result.get('model_essay', '')
    if model_essay:
        # 如果范文中没有**粗体标记，自动对比原文找出修改处并加粗
        student_text = result.get('student_text', '')
        if '**' not in model_essay and student_text:
            model_essay = diff_and_bold_model_essay(student_text, model_essay)
        
        # 按段落分割
        paragraphs = model_essay.split('\n')
        for para_text in paragraphs:
            if para_text.strip():
                # 解析**粗体**标记
                bold_parts = parse_bold_markdown(para_text)
                runs = [(text, bold, BLACK) for text, bold in bold_parts]
                add_para(doc, runs, size=12, indent=24)
            else:
                # 空行
                add_para(doc, [('', False, BLACK)], size=6, indent=24)

    doc.save(output_path)


def generate_objective_doc(result, output_path, title, uploaded_images=None):
    """生成客观题批改Word文档
    包含：学生原文 + 错题分析
    uploaded_images: 上传的学生作业图片路径列表，用于在原文部分展示手写原图
    """
    score = result.get('score', 0)
    full_score = result.get('full_score', 100)
    doc = create_doc_with_title(title, f"满分：{full_score}分  得分：{score}分")
    
    # 得分概览
    total = result.get('total_questions', 0)
    correct = result.get('correct_count', 0)
    wrong = result.get('wrong_count', 0)
    score = result.get('score', 0)
    full_score = result.get('full_score', 100)
    
    add_para(doc, [('得分概览', True, BLUE)], size=14, space_before=6, space_after=4)
    
    overview_text = f'总题数：{total}题    做对：{correct}题    做错：{wrong}题    得分：{score}/{full_score}'
    if total > 0:
        accuracy = round(correct / total * 100, 1)
        overview_text += f'    正确率：{accuracy}%'
    add_para(doc, [(overview_text, False, BLACK)], size=12, space_after=6)
    
    # 原文
    add_para(doc, [('原文', True, BLUE)], size=14, space_before=10, space_after=4)
    
    # OCR识别的文字内容
    student_text = result.get('student_text', '（未识别到学生作业内容）')
    if student_text:
        # 按行显示
        lines = student_text.split('\n')
        for line in lines:
            if line.strip():
                add_para(doc, [(line.strip(), False, BLACK)], size=12, indent=24)
            else:
                add_para(doc, [('', False, BLACK)], size=6, indent=24)
    
    # 错题分析
    add_para(doc, [('错题分析', True, BLUE)], size=14, space_before=10, space_after=4)
    
    wrong_questions = result.get('wrong_questions', [])
    if wrong_questions:
        for i, q in enumerate(wrong_questions, 1):
            analysis = q.get('error_analysis', '')
            
            # 直接输出完整的错题分析内容（已由AI按模板格式生成）
            if analysis:
                analysis_lines = analysis.split('\n')
                for line in analysis_lines:
                    stripped = line.strip()
                    if stripped:
                        add_para(doc, [(stripped, False, BLACK)], size=12, indent=24, space_before=2, space_after=2)
                    else:
                        add_para(doc, [('', False, BLACK)], size=6, indent=24)
                # 每道题之间加一点间距
                add_para(doc, [('', False, BLACK)], size=6, indent=24)
    else:
        add_para(doc, [('🎉 恭喜！所有题目都做对了，没有错题。', True, GREEN)], size=12, indent=24)
    
    # 整体评价
    overall_comment = result.get('overall_comment', '')
    if overall_comment:
        add_para(doc, [('整体评价', True, BLUE)], size=14, space_before=10, space_after=4)
        comment_lines = overall_comment.split('\n')
        for line in comment_lines:
            if line.strip():
                add_para(doc, [(line.strip(), False, BLACK)], size=12, indent=24)
    
    doc.save(output_path)


# ==================== 路由 ====================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download/<filename>')
def download(filename):
    """下载批改结果"""
    return send_file(
        os.path.join(app.config['OUTPUT_FOLDER'], filename),
        as_attachment=True,
        download_name=filename
    )


# ==================== AI配置API ====================

@app.route('/api/config', methods=['GET'])
def api_config_get():
    """获取系统配置（不含API Key明文）"""
    config = load_config()
    # 隐藏API Key，只显示是否已配置
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
    """获取批改历史记录"""
    try:
        history = load_history()
        # 过滤掉文件已不存在的记录
        valid_history = []
        for record in history:
            file_path = os.path.join(app.config['OUTPUT_FOLDER'], record['filename'])
            if os.path.exists(file_path):
                valid_history.append(record)
        # 如果有文件被删除了，更新历史记录
        if len(valid_history) != len(history):
            save_history(valid_history)
        return jsonify({'success': True, 'history': valid_history})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/history/<record_id>', methods=['DELETE'])
def api_history_delete(record_id):
    """删除历史记录（同时删除文件）"""
    try:
        history = load_history()
        record = next((r for r in history if r['id'] == record_id), None)
        if record:
            # 删除文件
            file_path = os.path.join(app.config['OUTPUT_FOLDER'], record['filename'])
            if os.path.exists(file_path):
                os.remove(file_path)
            # 删除记录
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
                'has_template_file': is_overridden and overrides[key].get('template_file') is not None,
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
        
        # 先查找内置模板（可能有用户上传的覆盖文件）
        for key, t in BUILTIN_TEMPLATES.items():
            if t['id'] == template_id:
                template_name = t['name']
                if key in overrides:
                    template_file = overrides[key].get('template_file')
                break
        
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
        
        template_path = os.path.join(app.config['TEMPLATE_FOLDER'], template_file)
        if not os.path.exists(template_path):
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
            # PDF类型，将所有页转为图片预览
            try:
                import base64
                import io
                from pdf2image import convert_from_path
                images = convert_from_path(template_path)
                if images:
                    pages = []
                    for img in images:
                        img_buffer = io.BytesIO()
                        img.save(img_buffer, format='JPEG')
                        page_data = base64.b64encode(img_buffer.getvalue()).decode('utf-8')
                        pages.append(page_data)
                    # 如果只有一页，保持向后兼容
                    if len(pages) == 1:
                        content = pages[0]
                    else:
                        # 多页情况，用特殊标记分隔
                        content = '|||'.join(pages)
                    file_type = 'pdf_image'
                else:
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
                import io
                from pdf2image import convert_from_path
                images = convert_from_path(file_path)
                if images:
                    pages = []
                    for img in images:
                        img_buffer = io.BytesIO()
                        img.save(img_buffer, format='JPEG')
                        page_data = base64.b64encode(img_buffer.getvalue()).decode('utf-8')
                        pages.append(page_data)
                    if len(pages) == 1:
                        content = pages[0]
                    else:
                        content = '|||'.join(pages)
                    file_type = 'pdf_image'
                else:
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
        import time
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
            
            # 如果有PDF，需要转图片
            pdf_files = [
                f for f in uploaded_student_files
                if os.path.splitext(f)[1].lower() == '.pdf'
            ]
            if pdf_files:
                try:
                    from pdf2image import convert_from_path
                    for pdf_file in pdf_files:
                        images = convert_from_path(pdf_file)
                        for i, img in enumerate(images):
                            img_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_pdf_{i}.jpg")
                            img.save(img_path, 'JPEG')
                            student_images.append(img_path)
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
            
            # 生成Word文档
            today = datetime.date.today()
            date_str = f"{today.month}.{today.day}"
            project_name = '客观题'
            base_filename = f"{date_str}-{project_name}批改"
            
            output_filename = f"{base_filename}.docx"
            counter = 1
            while os.path.exists(os.path.join(app.config['OUTPUT_FOLDER'], output_filename)):
                output_filename = f"{base_filename}({counter}).docx"
                counter += 1
            
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            generate_objective_doc(result, output_path, title, student_images)
            
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
                'download_url': f'/download/{output_filename}',
                'filename': output_filename,
                'score': result.get('score', 0),
                'full_score': result.get('full_score', 100),
                'ai_used': True,
                'rules_applied': len(custom_rules),
                'rule_names': [r['name'] for r in custom_rules],
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
        
        # 获取评分标准内容
        grading_standard = None
        if grade_type and grade_type in BUILTIN_GRADING_STANDARDS:
            grading_standard = get_grading_standard_content(grade_type)
        
        # 使用AI智能批改
        ai_result = grade_with_doubao(
            student_input, template,
            image_paths=uploaded_images,
            api_key=api_key,
            model_name=model_name,
            custom_rules=custom_rules,
            grading_standard=grading_standard
        )

        if not ai_result or not ai_result.get('success'):
            error_msg = ai_result.get('error', 'AI返回为空') if ai_result else 'AI调用失败'
            return jsonify({'success': False, 'error': f'AI批改失败：{error_msg}'})

        result = ai_result
        result['ai_used'] = True

        # 设置学生原文（优先使用AI识别的结果）
        if not result.get('student_text'):
            result['student_text'] = student_input or '（已上传图片，文字已由AI识别提取）'

        # 生成Word文档
        # 文件名格式：日期-批改项目，如 "9.9-中考作文批改.docx"
        today = datetime.date.today()
        date_str = f"{today.month}.{today.day}"
        
        # 简化标题（去掉"批改"和"模板"字样）
        project_name = title.replace('批改', '').replace('模板', '').strip()
        if not project_name:
            project_name = '作业'
        
        # 基础文件名
        base_filename = f"{date_str}-{project_name}批改"
        
        # 检查是否重名，避免覆盖
        output_filename = f"{base_filename}.docx"
        counter = 1
        while os.path.exists(os.path.join(app.config['OUTPUT_FOLDER'], output_filename)):
            output_filename = f"{base_filename}({counter}).docx"
            counter += 1
        
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        generate_essay_doc(result, output_path, title, uploaded_images)

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
            'download_url': f'/download/{output_filename}',
            'filename': output_filename,
            'score': result['score'],
            'full_score': result['full_score'],
            'ai_used': True,
            'rules_applied': len(custom_rules),
            'rule_names': [r['name'] for r in custom_rules],
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


if __name__ == '__main__':
    print("=" * 50)
    print("  英语作业自动批改工具")
    print("  访问地址：http://localhost:5005")
    print("=" * 50)
    app.run(host='0.0.0.0', debug=True, port=5005)
