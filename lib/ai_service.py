"""
AI 调用相关服务
包含：豆包 API 调用、Prompt 构建、文本提取等
"""

import os
import json
import re
import base64
import requests

from docx import Document


# ==================== 模板文件路径辅助函数 ====================

def get_builtin_template_file_path(template_file, builtin_dir=''):
    """获取内置模板文件的完整路径"""
    if not template_file:
        return None
    if builtin_dir:
        path = os.path.join(builtin_dir, template_file)
        if os.path.exists(path):
            return path
    return None


def get_template_file_path(template_file, template_folder='', builtin_dir='', is_overridden=False):
    """获取模板文件的完整路径（优先用户覆盖版本，其次内置版本）"""
    if not template_file:
        return None
    # 用户覆盖的文件在 template_folder
    if is_overridden:
        user_path = os.path.join(template_folder, template_file)
        if os.path.exists(user_path):
            return user_path
    # 内置模板文件在 builtin_dir
    builtin_path = get_builtin_template_file_path(template_file, builtin_dir)
    if builtin_path:
        return builtin_path
    # 兜底：在 template_folder 中找
    user_path = os.path.join(template_folder, template_file)
    if os.path.exists(user_path):
        return user_path
    return None


# ==================== 豆包API调用 ====================

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


# ==================== Prompt 构建 ====================

def build_grading_prompt(student_text, template, has_image=False, custom_rules=None,
                         grading_standard=None, template_folder=None,
                         builtin_template_files_dir='', reference_documents=''):
    """构建批改prompt，让AI按指定模板格式返回JSON
    custom_rules: 用户自定义规则列表，每条规则包含name和content
    grading_standard: 评分标准文本内容，用于指导AI评分
    template_folder: 模板文件夹路径，用于读取模板文件内容
    builtin_template_files_dir: 内置模板文件目录
    reference_documents: 用户上传的docx/pdf等参考文档内容（仅作参考，不是学生答案）
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

    # 构建格式模板部分
    format_template_section = ''
    template_file = template.get('template_file')
    is_overridden = template.get('is_overridden', False)
    if template_file and template_folder:
        template_path = get_template_file_path(template_file, template_folder, builtin_template_files_dir, is_overridden)
        if template_path and os.path.exists(template_path):
            ext = os.path.splitext(template_path)[1].lower()
            if ext in {'.docx', '.doc'}:
                template_content = extract_text_from_docx(template_path)
                if template_content:
                    # 限制长度
                    if len(template_content) > 2000:
                        template_content = template_content[:2000] + '\n...（内容过长，已截取核心部分）'
                    format_template_section = f"\n【批改格式参考模板】\n以下是批改结果的格式模板范例，请严格按照这个模板的结构、用语习惯和排版方式来组织你的批改结果：\n{template_content}\n"

    # 根据模板风格调整输出要求
    style = template.get('style', 'default')

    # 高考风格：需要改进建议
    improvement_field = ''
    if style == 'gaokao':
        improvement_field = '  "improvement_suggestion": "综合改进建议：针对学生的问题，给出具体的、可操作的改进建议（如错题本、练习重点等）",\n'

    # 范文字段说明
    if style == 'gaokao':
        model_essay_desc = '"基于原文的修改范文（在学生原文基础上改写提升，所有修改/提升的部分请用**两个星号**包裹标记，例如：**improved** 表示此处有修改。结构更完整、用词更丰富、句式更多样）"'
        good_sentences_field = ''
    else:
        model_essay_desc = '"参考范文（在学生原文基础上改写提升，所有修改/提升的部分请用**两个星号**包裹标记，例如：**improved** 表示此处有修改。结构更完整、用词更丰富、句式更多样）\\n\\n【好词好句学习】\\n1. 好词/好句型 - 解释\\n2. ..."'
        good_sentences_field = '  "good_sentences": [\n    "好词1 - 释义",\n    "好词2 - 释义"\n  ],\n'

    # 构建参考文档部分
    reference_section = ''
    if reference_documents and reference_documents.strip():
        ref_text = reference_documents.strip()
        if len(ref_text) > 3000:
            ref_text = ref_text[:3000] + '\n...（内容过长，已截取核心部分）'
        reference_section = f"\n【参考文档】\n以下是用户上传的参考资料，仅用于辅助批改（如参考答案、题目原文等），不是学生答案：\n{ref_text}\n"

    prompt = f"""你是一位专业的英语作文批改老师，请按照以下要求批改学生作文。

【批改模板】
模板名称：{template['name']}
模板说明：{template.get('description', '英语作文批改')}
满分：{template.get('full_score', 20)}分
评分维度：
{dims_desc}
{format_template_section}
{rules_section}
{standard_section}
{reference_section}
⚠️ 重要规则：学生答案只来自【学生作文】部分（文本或图片识别内容），【参考文档】【批改格式参考模板】中的内容都不是学生答案，仅作批改参考使用。

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
{improvement_field}{good_sentences_field}  "model_essay": {model_essay_desc}
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


def build_summary_handout_prompt(passage_text, template, scoring_standard_text='',
                                 custom_rules=None, template_folder=None,
                                 builtin_template_files_dir='', format_reference_text=''):
    """构建概要写作讲义生成的prompt
    format_reference_text: 用户上传的docx/pdf文档内容（优先级高于内置模板，作为格式参考）
    """
    template_name = template.get('name', '概要写作讲义')

    # 格式参考内容：优先使用用户上传的文档，其次使用内置模板
    format_template_section = ''
    ref_content = ''
    ref_source = ''

    if format_reference_text and format_reference_text.strip():
        ref_content = format_reference_text.strip()
        ref_source = '用户上传的参考文档'
    else:
        # 读取内置模板文件内容作为格式参考
        template_file = template.get('template_file')
        is_overridden = template.get('is_overridden', False)
        if template_file and template_folder:
            template_path = get_template_file_path(template_file, template_folder, builtin_template_files_dir, is_overridden)
            if template_path and os.path.exists(template_path):
                ext = os.path.splitext(template_path)[1].lower()
                if ext in {'.docx', '.doc'}:
                    template_content = extract_text_from_docx(template_path)
                    if template_content:
                        ref_content = template_content
                        ref_source = '内置模板'

    if ref_content:
        if len(ref_content) > 3000:
            ref_content = ref_content[:3000] + '\n...（内容过长，已截取核心部分）'
        format_template_section = f"\n【讲义格式参考】（来自{ref_source}）\n以下是讲义的格式参考范例，仅用于参考讲义的结构、用语习惯和排版方式。\n⚠️ 重要提醒：参考文档中的原文、题目、答案等内容只是示例，你必须基于【原文】部分的真实内容来生成讲义，绝对不能使用或照搬参考文档中的示例题目和内容！\n{ref_content}\n"

    custom_rules_text = ''
    if custom_rules and len(custom_rules) > 0:
        custom_rules_text = '\n【自定义规则】\n'
        for i, rule in enumerate(custom_rules, 1):
            custom_rules_text += f"{i}. {rule.get('content', '')}\n"

    scoring_standard_section = ''
    if scoring_standard_text:
        scoring_standard_section = f"\n【3+2评分标准参考】\n以下是概要写作的3+2评分标准，请根据这个标准来确定3个大点和2个小点：\n{scoring_standard_text}\n"

    prompt = f"""你是一位专业的英语概要写作老师，请根据以下原文，生成一份完整的概要写作讲义。

【原文】
{passage_text}
{format_template_section}{scoring_standard_section}
【任务要求】
请仔细阅读【原文】部分的内容，分析其语篇结构，生成一份概要写作讲义。

⚠️ 核心规则：你必须基于用户上传的【原文】内容生成讲义，绝对不能使用【讲义格式参考模板】中的示例题目或内容。模板仅用于参考排版格式和用语风格。

讲义需要包含以下部分：

1. **标题**：包含日期和主题，如"X.X 概要写作讲义 主张-反主张型"

2. **原文**：Directions + 英文原文 + 中文翻译（逐段中英对照）

3. **一、语篇模式**
   - 判断语篇类型（如：主张-反主张型、问题-解决型、现象-分析型、原因-结果型、对比型等）
   - 给出该类型的定义
   - 结构分析表格：结构环节 / 原文对应段落 / 核心内容

4. **二、"3+2"得分**
   - 3个大点（main points）：每个大点一句话概括，说明对应原文哪几段，以及为什么是大点
   - 2个小点（supporting details）：从大点中选2个关键细节补充，说明对应哪个大点

5. **做题方法**
   - 分步骤说明这类文章的概要写作方法和技巧（4-6步）

6. **参考答案**
   - 60词以内的英文参考范文
   - 对应的中文翻译

【输出格式】
返回严格的JSON格式，结构如下：
{{
  "title": "讲义标题（如：9.26 概要写作讲义 主张-反主张型）",
  "directions": "题目要求（Read the following passage...）",
  "passage_english": "完整的英文原文",
  "passage_chinese": "完整的中文翻译，逐段对应",
  "pattern_name": "语篇模式名称（如：主张-反主张型）",
  "pattern_english_name": "英文名称（如：Claim-Counterclaim Pattern）",
  "pattern_definition": "语篇模式定义",
  "structure_analysis": [
    {{
      "section": "结构环节名称（如：Background / Claim / Counterclaim / Consensus）",
      "paragraph": "对应原文段落（如：第1段）",
      "content": "核心内容概括"
    }}
  ],
  "three_main_points": [
    {{
      "point": "大点①的英文概括",
      "explanation": "中文解释，说明对应原文哪几段，为什么是大点",
      "paragraph": "对应段落"
    }}
  ],
  "two_minor_points": [
    {{
      "point": "小点的英文概括",
      "explanation": "中文解释，说明对应哪个大点，为什么是小点",
      "corresponds_to": "对应大点①/②/③"
    }}
  ],
  "method_steps": [
    "第一步：...",
    "第二步：...",
    "第三步：..."
  ],
  "sample_answer_english": "60词以内的参考范文（英文）",
  "sample_answer_chinese": "参考范文的中文翻译"
}}

注意：
1. 3个大点必须覆盖全文核心要点，不能遗漏任何主要观点
2. 2个小点是大点的关键细节支撑，不能和大点重复
3. 参考范文必须控制在60词以内，使用自己的话表达
4. 语篇模式判断要准确，结构分析要清晰对应
5. 做题方法要针对该类型文章的特点，给出实用的技巧
6. 返回纯JSON，不要有markdown格式或其他说明文字
7. 如果提供了【讲义格式参考模板】，请严格按照模板的结构和用语风格来组织内容
{custom_rules_text}
"""
    return prompt


# ==================== 文本提取 ====================

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


def extract_text_from_file(file_path):
    """从文件中提取纯文本内容（支持docx/doc/pdf/txt）"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in {'.docx', '.doc'}:
        return extract_text_from_docx(file_path)
    elif ext == '.txt':
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except:
            return ''
    elif ext == '.pdf':
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(file_path)
            text = ''
            for page in doc:
                text += page.get_text()
            doc.close()
            return text
        except Exception as e:
            print(f"读取PDF失败: {e}")
            return ''
    return ''


# ==================== 客观题 Prompt 构建 ====================

def build_objective_grading_prompt(answer_key_text, template, custom_rules=None,
                                   template_folder=None):
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


def build_objective_text_grading_prompt(wrong_questions_text, answer_key_text, template,
                                        custom_rules=None, template_folder=None,
                                        builtin_template_files_dir=''):
    """构建基于文字输入的客观题批改prompt
    wrong_questions_text: 用户手动输入的错题信息（题号、学生答案、正确答案）
    answer_key_text: 题目解析文档内容
    template: 批改模板
    """
    template_name = template.get('name', '客观题批改')

    # 读取模板文件内容作为格式参考
    format_template_section = ''
    template_file = template.get('template_file')
    is_overridden = template.get('is_overridden', False)
    if template_file and template_folder:
        template_path = get_template_file_path(template_file, template_folder, builtin_template_files_dir, is_overridden)
        if template_path and os.path.exists(template_path):
            ext = os.path.splitext(template_path)[1].lower()
            if ext in {'.docx', '.doc'}:
                template_content = extract_text_from_docx(template_path)
                if template_content:
                    if len(template_content) > 2000:
                        template_content = template_content[:2000] + '\n...（内容过长，已截取核心部分）'
                    format_template_section = f"\n【批改格式参考模板】\n以下是批改结果的格式模板范例，请严格按照这个模板的结构、用语习惯和排版方式来组织你的批改结果：\n{template_content}\n"

    custom_rules_text = ''
    if custom_rules and len(custom_rules) > 0:
        custom_rules_text = '\n【自定义批改规则】\n'
        for i, rule in enumerate(custom_rules, 1):
            custom_rules_text += f"{i}. {rule.get('content', '')}\n"

    prompt = f"""你是一个专业的英语老师，负责批改学生的{template_name}作业。

【批改模板】
{template_name}
{format_template_section}
【学生错题信息】
以下是学生做错的题目（包含题号、学生答案、正确答案）：
{wrong_questions_text}

【题目解析参考】
{answer_key_text}

【批改要求】
1. 根据学生错题信息和题目解析，为每道错题生成详细的错题分析
2. 错题分析要包括：错误原因分析、正确答案解读、相关知识点讲解、解题思路和方法
3. 如果提供了【批改格式参考模板】，请严格按照模板的格式、结构、用语习惯来组织批改结果
4. 统计总题数、做对题数、做错题数、得分率
5. 给出整体评价和针对性建议
6. 返回严格的JSON格式，结构如下：
{{
  "total_questions": 总题数,
  "correct_count": 做对题数,
  "wrong_count": 做错题数,
  "score": 得分（百分制）,
  "full_score": 100,
  "wrong_questions": [
    {{
      "question_number": "题号",
      "student_answer": "学生答案",
      "correct_answer": "正确答案",
      "error_analysis": "完整的错题分析，严格按照模板格式，包含错误原因、解析、知识点、解题思路等",
      "knowledge_point": "知识点"
    }}
  ],
  "overall_comment": "整体评价和建议"
}}
7. 返回纯JSON，不要有markdown格式或其他说明文字
{custom_rules_text}
"""
    return prompt


# ==================== 客观题批改 ====================

def grade_objective_with_doubao(student_images, answer_key_text, template,
                                api_key='', model_name='doubao-seed-2-0-mini-260428',
                                custom_rules=None, template_folder=None):
    """使用豆包AI进行客观题批改
    student_images: 学生作业图片路径列表
    answer_key_text: 题目解析文本内容
    template: 批改模板
    template_folder: 模板文件夹路径
    """
    if not api_key:
        return None

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
    prompt = build_objective_grading_prompt(answer_key_text, template,
                                            custom_rules=custom_rules,
                                            template_folder=template_folder)

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
        json_match = re.search(r'\{[\s\S]*\}', output_text)
        if json_match:
            result_data = json.loads(json_match.group(0))
        else:
            return {'success': False, 'error': f'AI返回格式错误：{output_text[:200]}'}

    result_data['success'] = True
    return result_data


def grade_objective_text_with_doubao(wrong_questions_text, answer_key_text, template,
                                     api_key='', model_name='doubao-seed-2-0-mini-260428',
                                     custom_rules=None, template_folder=None,
                                     builtin_template_files_dir=''):
    """使用豆包AI进行文字输入的客观题批改
    wrong_questions_text: 用户输入的错题信息
    answer_key_text: 题目解析文本
    """
    if not api_key:
        return None

    prompt = build_objective_text_grading_prompt(
        wrong_questions_text, answer_key_text, template,
        custom_rules=custom_rules, template_folder=template_folder,
        builtin_template_files_dir=builtin_template_files_dir
    )

    try:
        # 使用Responses API格式
        url = 'https://ark.cn-beijing.volces.com/api/v3/responses'
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }

        content = [{'type': 'input_text', 'text': prompt}]

        payload = {
            'model': model_name,
            'input': content,
        }

        response = requests.post(url, json=payload, headers=headers, timeout=180)
        response.raise_for_status()
        result = response.json()

        # 解析结果
        output_text = ''
        if 'output' in result:
            for item in result['output']:
                if item.get('type') == 'message' and 'content' in item:
                    for c in item['content']:
                        if c.get('type') == 'output_text':
                            output_text += c.get('text', '')

        if not output_text:
            return {'success': False, 'error': 'AI返回为空'}

        # 尝试解析JSON
        try:
            # 清理可能的markdown标记
            cleaned = output_text.strip()
            if cleaned.startswith('```json'):
                cleaned = cleaned[7:]
            if cleaned.startswith('```'):
                cleaned = cleaned[3:]
            if cleaned.endswith('```'):
                cleaned = cleaned[:-3]
            cleaned = cleaned.strip()

            data = json.loads(cleaned)
            data['success'] = True
            return data
        except json.JSONDecodeError as e:
            return {'success': False, 'error': f'JSON解析失败: {str(e)}', 'raw_output': output_text[:500]}

    except Exception as e:
        return {'success': False, 'error': str(e)}


# ==================== 作文批改 ====================

def grade_with_doubao(student_text, template, image_paths=None, api_key='',
                      model_name='doubao-seed-2-0-mini-260428', custom_rules=None,
                      grading_standard=None, template_folder=None,
                      builtin_template_files_dir='', reference_documents=''):
    """使用豆包AI进行智能批改
    custom_rules: 用户自定义规则列表
    grading_standard: 评分标准文本内容
    template_folder: 模板文件夹路径
    builtin_template_files_dir: 内置模板文件目录
    reference_documents: 用户上传的docx/pdf等参考文档内容（仅作参考，不是学生答案）
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
    prompt = build_grading_prompt(
        prompt_text, template, has_image=bool(image_paths),
        custom_rules=custom_rules, grading_standard=grading_standard,
        template_folder=template_folder,
        builtin_template_files_dir=builtin_template_files_dir,
        reference_documents=reference_documents
    )
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
        content_text = result['content']
        # 尝试提取JSON（去掉可能的markdown代码块标记）
        content_text = content_text.strip()
        if content_text.startswith('```json'):
            content_text = content_text[7:]
        if content_text.startswith('```'):
            content_text = content_text[3:]
        if content_text.endswith('```'):
            content_text = content_text[:-3]
        content_text = content_text.strip()

        ai_result = json.loads(content_text)

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


# ==================== 概要写作讲义生成 ====================

def generate_summary_handout_with_doubao(passage_text, template, scoring_standard_text='',
                                         api_key='', model_name='doubao-seed-2-0-mini-260428',
                                         custom_rules=None, template_folder=None,
                                         builtin_template_files_dir='', image_paths=None,
                                         format_reference_text=''):
    """使用豆包AI生成概要写作讲义
    image_paths: 图片文件路径列表（作为原文内容，AI会自动识别）
    format_reference_text: 用户上传的docx/pdf文档内容（作为格式参考模板）
    """
    if not api_key:
        return None

    prompt = build_summary_handout_prompt(
        passage_text, template,
        scoring_standard_text=scoring_standard_text,
        custom_rules=custom_rules,
        template_folder=template_folder,
        builtin_template_files_dir=builtin_template_files_dir,
        format_reference_text=format_reference_text
    )

    # 构建 input 列表（Responses API格式）
    content = []

    # 如果有图片，添加所有图片（作为原文）
    if image_paths and len(image_paths) > 0:
        for img_path in image_paths:
            img_base64 = encode_image_to_base64(img_path)
            content.append({
                "type": "input_image",
                "image_url": f"data:image/jpeg;base64,{img_base64}"
            })

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

    # 尝试解析JSON
    try:
        cleaned = output_text.strip()
        if cleaned.startswith('```json'):
            cleaned = cleaned[7:]
        if cleaned.startswith('```'):
            cleaned = cleaned[3:]
        if cleaned.endswith('```'):
            cleaned = cleaned[:-3]
        cleaned = cleaned.strip()

        data = json.loads(cleaned)
        data['success'] = True
        return data
    except json.JSONDecodeError as e:
        return {'success': False, 'error': f'JSON解析失败: {str(e)}', 'raw_output': output_text[:500]}
