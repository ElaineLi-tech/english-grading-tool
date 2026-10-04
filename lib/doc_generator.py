"""
Word 文档生成模块
包含：作文批改文档、客观题批改文档、概要写作讲义等
"""

import os
import re
from io import BytesIO

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ==================== 常量定义 ====================

BLACK = RGBColor(0, 0, 0)
BLUE = RGBColor(0x00, 0x70, 0xC0)
RED = RGBColor(0xFF, 0x00, 0x00)
GRAY = RGBColor(0x80, 0x80, 0x80)
GREEN = RGBColor(0x00, 0x80, 0x00)
ORANGE = RGBColor(0xFF, 0x8C, 0x00)
FONT_CN = '微软雅黑'
FONT_EN = 'Times New Roman'


# ==================== 文档保存辅助 ====================

def _save_doc(doc, output_path):
    """保存文档：output_path 为 None 时返回 BytesIO，否则保存到文件并返回路径"""
    if output_path is None:
        # 流式返回：写入内存
        bio = BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio
    else:
        # 保存到文件
        doc.save(output_path)
        return output_path


# ==================== 字体与段落辅助函数 ====================

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


# ==================== 文本处理辅助函数 ====================

def extract_english_correction(suggestion):
    """从AI的suggestion中提取纯英文的正确答案
    去掉中文解释、"改为："等前缀，只保留英文内容
    """
    if not suggestion:
        return ''

    text = suggestion.strip()

    # 去掉开头的"改为："、"答案："等中文前缀
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


# ==================== 作文批改文档生成 ====================

def generate_essay_doc(result, output_path, title, uploaded_images=None, template=None):
    """生成作文批改Word文档
    uploaded_images: 上传的学生作业图片路径列表，用于在原文部分展示手写原图
    template: 批改模板，用于确定文档格式风格
    """
    doc = create_doc_with_title(title, f"满分：{result['full_score']}分  得分：{result['score']}分")

    # 获取模板风格
    style = 'default'
    has_model_essay = True
    has_good_sentences = False
    if template:
        style = template.get('style', 'default')
        has_model_essay = template.get('has_model_essay', True)
        has_good_sentences = template.get('has_good_sentences', False)

    # 高考风格：题目要求
    if style == 'gaokao':
        topic_text = result.get('topic', '')
        if topic_text:
            add_para(doc, [('题目要求', True, BLUE)], size=14, space_before=6, space_after=4)
            for line in topic_text.split('\n'):
                if line.strip():
                    add_para(doc, [(line.strip(), False, BLACK)], size=12, indent=24, space_after=2)

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

    # 高考风格：改进建议
    if style == 'gaokao':
        improvement = result.get('improvement_suggestion', '')
        if not improvement:
            # 从维度评价中提取改进建议
            dim_evals = []
            if 'scoring_dimensions' in result:
                for dim in result['scoring_dimensions']:
                    ev = dim.get('evaluation', '')
                    if ev:
                        dim_evals.append(ev)
            if dim_evals:
                improvement = '继续加强语言准确性，注意拼写、语法和搭配；丰富句式结构，提升表达的多样性。'
        if improvement:
            add_para(doc, [('改进建议：', False, RED), (improvement, False, BLACK)], size=12, space_after=4, indent=24)

    # 范文 — 修改部分用黑色粗体标出
    if has_model_essay:
        # 高考风格：范文标题不同
        if style == 'gaokao':
            model_title = '基于原文的修改范文'
        else:
            model_title = '范文参考'
        add_para(doc, [(model_title, True, BLUE)], size=14, space_before=10, space_after=4)
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

        # 好词好句（仅中考等模板有）
        if has_good_sentences:
            good_sentences = result.get('good_sentences', [])
            if good_sentences and len(good_sentences) > 0:
                add_para(doc, [('好词好句学习', True, BLUE)], size=14, space_before=10, space_after=4)
                for i, gs in enumerate(good_sentences, 1):
                    add_para(doc, [(f"{i}. {gs}", False, BLACK)], size=12, space_after=2, indent=24)

    return _save_doc(doc, output_path)


# ==================== 客观题批改文档生成 ====================

def generate_objective_doc(result, output_path, title, uploaded_images=None, template=None):
    """生成客观题批改Word文档
    包含：学生原文 + 错题分析
    uploaded_images: 上传的学生作业图片路径列表，用于在原文部分展示手写原图
    template: 批改模板，用于根据不同题型调整文档格式
    """
    # 获取模板风格
    style = 'objective'
    if template:
        style = template.get('style', 'objective')

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

    # 如果是客观题子题型且result中没有student_text但有wrong_questions_text，显示学生输入的错题信息
    objective_sub_styles = ['objective_reading', 'objective_six_choose_four', 'objective_cloze', 'objective_grammar', 'objective_cat_fishing']
    if style in objective_sub_styles:
        wrong_questions_text = result.get('wrong_questions_text', '')
        if wrong_questions_text and not result.get('student_text'):
            add_para(doc, [('学生错题信息', True, BLUE)], size=14, space_before=10, space_after=4)
            lines = wrong_questions_text.split('\n')
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

    return _save_doc(doc, output_path)


# ==================== 概要写作讲义文档生成 ====================

def generate_summary_handout_doc(handout_data, output_path, title):
    """生成概要写作讲义Word文档"""
    BLUE_LOCAL = RGBColor(0x00, 0x70, 0xC0)
    RED_LOCAL = RGBColor(0xC0, 0x00, 0x00)
    GREEN_LOCAL = RGBColor(0x00, 0x80, 0x00)
    BLACK_LOCAL = RGBColor(0, 0, 0)
    GRAY_LOCAL = RGBColor(0x66, 0x66, 0x66)
    FONT_CN_LOCAL = '微软雅黑'
    FONT_EN_LOCAL = 'Times New Roman'

    def set_font(run, size=12, bold=False, color=None, font_cn=FONT_CN_LOCAL):
        run.font.name = FONT_EN_LOCAL
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = color
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), FONT_EN_LOCAL)
        rFonts.set(qn('w:hAnsi'), FONT_EN_LOCAL)
        rFonts.set(qn('w:eastAsia'), font_cn)

    doc = Document()

    # 设置页面边距
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # ===== 标题 =====
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(handout_data.get('title', title))
    set_font(run, 16, True, BLUE_LOCAL)
    p.paragraph_format.space_after = Pt(12)

    # ===== Directions =====
    if handout_data.get('directions'):
        p = doc.add_paragraph()
        run = p.add_run('Directions: ')
        set_font(run, 11, True, BLACK_LOCAL)
        run2 = p.add_run(handout_data['directions'])
        set_font(run2, 11, False, BLACK_LOCAL)
        p.paragraph_format.space_after = Pt(8)

    # ===== 原文 =====
    p = doc.add_paragraph()
    run = p.add_run('原文')
    set_font(run, 14, True, BLUE_LOCAL)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(6)

    # 英文原文
    eng_text = handout_data.get('passage_english', '')
    eng_paragraphs = eng_text.split('\n')
    for para in eng_paragraphs:
        if para.strip():
            p = doc.add_paragraph()
            run = p.add_run(para.strip())
            set_font(run, 11, False, BLACK_LOCAL)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.5

    # 中文翻译
    ch_text = handout_data.get('passage_chinese', '')
    if ch_text:
        doc.add_paragraph()  # 空行
        ch_paragraphs = ch_text.split('\n')
        for para in ch_paragraphs:
            if para.strip():
                p = doc.add_paragraph()
                run = p.add_run(para.strip())
                set_font(run, 10, False, GRAY_LOCAL)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.3

    # ===== 一、语篇模式 =====
    p = doc.add_paragraph()
    run = p.add_run('一、语篇模式')
    set_font(run, 14, True, BLUE_LOCAL)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)

    pattern_name = handout_data.get('pattern_name', '')
    pattern_eng = handout_data.get('pattern_english_name', '')
    if pattern_name:
        p = doc.add_paragraph()
        run = p.add_run(f'本篇属于{pattern_name}')
        if pattern_eng:
            run2 = p.add_run(f'（{pattern_eng}）')
            set_font(run2, 11, False, BLACK_LOCAL)
        set_font(run, 12, True, RED_LOCAL)
        p.paragraph_format.space_after = Pt(4)

    if handout_data.get('pattern_definition'):
        p = doc.add_paragraph()
        run = p.add_run('定义：')
        set_font(run, 11, True, BLACK_LOCAL)
        run2 = p.add_run(handout_data['pattern_definition'])
        set_font(run2, 11, False, BLACK_LOCAL)
        p.paragraph_format.space_after = Pt(6)

    # 结构分析表格
    structure = handout_data.get('structure_analysis', [])
    if structure:
        p = doc.add_paragraph()
        run = p.add_run('结构分析')
        set_font(run, 12, True, BLACK_LOCAL)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(4)

        table = doc.add_table(rows=1, cols=3)
        table.style = 'Light Grid Accent 1'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # 表头
        hdr_cells = table.rows[0].cells
        headers = ['结构环节', '原文对应', '核心内容']
        for i, h in enumerate(headers):
            hdr_cells[i].text = h
            for para in hdr_cells[i].paragraphs:
                for run in para.runs:
                    set_font(run, 10, True, BLACK_LOCAL)
            # 表头底纹
            tc = hdr_cells[i]._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), 'D9E2F3')
            tcPr.append(shd)

        # 设置列宽
        for row in table.rows:
            row.cells[0].width = Cm(3)
            row.cells[1].width = Cm(2.5)
            row.cells[2].width = Cm(10)

        for item in structure:
            row_cells = table.add_row().cells
            row_cells[0].text = item.get('section', '')
            row_cells[1].text = item.get('paragraph', '')
            row_cells[2].text = item.get('content', '')
            for cell in row_cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        set_font(run, 10, False, BLACK_LOCAL)

        doc.add_paragraph()  # 空行

    # ===== 二、"3+2"得分 =====
    p = doc.add_paragraph()
    run = p.add_run('二、"3+2"得分')
    set_font(run, 14, True, BLUE_LOCAL)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)

    # 3个大点
    p = doc.add_paragraph()
    run = p.add_run('3 个大点')
    set_font(run, 12, True, RED_LOCAL)
    p.paragraph_format.space_after = Pt(4)

    main_points = handout_data.get('three_main_points', [])
    for i, mp in enumerate(main_points, 1):
        # 大点英文
        p = doc.add_paragraph()
        run = p.add_run(f'大点{"①②③④⑤"[i-1]}：')
        set_font(run, 11, True, BLACK_LOCAL)
        run2 = p.add_run(mp.get('point', ''))
        set_font(run2, 11, False, BLACK_LOCAL)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_after = Pt(2)

        # 大点解释
        if mp.get('explanation'):
            p = doc.add_paragraph()
            run = p.add_run(mp['explanation'])
            set_font(run, 10, False, GRAY_LOCAL)
            p.paragraph_format.left_indent = Cm(1)
            p.paragraph_format.space_after = Pt(4)

    # 2个小点
    p = doc.add_paragraph()
    run = p.add_run('2 个小点')
    set_font(run, 12, True, GREEN_LOCAL)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)

    minor_points = handout_data.get('two_minor_points', [])
    for i, mp in enumerate(minor_points, 1):
        p = doc.add_paragraph()
        run = p.add_run(f'小点{i}：')
        set_font(run, 11, True, BLACK_LOCAL)
        run2 = p.add_run(mp.get('point', ''))
        set_font(run2, 11, False, BLACK_LOCAL)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_after = Pt(2)

        if mp.get('explanation'):
            p = doc.add_paragraph()
            cor = mp.get('corresponds_to', '')
            prefix = f'（{cor}）' if cor else ''
            run = p.add_run(prefix + mp['explanation'])
            set_font(run, 10, False, GRAY_LOCAL)
            p.paragraph_format.left_indent = Cm(1)
            p.paragraph_format.space_after = Pt(4)

    # ===== 做题方法 =====
    p = doc.add_paragraph()
    run = p.add_run('做题方法')
    set_font(run, 14, True, BLUE_LOCAL)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)

    steps = handout_data.get('method_steps', [])
    for step in steps:
        p = doc.add_paragraph()
        run = p.add_run(step)
        set_font(run, 11, False, BLACK_LOCAL)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_after = Pt(3)

    # ===== 答案 =====
    p = doc.add_paragraph()
    run = p.add_run('答案：')
    set_font(run, 14, True, BLUE_LOCAL)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)

    if handout_data.get('sample_answer_english'):
        p = doc.add_paragraph()
        run = p.add_run(handout_data['sample_answer_english'])
        set_font(run, 11, False, BLACK_LOCAL)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5

    if handout_data.get('sample_answer_chinese'):
        p = doc.add_paragraph()
        run = p.add_run(handout_data['sample_answer_chinese'])
        set_font(run, 10, False, GRAY_LOCAL)
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.3

    return _save_doc(doc, output_path)
