"""
内置批改模板配置常量
"""

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
        'style': 'gaokao',
        'template_file': 'gaokao_essay_template.docx',
        'has_model_essay': True,
        'has_good_sentences': True
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
    'summary_handout': {
        'id': 'builtin_summary_handout',
        'name': '概要写作讲义',
        'description': '概要写作讲义生成，包含语篇模式分析、3+2得分点拆解、做题方法和参考答案。',
        'full_score': 10,
        'scoring_dimensions': [
            {'name': '内容要点', 'full_score': 5},
            {'name': '语言表达', 'full_score': 3},
            {'name': '组织结构', 'full_score': 2}
        ],
        'style': 'summary_handout',
        'template_file': 'summary_handout_template.docx'
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
    },
    'objective_reading': {
        'id': 'builtin_obj_reading',
        'name': '阅读理解',
        'description': '阅读理解题批改，生成错题分析和解题思路。',
        'full_score': 100,
        'scoring_dimensions': [{'name': '正确率', 'full_score': 100}],
        'style': 'objective_reading',
        'has_error_analysis': True,
        'template_file': 'objective_reading_template.docx'
    },
    'objective_six_choose_four': {
        'id': 'builtin_obj_six_choose_four',
        'name': '六选四',
        'description': '六选四题型批改，分析句子衔接和逻辑关系。',
        'full_score': 100,
        'scoring_dimensions': [{'name': '正确率', 'full_score': 100}],
        'style': 'objective_six_choose_four',
        'has_error_analysis': True,
        'template_file': 'objective_six_choose_four_template.docx'
    },
    'objective_cloze': {
        'id': 'builtin_obj_cloze',
        'name': '完形填空',
        'description': '完形填空批改，分析词汇、语法和语境理解。',
        'full_score': 100,
        'scoring_dimensions': [{'name': '正确率', 'full_score': 100}],
        'style': 'objective_cloze',
        'has_error_analysis': True,
        'template_file': 'objective_cloze_template.docx'
    },
    'objective_grammar': {
        'id': 'builtin_obj_grammar',
        'name': '语法填空',
        'description': '语法填空批改，分析语法知识点和词形变化。',
        'full_score': 100,
        'scoring_dimensions': [{'name': '正确率', 'full_score': 100}],
        'style': 'objective_grammar',
        'has_error_analysis': True,
        'template_file': 'objective_grammar_template.docx'
    },
    'objective_cat_fishing': {
        'id': 'builtin_obj_cat_fishing',
        'name': '小猫钓鱼',
        'description': '小猫钓鱼（选词填空）批改，分析词义辨析和搭配。',
        'full_score': 100,
        'scoring_dimensions': [{'name': '正确率', 'full_score': 100}],
        'style': 'objective_cat_fishing',
        'has_error_analysis': True,
        'template_file': 'objective_cat_fishing_template.docx'
    }
}

# 内置默认评分标准（题型 -> 默认文件名）
BUILTIN_GRADING_STANDARDS = {
    'summary': {
        'default_file': 'standard_summary_default.docx',
        'name': '概要写作评分标准（3+2得分点）'
    }
}
