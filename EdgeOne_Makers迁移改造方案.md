# 英语作业批改工具 - EdgeOne Makers 迁移改造方案

## 一、可行性结论

**总体评估：可行，改造工作量中等（约 2 天）**

EdgeOne Makers 原生支持 Python + Flask（WSGI 模式），可以部署全栈应用。核心改造工作是将本地文件存储替换为 KV 存储，并调整目录结构适配 Makers 规范。

> **已确认决策**：生成的 Word 文档直接流式返回给用户下载，**不落地存储**；仅在 KV 中保留批改摘要记录。因此不需要 Blob 存储生成文件，大幅简化了存储改造。

---

## 二、当前项目架构分析

### 2.1 项目结构

```
grading_tool/
├── app.py                    # Flask 主应用（所有路由和业务逻辑）
├── requirements.txt          # Python 依赖
├── vercel.json               # Vercel 部署配置（需替换）
├── api/index.py              # Vercel WSGI 入口（需替换）
├── builtin_templates/        # 内置批改模板（7个docx）
│   ├── gaokao_essay_template.docx
│   ├── summary_handout_template.docx
│   ├── objective_reading_template.docx
│   ├── objective_six_choose_four_template.docx
│   ├── objective_cloze_template.docx
│   ├── objective_grammar_template.docx
│   └── objective_cat_fishing_template.docx
├── templates/
│   └── index.html            # 前端单页应用
├── templates_user/           # 用户数据目录（运行时写入）
│   ├── config.json           # 用户配置（API Key等）
│   ├── templates.json        # 自定义模板列表
│   ├── rules.json            # 自定义规则
│   ├── history.json          # 批改历史记录
│   ├── builtin_overrides.json # 内置模板覆盖配置
│   ├── grading_standards.json # 评分标准配置
│   ├── standard_summary_default.docx  # 内置评分标准文件
│   └── override_*.docx       # 用户上传的模板覆盖文件
├── uploads/                  # 上传的临时文件（运行时写入）
└── outputs/                  # 生成的批改文档（运行时写入）
```

### 2.2 主要依赖

| 依赖 | 用途 | Makers 兼容性 |
|------|------|-------------|
| flask==3.0.0 | Web 框架 | ✅ 原生支持（WSGI 模式） |
| python-docx==1.1.0 | Word 文档生成 | ✅ 可安装 |
| requests==2.31.0 | HTTP 请求（AI API） | ✅ 可安装 |
| Pillow==10.2.0 | 图片处理 | ✅ 可安装 |
| pdf2image==1.17.0 | PDF 转图片 | ⚠️ 依赖 poppler，需确认 |
| Werkzeug==3.0.1 | Flask 底层 | ✅ 随 Flask 安装 |

### 2.3 本地文件依赖分析

| 功能 | 当前实现 | 改造方向 |
|------|---------|---------|
| 内置模板读取 | `builtin_templates/` 本地读取 | 通过 `includeFiles` 打包进部署包（只读） |
| 用户配置 | `templates_user/config.json` 读写 | 环境变量 + KV 存储 |
| 自定义模板 | `templates_user/templates.json` + 上传的 docx | KV 存元数据 + Blob 存文件 |
| 自定义规则 | `templates_user/rules.json` 读写 | KV 存储 |
| 历史记录 | `templates_user/history.json` 读写 | KV 存储（仅存摘要，不含文件） |
| 评分标准 | `templates_user/standard_*.docx` 读写 | KV 存配置 + Blob 存文件 |
| 模板覆盖 | `builtin_overrides.json` + 上传文件 | KV 存配置 + Blob 存文件 |
| 文件上传 | `uploads/` 本地暂存 | 内存处理，不落地 |
| 生成的 Word 文档 | `outputs/` 本地存储 + 下载 | **流式返回，不落地存储**（✅ 已确认） |

> **存储使用汇总**：
> - **KV 存储**：配置、规则、历史摘要、模板元数据（小数据、读写频繁）
> - **Blob 存储**：用户自定义模板文件、评分标准文件、模板覆盖文件（大文件、低频读写）
> - **不存储**：生成的批改文档、上传的临时文件

---

## 三、目标架构（EdgeOne Makers）

### 3.1 目录结构

```
grading_tool/                    # 项目根目录
├── public/                      # 静态文件（自动部署为静态站点）
│   └── index.html               # 前端页面（原 templates/index.html）
├── cloud-functions/             # 后端函数目录
│   └── api/
│       └── [[default]].py       # Flask WSGI 入口（Catch-all 路由）
├── lib/                         # 业务逻辑模块（被 cloud-functions 引用）
│   ├── __init__.py
│   ├── app_core.py              # Flask app 创建 + 路由注册
│   ├── ai_service.py            # AI 调用相关
│   ├── doc_generator.py         # Word 文档生成
│   ├── storage.py               # 存储抽象层（KV + Blob）
│   └── templates_data.py        # 内置模板数据
├── builtin_templates/           # 内置模板（通过 includeFiles 打包）
│   └── *.docx
├── requirements.txt             # Python 依赖
├── edgeone.json                 # Makers 构建配置
└── README.md
```

### 3.2 架构说明

- **静态前端**：`public/index.html` 直接由 EdgeOne CDN 提供，全球加速
- **后端 API**：`cloud-functions/api/[[default]].py` 使用 Flask WSGI 模式，Catch-all 路由将所有 `/api/*` 请求转发给 Flask 处理
- **内置模板**：通过 `edgeone.json` 的 `includeFiles` 配置打包进函数部署包，运行时只读
- **用户数据**：通过 KV 存储（JSON 配置）和 Blob 存储（文件）持久化
- **文件上传/下载**：上传文件直接在内存中处理，生成的 Word 文档流式返回

---

## 四、详细改造步骤

### 阶段一：架构重构（第 1 天）

#### 1.1 拆分 app.py 为模块化结构

**目标**：将单体 `app.py` 拆分为多个模块，便于 Cloud Functions 环境使用。

**改动**：
- 新建 `lib/app_core.py`：创建 Flask app，注册所有路由
- 新建 `lib/ai_service.py`：AI 调用相关函数（`call_doubao_chat`、`grade_with_doubao`、`generate_summary_handout_with_doubao`）
- 新建 `lib/doc_generator.py`：Word 文档生成函数（`generate_essay_doc`、`generate_summary_handout_doc` 等）
- 新建 `lib/storage.py`：存储抽象层，封装 KV 和 Blob 操作
- 新建 `lib/templates_data.py`：内置模板配置常量

#### 1.2 存储抽象层设计

```python
# lib/storage.py

class Storage:
    """存储抽象层：本地开发用文件，生产用 KV + Blob"""
    
    def __init__(self, mode='local'):
        self.mode = mode  # 'local' or 'cloud'
    
    # KV 操作
    def get_json(self, key):
        """获取 JSON 数据"""
        if self.mode == 'local':
            # 本地文件读取
            ...
        else:
            # Makers KV 存储
            ...
    
    def put_json(self, key, value):
        """写入 JSON 数据"""
        ...
    
    # Blob 操作
    def get_blob(self, key):
        """获取二进制文件内容"""
        ...
    
    def put_blob(self, key, data):
        """存储二进制文件"""
        ...
    
    def delete_blob(self, key):
        """删除二进制文件"""
        ...
    
    def list_blobs(self, prefix=''):
        """列出 Blob 文件"""
        ...
```

**说明**：开发环境继续使用本地文件，生产环境自动切换到 KV + Blob。通过环境变量 `STORAGE_MODE` 控制。

#### 1.3 创建 Makers 入口文件

```python
# cloud-functions/api/[[default]].py

import sys
import os

# 将项目根目录和 lib 加入 Python 路径
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, PROJECT_ROOT)
sys.path.insert(0, os.path.join(PROJECT_ROOT, 'lib'))

from app_core import app  # Flask WSGI app
```

**路由说明**：使用 `[[default]].py` Catch-all 路由，所有 `/api/*` 请求都进入 Flask 处理。Flask 内部路由按相对路径定义（如 `/grade` 实际映射为 `/api/grade`）。

### 阶段二：存储迁移（第 1-2 天）

#### 2.1 配置数据迁移到 KV

| 原有文件 | KV Key 设计 | 说明 |
|---------|------------|------|
| `config.json` | `config:global` | 全局配置（API Key 等用环境变量，不存 KV） |
| `templates.json` | `templates:custom` | 自定义模板列表 |
| `rules.json` | `rules:list` | 自定义规则列表 |
| `history.json` | `history:list` | 批改历史记录 |
| `builtin_overrides.json` | `overrides:builtin` | 内置模板覆盖配置 |
| `grading_standards.json` | `standards:config` | 评分标准配置 |

**KV 操作封装**：在 `lib/storage.py` 中实现统一接口。

#### 2.2 文件数据迁移到 Blob 存储

| 原有文件 | Blob Key 设计 | 说明 |
|---------|--------------|------|
| 用户上传的模板 | `templates/custom/{template_id}.docx` | 自定义模板文件 |
| 模板覆盖文件 | `templates/overrides/{grade_type}.docx` | 内置模板覆盖文件 |
| 评分标准文件 | `standards/{grade_type}/{filename}` | 用户自定义评分标准 |

**Blob 操作封装**：在 `lib/storage.py` 中实现统一接口。

#### 2.3 上传文件处理改造

**当前逻辑**：文件上传 → 保存到 `uploads/` → 读取处理 → 删除

**改造后逻辑**：文件上传 → 内存中直接处理（base64 编码、文本提取等） → 不需要时自动释放

**改动点**：
- `api_grade` 接口中，上传的文件直接从 `request.files` 读取到内存
- 图片转 base64 直接在内存中完成
- docx/pdf 文本提取直接从内存流读取
- 处理完后无需删除本地文件（本来就没落地）

#### 2.4 生成文件下载改造

**当前逻辑**：生成 Word → 保存到 `outputs/` → 返回下载 URL

**改造后逻辑**：生成 Word → 写入 BytesIO → 直接流式返回响应

**改动点**：
- `generate_essay_doc` 和 `generate_summary_handout_doc` 增加支持输出到 BytesIO
- 下载接口改为直接返回文件流，`Content-Disposition: attachment`
- 不再需要 `outputs/` 目录

### 阶段三：配置与部署（第 2-3 天）

#### 3.1 edgeone.json 配置

```json
{
  "build": {
    "python": {
      "maxDuration": 300,
      "includeFiles": [
        "lib/**",
        "builtin_templates/**",
        "requirements.txt"
      ]
    }
  },
  "env": {
    "STORAGE_MODE": "cloud",
    "DOUBAO_API_KEY": "@secret/doubao_api_key"
  },
  "bindings": {
    "KV": [
      {
        "name": "GRADING_KV",
        "namespace": "english-grading-data"
      }
    ],
    "BLOB": [
      {
        "name": "GRADING_BLOB",
        "bucket": "english-grading-files"
      }
    ]
  }
}
```

> **注意**：具体配置项名称和格式以 EdgeOne Makers 官方文档为准，上述为参考结构。

#### 3.2 环境变量

| 变量名 | 说明 | 备注 |
|--------|------|------|
| `DOUBAO_API_KEY` | 豆包 API Key | 使用 Makers Secret 管理 |
| `STORAGE_MODE` | 存储模式 | `local` 或 `cloud` |
| `KV_NAMESPACE` | KV 命名空间名 | 生产环境配置 |
| `BLOB_BUCKET` | Blob 存储桶名 | 生产环境配置 |

#### 3.3 前端页面迁移

- 将 `templates/index.html` 移动到 `public/index.html`
- 前端 API 请求路径从 `/api/xxx` 保持不变（Cloud Functions 路由前缀也是 `/api`）
- 静态资源路径调整

### 阶段四：测试与验证（第 3 天）

#### 4.1 本地测试

- 确保 `STORAGE_MODE=local` 模式下功能完整
- 所有批改类型正常工作

#### 4.2 部署测试

- 连接 Git 仓库部署到 Makers
- 测试各功能模块：
  - [ ] 首页加载
  - [ ] 作文批改（中考/高考/雅思）
  - [ ] 概要写作批改
  - [ ] 概要写作讲义生成
  - [ ] 客观题批改（5种子题型）
  - [ ] 翻译批改
  - [ ] 模板管理
  - [ ] 自定义规则
  - [ ] 评分标准配置
  - [ ] 历史记录
  - [ ] 文件下载

---

## 五、风险与应对

| 风险 | 影响 | 应对方案 |
|------|------|---------|
| `pdf2image` 依赖系统 poppler | PDF 转图片功能可能失效 | 用 PyMuPDF 替代 pdf2image，PyMuPDF 纯 Python 可直接 pip 安装 |
| KV/Blob Python SDK 不明确 | 存储层开发受阻 | 先使用 HTTP API 直接调用，或参考 Makers 官方 Python 示例 |
| 冷启动延迟 | 首次请求较慢 | Cloud Functions 冷启动通常在 1-3 秒，可接受；可配置预热 |
| 单函数超时 300s | 超长 AI 生成可能超时 | 目前 AI 生成在 30-120s，在限制内；可配置 maxDuration |
| 部署包大小限制 | 内置模板 docx 文件较大 | 单文件 25MB 限制，模板文件远小于此，无问题 |
| `python-docx` 兼容性 | 文档生成异常 | 纯 Python 库，兼容性好，无需系统依赖 |

---

## 六、迁移前后对比

| 维度 | 迁移前（本地/Vercel） | 迁移后（EdgeOne Makers） |
|------|-------------------|----------------------|
| 部署方式 | 手动 / Vercel | Git 自动部署 |
| 访问速度 | 取决于服务器位置 | 全球边缘加速 |
| 扩展性 | 单服务器 | 自动弹性伸缩 |
| 数据存储 | 本地文件 | KV（配置/历史摘要）+ Blob（用户模板文件） |
| 生成文档存储 | 本地文件持久化 | **流式返回，不存储**（✅ 已确认） |
| 运维成本 | 需要管理服务器 | 全托管，免运维 |
| 成本 | 服务器费用 | 免费额度充足，超出按量付费 |
| 可用性 | 单点 | 高可用架构 |

---

## 七、实施建议

### 推荐方案：渐进式迁移

1. **第一步**：完成目录结构调整和存储抽象层，确保本地运行正常
2. **第二步**：部署到 Makers，先用只读功能验证（内置模板、批改功能）
3. **第三步**：接入 KV 存储，实现配置读写
4. **第四步**：接入 Blob 存储，实现用户文件上传功能
5. **第五步**：全功能验证，切换正式使用

### 不推荐的做法

- ❌ 一次性全部重写
- ❌ 不做存储抽象直接硬编码 KV/Blob 调用
- ❌ 忽略本地开发体验

---

## 八、参考资源

- [EdgeOne Makers Python 官方文档](https://cloud.tencent.com.cn/document/product/1552/130012)
- [EdgeOne Makers 限制与配额](https://cloud.tencent.com/document/product/1552/132789)
- [EdgeOne Makers KV 存储](https://cloud.tencent.com/document/product/1552/127431)
- [edgeone.json 配置说明](https://cloud.tencent.com/document/product/1552/127389)
