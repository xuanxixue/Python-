# 📄 文档句子提取工具

<div align="center">

![Python Version](https://img.shields.io/badge/Python-3.7+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)
![Status](https://img.shields.io/badge/Status-Production%20Ready-brightgreen.svg)
![NLP](https://img.shields.io/badge/NLP-spaCy%20Powered-orange.svg)

**基于spaCy的多格式文档智能句子提取工具**

[功能特性](#-功能特性) • [快速开始](#-快速开始) • [使用指南](#-使用指南) • [技术架构](#-技术架构) • [注意事项](#-注意事项)

</div>

## 📋 项目简介

**文档句子提取工具**是一款专为文档处理设计的智能工具，能够自动从多种格式的文档中提取完整的句子。无论是Word文档还是文本文件，都能快速准确地识别和分离每一个完整的句子，为后续的文本分析、数据挖掘和自然语言处理任务提供基础支持。

> **核心理念**：让文档句子提取变得简单、准确、高效

### 🎯 项目背景
- **概念提出**：玄曦雪-张悦
- **技术实现**：Deepseek AI
- **版本**：v1.0 稳定版
- **许可证**：Apache-2.0 license

---

## ✨ 功能特性

### 🎯 核心提取功能
| 功能 | 描述 | 状态 |
|------|------|------|
| 多格式支持 | DOCX、TXT等多种文档格式 | ✅ 稳定 |
| 智能句子分割 | 基于spaCy的智能句子边界识别 | ✅ 稳定 |
| 批量处理 | 支持多个文档同时处理 | ✅ 稳定 |
| 多语言支持 | 自动适配中英文语言模型 | ✅ 稳定 |
| 可视化界面 | 直观的Tkinter图形界面 | ✅ 稳定 |

### 📊 文档处理能力
| 处理模式 | 适用场景 | 特点 |
|----------|----------|------|
| DOCX解析 | Word文档处理 | 保留完整文本结构 |
| TXT解析 | 纯文本文件处理 | 快速轻量级处理 |
| 批量合并 | 多文档统一处理 | 自动内容合并 |
| 句子保持 | 原始句子完整性 | 不进行去重处理 |

### 🔧 输出管理
- **实时预览**：表格形式展示提取结果
- **统计信息**：详细的句子数量统计
- **导出功能**：标准TXT格式导出
- **序号标记**：自动为句子添加序号

### 🎨 用户界面
- **文件列表**：清晰的已上传文件显示
- **进度指示**：处理状态实时反馈
- **结果表格**：可滚动的句子展示区域
- **一键操作**：简化的用户交互流程

---

## 🚀 快速开始

### 系统要求
- **操作系统**：Windows 7/8/10/11, Linux, macOS 10.14+
- **Python版本**：3.7+（推荐3.8-3.11）
- **内存**：2GB RAM（推荐4GB）
- **存储空间**：500MB可用空间（包含语言模型）

### 安装步骤

1. **下载项目**
```bash
git clone https://github.com/your-username/document-sentence-extractor.git
cd document-sentence-extractor
```

2. **创建虚拟环境（推荐）**
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux/macOS
source venv/bin/activate
```

3. **安装Python依赖**
```bash
# 基础GUI和文档处理依赖
pip install tkinter python-docx spacy

# 如果tkinter未安装（Linux系统）
sudo apt-get install python3-tk
```

4. **安装spaCy语言模型**
```bash
# 安装中文模型（推荐用于中文文档）
python -m spacy download zh_core_web_sm

# 或者安装英文模型
python -m spacy download en_core_web_sm
```

5. **验证安装**
```bash
python -c "import tkinter; from docx import Document; import spacy; print('所有依赖安装成功!')"
```

6. **运行程序**
```bash
python sdhv文档句子提取工具.py
```

### 语言模型选择建议

| 语言 | 推荐模型 | 安装命令 | 特点 |
|------|----------|----------|------|
| 中文 | zh_core_web_sm | `python -m spacy download zh_core_web_sm` | 专门优化中文句子分割 |
| 英文 | en_core_web_sm | `python -m spacy download en_core_web_sm` | 英文处理精度高 |
| 多语言 | xx_sent_ud_sm | `python -m spacy download xx_sent_ud_sm` | 支持55种语言 |

---

## 📖 使用指南

### 基础操作流程
1. **启动程序**：运行 `python sdhv文档句子提取工具.py`
2. **上传文档**：点击"上传文档"选择DOCX或TXT文件
3. **处理文档**：点击"处理文档"开始句子提取
4. **查看结果**：在结果表格中查看提取的句子
5. **导出结果**：点击"导出为TXT"保存提取结果

### 详细操作步骤

#### 第一步：文档上传
```python
# 支持的文件格式：
- 📄 Word文档 (.docx)
- 📝 文本文件 (.txt)
- 🔄 批量选择多个文件
```

#### 第二步：智能处理
```python
# 处理过程：
1. 文档内容读取和合并
2. spaCy模型加载和文本解析
3. 智能句子边界识别
4. 句子清洗和格式化
```

#### 第三步：结果查看
```python
# 结果展示：
- 📊 统计信息：提取句子总数
- 📋 表格预览：带序号的句子列表
- 🔍 滚动查看：支持大量句子浏览
```

#### 第四步：数据导出
```python
# 导出格式：
- 标准TXT文本文件
- 包含统计头部信息
- 带序号的句子列表
- UTF-8编码确保中文兼容
```

### 界面功能区域详解

#### 📁 文档上传区
- **多文件选择**：支持一次性选择多个文档
- **格式过滤**：自动过滤DOCX和TXT文件
- **列表显示**：清晰显示已上传文件路径
- **状态管理**：根据上传状态启用处理按钮

#### ⚙️ 处理控制区
- **一键处理**：简单的开始处理按钮
- **进度反馈**：处理状态实时提示
- **错误处理**：友好的错误信息提示

#### 📊 结果显示区
- **统计面板**：实时显示提取句子数量
- **表格视图**：序号和句子的双列表格
- **滚动支持**：处理大量句子时的流畅浏览
- **完整显示**：保持句子原始格式和内容

#### 💾 导出管理区
- **格式选择**：标准TXT格式导出
- **自定义保存**：用户选择保存路径
- **完整信息**：包含统计信息的完整导出

### 处理示例

#### 输入文档示例
```
文档1.docx 内容：
自然语言处理是人工智能的重要分支。
它让计算机能够理解、解释和生成人类语言。

文档2.txt 内容：
深度学习模型在文本分析中表现出色。特别是Transformer架构。
```

#### 输出结果示例
```txt
文档句子提取结果
原始文档数量: 2
提取句子数量: 3
==================================================

1. 自然语言处理是人工智能的重要分支。

2. 它让计算机能够理解、解释和生成人类语言。

3. 深度学习模型在文本分析中表现出色。

4. 特别是Transformer架构。
```

---

## 🏗️ 技术架构

### 系统架构
```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   用户界面层     │    │   处理引擎层      │    │   数据输出层     │
│                 │    │                  │    │                 │
│ • Tkinter GUI   │◄──►│ • 文档解析器     │◄──►│ • 表格显示器    │
│ • 文件管理       │    │ • 句子提取器     │    │ • 文件导出器    │
│ • 状态反馈       │    │ • 文本清洗器     │    │ • 统计生成器    │
└─────────────────┘    └──────────────────┘    └─────────────────┘
         │                        │                       │
         ▼                        ▼                       ▼
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   文件输入层     │    │   NLP核心层       │    │   配置管理层     │
│                 │    │                  │    │                 │
│ • 多格式读取     │    │ • spaCy引擎      │    │ • 模型管理      │
│ • 编码处理       │    │ • 句子分割       │    │ • 参数配置      │
│ • 批量合并       │    │ • 语言检测       │    │ • 错误处理      │
└─────────────────┘    └──────────────────┘    └─────────────────┘
```

### 核心技术栈
| 组件 | 技术选型 | 用途 | 特点 |
|------|----------|------|------|
| 图形界面 | Tkinter | 用户交互界面 | 轻量级、原生支持 |
| 文档处理 | python-docx | Word文档解析 | 专业格式支持 |
| NLP引擎 | spaCy | 句子分割和NLP处理 | 工业级精度 |
| 文本处理 | 标准库 | 文本文件读取 | 稳定可靠 |
| 数据展示 | ttk.Treeview | 表格形式展示 | 美观易用 |

### 核心算法原理

#### 文档读取流程
```python
def read_document(self, file_path):
    """统一的文档读取接口"""
    if file_path.endswith('.docx'):
        # Word文档解析
        doc = docx.Document(file_path)
        content = [paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()]
        return '\n'.join(content)
    elif file_path.endswith('.txt'):
        # 文本文件读取
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
```

#### 句子提取算法
```python
def extract_sentences_spacy(self, text):
    """基于spaCy的智能句子提取"""
    if self.nlp is None:
        return []
    
    # spaCy文档处理
    doc = self.nlp(text)
    
    # 句子分割和清洗
    sentences = [
        sent.text.strip() 
        for sent in doc.sents 
        if sent.text.strip()
    ]
    
    return sentences
```

#### 多文档合并策略
```python
def merge_documents(self):
    """智能文档合并"""
    combined_content = []
    for file_path in self.file_paths:
        content = self.read_document(file_path)
        if content:
            combined_content.append(content)
    
    # 用换行符连接，保持段落结构
    return '\n'.join(combined_content)
```

### 处理流程详解

1. **文档读取阶段**
   - 格式自动检测
   - 编码正确处理
   - 内容完整性验证

2. **文本预处理阶段**
   - 多文档内容合并
   - 空白字符清理
   - 编码统一处理

3. **句子分割阶段**
   - spaCy模型加载
   - 智能边界检测
   - 句子完整性保证

4. **结果后处理阶段**
   - 句子清洗和格式化
   - 统计信息生成
   - 用户界面更新

---

## ⚠️ 注意事项

### 环境配置要点

#### 1. 语言模型安装
**问题**：spaCy模型下载失败
```bash
# 解决方案：使用国内镜像
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple spacy
python -m spacy download zh_core_web_sm
```

**问题**：模型加载失败
```python
# 备用方案：代码自动降级处理
try:
    self.nlp = spacy.load("zh_core_web_sm")
except OSError:
    try:
        self.nlp = spacy.load("en_core_web_sm")
    except OSError:
        # 提供友好的错误提示
        messagebox.showerror("错误", "请安装spaCy语言模型")
```

#### 2. 文件编码处理
**问题**：中文乱码
```python
# 确保使用UTF-8编码
with open(file_path, 'r', encoding='utf-8') as file:
    content = file.read()
```

#### 3. 大文件处理
**问题**：内存不足
```python
# 优化建议：
# - 分批处理特大文件
# - 增加进度指示
# - 提供内存使用监控
```

### 使用限制说明

#### 支持的文件格式
- ✅ DOCX (Microsoft Word文档)
- ✅ TXT (纯文本文件)
- ❌ PDF (不支持)
- ❌ HTML (不支持)
- ❌ RTF (不支持)

#### 处理能力限制
- **文件大小**：建议单个文件不超过10MB
- **句子数量**：支持处理数千个句子
- **语言支持**：依赖安装的spaCy模型
- **特殊格式**：复杂的表格和图片无法提取

### 性能优化建议

#### 1. 内存优化
```python
# 对于大文件，可以分块处理
chunk_size = 100000  # 10万字符为一个处理块
```

#### 2. 处理速度优化
```python
# 禁用不需要的spaCy管道组件
nlp = spacy.load("zh_core_web_sm", disable=["parser", "ner", "lemmatizer"])
```

#### 3. 用户体验优化
- 添加处理进度条
- 支持处理取消操作
- 提供更详细的错误信息

### 常见问题解答

**Q: 程序无法启动，提示模块找不到**
A: 请确保已安装所有依赖包：`pip install tkinter python-docx spacy`

**Q: 处理中文文档时句子分割不准确**
A: 请安装中文语言模型：`python -m spacy download zh_core_web_sm`

**Q: 导出文件乱码**
A: 程序使用UTF-8编码，请确保文本编辑器支持UTF-8

**Q: 处理大型文档时程序无响应**
A: 这是正常现象，程序正在处理数据。建议先处理较小的文档。

---

## 🔧 扩展开发

### 自定义句子分割规则
```python
def custom_sentence_splitter(self, text):
    """自定义句子分割逻辑"""
    # 添加基于标点的简单分割
    sentences = re.split(r'[。！？!?]', text)
    return [s.strip() for s in sentences if s.strip()]
```

### 支持新文件格式
```python
def read_pdf_document(self, file_path):
    """扩展PDF文件支持"""
    # 需要安装PyPDF2等额外库
    import PyPDF2
    # PDF读取逻辑...
```

### 界面功能扩展
```python
def add_search_functionality(self):
    """添加搜索功能"""
    # 在结果表格中添加搜索框
    search_frame = ttk.Frame(self.result_frame)
    search_frame.grid(row=0, column=1, sticky=tk.E)
```

---

## 📞 技术支持

### 问题反馈
如果您在使用过程中遇到问题：

1. **检查环境配置**：确保Python版本和依赖包正确安装
2. **查看错误信息**：程序会提供详细的错误提示
3. **提交Issue**：在项目页面描述具体问题
4. **联系开发者**：通过邮件或项目讨论区联系

### 反馈信息要求
- 操作系统和Python版本
- 错误信息和堆栈跟踪
- 涉及的文件类型和大小
- 复现问题的具体步骤

---

## 🚀 未来规划

### 功能扩展
- [ ] PDF文档支持
- [ ] 批量导出功能
- [ ] 句子去重选项
- [ ] 高级过滤规则
- [ ] 处理进度显示

### 性能优化
- [ ] 异步处理支持
- [ ] 内存使用优化
- [ ] 处理速度提升
- [ ] 大文件分块处理

### 用户体验
- [ ] 暗色主题支持
- [ ] 多语言界面
- [ ] 快捷键支持
- [ ] 配置持久化

---

<div align="center">

## 🎉 开始使用

**文档句子提取工具 - 让文档处理变得更简单、更智能！**

[快速开始](#-快速开始) • [报告问题](https://github.com/your-username/document-sentence-extractor/issues)

*从繁杂的文档中提取清晰的句子，为文本分析奠定坚实基础*

**开发者**: Deepseek, 玄曦雪-张悦  
**概念设计**: 玄曦雪-张悦  
**最后更新**: 2025年11月6日  

</div>
