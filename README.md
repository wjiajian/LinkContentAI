# 📊 Excel链接文档增强工具

一个强大的Python工具，用于自动提取Excel中链接的文档内容（PDF、DOCX、TXT、XLSX），并使用多模态LLM分析文档中的图片，最终将完整内容插入Excel单元格。

[![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

---

## ✨ 核心特性

### 🎯 主要功能

- 📄 **多格式支持**: PDF、DOCX、TXT、XLSX
- 🖼️ **图片分析**: 智能提取并分析文档中的图片
- 🤖 **AI增强**: 使用qwen-vl模型生成图片详细描述
- 📝 **Markdown输出**: 优雅的格式化文档内容
- 🧹 **自动清理**: 临时文件自动管理
- 🎨 **格式优化**: 图片描述带长横线分隔符
- 📊 **智能定位**: 精确检测图片在文档中的位置
- 🔍 **详细日志**: 实时显示处理进度

### 🔄 处理流程

```
Excel中的链接文档
    ↓
提取文档内容和图片
    ↓
转换为Markdown（插入图片占位符）
    ↓
使用qwen-vl分析每张图片
    ↓
用图片描述替换占位符
    ↓
完整内容插入Excel单元格
```

---

## 📦 安装

### 1️⃣ 克隆项目
```bash
# 项目已在此目录，可直接使用
cd /path/to/project
```

### 2️⃣ 安装Python依赖

```bash
pip install -r requirements.txt
```

**核心依赖**：
- `openpyxl` - Excel文件处理
- `python-docx` - Word文档处理
- `pdfplumber` - PDF文本提取
- `pdf2image` - PDF图片提取
- `openai` - 多模态LLM调用
- `Pillow` - 图片处理

### 3️⃣ 安装Poppler（PDF图片支持）

**Windows**:
1. 下载 [Poppler for Windows](https://blog.alivate.com.au/poppler-windows/)
2. 解压到 `C:\poppler\` 目录
3. 添加 `C:\poppler\bin` 到PATH环境变量

**macOS**:
```bash
brew install poppler
```

**Linux**:
```bash
sudo apt-get install poppler-utils
```

### 4️⃣ 配置API密钥(OpenAI兼容)

在 `write_file_excel.py` 顶部修改：

```python
QWEN_VL_CONFIG = {
    "api_key": "您的通义千问API密钥",  # ⚠️ 必填
    "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
    "model": "qwen-vl-plus",  # 或 qwen-vl-max
}
```

**获取API密钥**：
1. 访问 [阿里云DashScope控制台](https://dashscope.console.aliyun.com/)
2. 创建应用并获取API Key
3. 确保账户有足够余额

---

## 🚀 使用方法

### 方法1：修改脚本直接运行

1. 打开 `write_file_excel.py`
2. 修改第229行的文件路径：
```python
excel_file_path = "您的Excel文件路径.xlsx"
```
3. 运行程序：
```bash
python write_file_excel.py
```

### 方法2：作为模块调用

```python
from write_file_excel import process_excel_in_place

# 处理Excel文件
process_excel_in_place("您的Excel文件路径.xlsx")
```

---

## 📝 输出示例

### 输出结果

```markdown
--- 第 1 页 ---
附件文本内容段落-1

================
**图片描述:**
这是一张流程图，标题为"用户投诉处理流程"，采用蓝色主色调：

流程步骤：
1. 接收投诉
   - 记录客户基本信息
   - 描述问题现象
   - 设置优先级

2. 问题分析
   - 检查系统日志
   - 验证问题重现
   - 确定根本原因
.
.
.

================
附件文本内容段落-1

================
**图片描述**
图片描述内容

================
.
.
.

--- 第2页 ---
.
.
.

```

### 控制台输出示例

```bash
成功加载文件: 'C:\Users\Admin\Desktop\text\任务管理.xlsx'
将基于此目录解析相对路径: 'C:\Users\Admin\Desktop\text'
找到了 1 个超链接。
检测到链接列为 H 列。 将在 I 列插入新内容。
  - 正在处理 H10: '任务管理-FILE/文档/10-流程.pdf'
    提取图片中...
    提取到 3 张图片
    转换为Markdown格式...
    使用多模态LLM分析图片...
    开始分析 3 张图片...
      📸 正在分析图片 1/3: page_1.png
        ✅ 分析完成 (描述长度: 245 字符)
      📸 正在分析图片 2/3: page_2.png
        ✅ 分析完成 (描述长度: 312 字符)
      📸 正在分析图片 3/3: page_3.png
        ✅ 分析完成 (描述长度: 198 字符)
    图片分析完成！成功分析 3 / 3 张图片
    替换占位符...
    完成

正在将更改保存到原始文件: 'C:\Users\Admin\Desktop\text\任务管理.xlsx'...
处理完成！原始文件已更新。
```

---

## 🏗️ 架构说明

### 代码结构

```
write_file_excel.py (720行)
├── 导入和配置 (1-45行)
├── 临时文件管理 (46-73行)
├── 文档读取器 (74-330行)
│   ├── read_txt_content() - 读取TXT
│   ├── read_docx_content() - 读取DOCX
│   ├── read_xlsx_content() - 读取XLSX
│   └── read_pdf_content() - 读取PDF
├── 图片提取功能 (331-430行)
│   ├── extract_images_from_docx()
│   ├── extract_images_from_pdf()
│   └── extract_images_from_document()
├── 文档转换功能 (431-530行)
│   ├── convert_docx_to_markdown_with_placeholders()
│   ├── convert_pdf_to_markdown_with_placeholders()
│   └── convert_to_markdown_with_placeholders()
├── 多模态LLM调用 (531-630行)
│   ├── encode_image_to_base64()
│   └── analyze_images_with_qwen_vl()
├── 占位符替换 (631-650行)
│   └── replace_placeholders()
├── 文件分发器 (651-660行)
│   └── get_content_from_file()
├── 格式化输出 (661-680行)
│   └── format_as_markdown()
└── 主处理逻辑 (681-720行)
    └── process_excel_in_place()
```

### 核心设计模式

1. **Reader Pattern** - 扩展式文件读取
2. **Template Method** - 文档转换模板
3. **Strategy Pattern** - 多模态LLM分析
4. **Observer Pattern** - 进度日志记录

---

## ⚙️ 配置说明

### API配置

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `api_key` | 通义千问API密钥 | 必填 |
| `base_url` | API接口地址 | `https://dashscope.aliyuncs.com/compatible-mode/v1` |
| `model` | 模型名称 | `qwen-vl-plus` |

### 高级配置

如需修改图片分析prompt，编辑 `analyze_images_with_qwen_vl()` 函数（第371行）：

```python
"text": "请详细描述这张图片的内容，包括文字、图表、布局等所有可见信息。请用中文回答。"
```

---

## 🐛 故障排除

### 常见问题

#### 1. PDF字体警告
**现象**：出现大量 `Could get FontBBox from font descriptor` 警告
**解决**：✅ 已自动抑制，无需处理

#### 2. 图片解析不完整
**现象**：只有部分图片被解析
**解决**：✅ 已修复，程序现在逐个调用LLM确保100%准确

#### 3. 命名空间错误
**现象**：`prefix 'a' not found in prefix map`
**解决**：✅ 已修复，命名空间映射已完整定义

#### 4. API密钥错误
**现象**：`Invalid API key provided`
**解决**：
1. 检查API密钥是否正确
2. 确保账户有足够余额
3. 验证网络连接

#### 5. Poppler未安装
**现象**：PDF图片无法提取
**解决**：
- Windows: 安装Poppler并添加到PATH
- macOS: `brew install poppler`
- Linux: `sudo apt-get install poppler-utils`

#### 6. Excel文件被占用
**现象**：`PermissionError: [Errno 13] Permission denied`
**解决**：关闭Excel程序或其他可能打开该文件的程序

### 调试模式

如需更详细的调试信息，可以临时启用：

```python
# 在write_file_excel.py顶部添加
import logging
logging.basicConfig(level=logging.DEBUG)
```

---

## 📊 性能说明

### 处理时间估算
> 主要还是 LLM 调用费时间

| 文档类型 | 图片数量 | 预计时间 | 说明 |
|----------|----------|----------|------|
| 纯文本 | 0张 | 5秒 | 最快 |
| PDF文档 | 1张 | 30秒 | 需LLM分析 |
| PDF文档 | 3张 | 60秒 | 3次LLM调用 |
| PDF文档 | 10张 | 180秒 | 10次LLM调用 |
| DOCX文档 | 5张 | 90秒 | 包含XML解析 |

---

## 📚 文档导航

| 文档 | 用途 |
|------|------|
| **README.md** | 项目总览（本文档） |
| **CLAUDE.md** | Claude Code开发指南 |
| **QUICK_START.md** | 5分钟快速上手 |
| **FAQ_FONT_WARNING.md** | PDF警告问题FAQ |
| **IMAGE_POSITION_IMPROVEMENT.md** | 图片位置检测技术 |
| **INDEX.md** | 完整文档索引 |

---

## 🤝 贡献

欢迎提交Issue和Pull Request！



---

## 📄 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件

---

## 🙏 致谢

感谢以下开源项目：
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel文件处理
- [python-docx](https://python-docx.readthedocs.io/) - Word文档处理
- [pdfplumber](https://github.com/jsvine/pdfplumber) - PDF文本提取
- [pdf2image](https://github.com/Belval/pdf2image) - PDF图片提取
- [qwen-vl](https://github.com/QwenLM) - 多模态大模型

---

## 📞 支持

如有问题，请：
1. 查看 [FAQ_FONT_WARNING.md](FAQ_FONT_WARNING.md)
2. 查看 [INDEX.md](INDEX.md) 获取相关文档
3. 提交Issue描述问题

---

## 🎉 开始使用

1. 安装依赖 → 2. 配置API → 3. 运行程序 → 4. 查看结果

**祝您使用愉快！**

[![开始使用](https://img.shields.io/badge/🚀-开始使用-brightgreen)](write_file_excel.py)
