# 🚀 快速开始指南

---

## ⚡ 立即开始

### 1️⃣ 安装依赖

```bash
pip install -r requirements.txt
```

**额外安装（PDF图片支持）**：
- **Windows**: 下载 [Poppler for Windows](https://blog.alivate.com.au/poppler-windows/)
- **macOS**: `brew install poppler`
- **Linux**: `sudo apt-get install poppler-utils`

### 2️⃣ 配置API密钥

在 `write_file_excel.py` 顶部修改：

```python
QWEN_VL_CONFIG = {
    "api_key": "您的通义千问API密钥",  # ⚠️ 必填
    "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
    "model": "qwen-vl-plus",  # 或 qwen-vl-max
}
```

### 3️⃣ 运行脚本

**方法一**：修改脚本中的文件路径
```python
excel_file_path = "您的Excel文件路径.xlsx"
```

然后运行：
```bash
python write_file_excel.py
```

**方法二**：导入模块使用
```python
from write_file_excel import process_excel_in_place
process_excel_in_place("您的Excel文件路径.xlsx")
```

---

## 📁 文件说明

| 文件 | 用途 |
|------|------|
| `write_file_excel.py` | 主程序文件（579行） |
| `requirements.txt` | 依赖列表 |
| `CLAUDE.md` | 代码指南 |
| `REFACTORING_SUMMARY.md` | 重构总结 |
| `FAQ_FONT_WARNING.md` | PDF警告问题FAQ |
| `QUICK_START.md` | 本文件 - 快速开始 |

---

## ✨ 核心功能

### 🎯 支持的文档类型
- ✅ `.txt` - 纯文本文件
- ✅ `.docx` - Word文档（含图片）
- ✅ `.xlsx` - Excel文件
- ✅ `.pdf` - PDF文档（含图片）

### 🔄 处理流程

```
Excel中的链接文档
    ↓
提取文档中的图片
    ↓
转换为Markdown（含占位符）
    ↓
调用qwen-vl分析图片
    ↓
替换占位符为图片描述
    ↓
插入完整的Markdown到Excel
```

### 📊 输出示例

```markdown
--- 第 1 页 ---
这是文档的文本内容...

**图片描述:**
这是一张柱状图，显示2023年Q1-Q4的销售数据...
- Q1: 150万元
- Q2: 180万元
- Q3: 210万元
- Q4: 260万元

更多内容...
```

---

## ⚠️ 重要提示

1. **备份Excel文件**：脚本会直接修改原始文件，请先备份
2. **API密钥必需**：必须配置有效的通义千问API密钥
3. **网络要求**：需要能访问阿里云DashScope API
4. **处理时间**：带图片的文档处理时间较长，请耐心等待
5. **自动清理**：临时图片文件会自动清理，无需手动处理

---

## 🆘 遇到问题？

**Q: 出现 "Could get FontBBox" 警告**
A: 已解决，参阅 `FAQ_FONT_WARNING.md`

**Q: PDF图片无法提取**
A: 检查是否安装了 Poppler（Windows需要单独安装）

**Q: API调用失败**
A: 检查API密钥是否正确配置，是否有网络连接

**Q: 处理速度慢**
A: 正常现象，图片分析需要时间，且受API调用速度影响

---

## 📚 更多信息

- **代码架构**: 查看 `CLAUDE.md`
- **重构详情**: 查看 `REFACTORING_SUMMARY.md`
- **PDF警告**: 查看 `FAQ_FONT_WARNING.md`

---

**开始使用吧！** 🎉
