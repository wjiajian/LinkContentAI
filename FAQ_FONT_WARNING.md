# FAQ: PDF字体警告问题

## 问题描述

运行代码时出现大量警告信息：
```
Could get FontBBox from font descriptor because None cannot be parsed as 4 floats
```

## 原因分析

这个警告是由于以下原因导致的：

1. **PDF文件格式问题**：某些PDF文件中使用了不完整或损坏的字体描述
2. **pdfplumber库特性**：pdfplumber在处理这类PDF时会尝试读取字体的边界框信息，但无法从损坏的字体描述中获取
3. **非致命错误**：这个警告**不会影响功能**，只是pdfplumber库在处理有问题的PDF时的正常反馈

## 解决方案

### 方案一：使用代码中的警告抑制（已实现）

代码已在以下位置添加了警告抑制：
- 文件顶部：全局警告抑制
- PDF处理函数：局部警告抑制

```python
logging.getLogger("pdfminer").setLevel(logging.ERROR)
```

### 方案二：升级pdfplumber版本

```bash
pip install --upgrade pdfplumber
```

### 方案三：使用PyMuPDF替代pdfplumber

如果问题持续存在，可以考虑使用PyMuPDF（也叫fitz）：

```bash
pip install pymupdf
```

示例代码：
```python
import fitz  # PyMuPDF

def read_pdf_with_pymupdf(file_path: str) -> str:
    """使用PyMuPDF读取PDF"""
    try:
        doc = fitz.open(file_path)
        all_text = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            all_text.append(f"--- 第 {page_num + 1} 页 ---\n{text}")
        doc.close()
        return "\n\n".join(all_text)
    except Exception as e:
        return f"读取PDF时出错: {e}"
```

### 方案四：完全关闭stderr输出

如果仍然想看到警告，可以在运行时重定向stderr：

```bash
python write_file_excel.py 2>/dev/null
```

## 技术细节

### FontBBox是什么？

- **FontBBox**：字体的边界框，定义了字体字符绘制的边界区域
- **4 floats**：通常表示 [x_min, y_min, x_max, y_max] 四个坐标值
- **None**：PDF中的字体描述缺少这个信息

### 为什么pdfplumber会警告？

pdfplumber尝试提取PDF中的字体信息以优化文本提取，但某些PDF（特别是旧版或生成质量不高的PDF）可能没有完整的字体描述。

## 结论

✅ **这个警告是正常的，不会影响功能**
✅ **代码已经包含警告抑制，无需额外操作**
✅ **如果不想看到任何警告，可以使用方案三或四**

---

**推荐做法**：保持现有代码不变，警告信息已被自动抑制，不会影响正常使用。
