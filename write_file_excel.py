import os
import sys
import warnings
import logging

# 更彻底地抑制所有PDF相关警告
logging.getLogger("pdfminer").setLevel(logging.ERROR)

import openpyxl
import docx
import tempfile
import shutil
import uuid
import re
import json
import base64
from typing import Dict, List, Tuple, Optional
from pathlib import Path
from openpyxl.utils import get_column_letter
from openai import OpenAI
# from xbot import print

# 抑制所有库的警告
for mod in ['pdfplumber', 'pdf2image', 'PIL']:
    try:
        warnings.filterwarnings('ignore', module=mod)
    except:
        pass

# 配置多模态LLM
# TODO: 请配置您的qwen-vl API信息
QWEN_VL_CONFIG = {
    "api_key": os.getenv("QWEN_V"),  # 替换为您的API密钥
    "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",  # 通义千问API endpoint
    "model": "qwen-vl-plus",  # 或 qwen-vl-max
}

# 临时文件管理类
class TempFileManager:
    """管理临时文件和目录的生命周期"""
    def __init__(self):
        self.temp_dir = None
        self.used_paths = set()

    def __enter__(self):
        self.temp_dir = tempfile.mkdtemp(prefix="excel_img_proc_")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def get_temp_path(self, suffix="") -> str:
        """生成唯一的临时文件路径"""
        unique_id = str(uuid.uuid4())
        filename = f"{unique_id}{suffix}"
        self.used_paths.add(filename)
        return os.path.join(self.temp_dir, filename)

# --- 模块化的内容读取区域 ---
# 未来若要添加对新文件类型（例如 .csv）的支持:
# 1. 编写一个新的函数 `read_csv_content(file_path)`。
# 2. 在 FILE_READERS 字典中增加一行映射：`'.csv': read_csv_content`。

def read_txt_content(file_path: str) -> str:
    """从 .txt 文件中读取内容。"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        return f"读取 TXT 文件 '{file_path}' 时出错: {e}"

def read_docx_content(file_path: str) -> str:
    """从 .docx 文件中读取内容。"""
    try:
        doc = docx.Document(file_path)
        full_text = [para.text for para in doc.paragraphs]
        return '\n'.join(full_text)
    except Exception as e:
        return f"读取 DOCX 文件 '{file_path}' 时出错: {e}"

def read_xlsx_content(file_path: str) -> str:
    """
    从 .xlsx 文件中的所有工作表读取可见的文本内容。
    """
    try:
        # 以只读模式加载工作簿，这样性能更好，且不会意外修改文件
        workbook = openpyxl.load_workbook(file_path, read_only=True)

        all_sheets_text = []

        # 遍历工作簿中的每一个工作表
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            sheet_text = []

            # 添加工作表标题，以便区分不同工作表的内容
            sheet_text.append(f"--- 工作表: {sheet.title} ---")

            # 遍历工作表中的每一行
            for row in sheet.iter_rows():
                # 获取行中每个单元格的值，并转换为字符串，忽略空单元格
                # str(cell.value) 可以安全地处理数字、日期等不同类型
                row_values = [str(cell.value) for cell in row if cell.value is not None]

                # 如果行中有内容，则将它们用制表符连接起来
                if row_values:
                    sheet_text.append("\t".join(row_values))

            # 将当前工作表的所有文本行用换行符连接起来
            all_sheets_text.append("\n".join(sheet_text))

        # 将所有工作表的内容用两个换行符隔开，使其更清晰
        return "\n\n".join(all_sheets_text)

    except FileNotFoundError:
        return f"错误：Excel 文件未找到 '{file_path}'"
    except Exception as e:
        return f"读取 XLSX 文件 '{file_path}' 时出错: {e}"


def read_pdf_content(file_path: str) -> str:
    """
    从 .pdf 文件中读取文本内容。
    """
    try:
        # 这里使用pdfplumber库来读取PDF文本
        import pdfplumber

        # 再次确保抑制警告
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            all_text = []
            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text()
                    if page_text:
                        all_text.append(f"--- 第 {page_num} 页 ---\n{page_text}")

        return "\n\n".join(all_text)

    except ImportError:
        return "错误：需要安装 pdfplumber 库来读取PDF文件: pip install pdfplumber"
    except FileNotFoundError:
        return f"错误：PDF 文件未找到 '{file_path}'"
    except Exception as e:
        return f"读取 PDF 文件 '{file_path}' 时出错: {e}"


# --- 图片提取功能 ---
def extract_images_from_docx(docx_path: str, temp_manager: TempFileManager) -> List[str]:
    """
    从 DOCX 文件中提取所有嵌入的图片。
    返回提取的图片路径列表。
    """
    try:
        import zipfile
        import os

        image_paths = []
        docx_dir = tempfile.mkdtemp(prefix="docx_extract_")

        # DOCX 实际上是一个ZIP文件
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(docx_dir)
            media_dir = os.path.join(docx_dir, "word", "media")

            if os.path.exists(media_dir):
                for filename in os.listdir(media_dir):
                    if any(filename.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']):
                        src_path = os.path.join(media_dir, filename)
                        # 复制到我们的临时目录
                        temp_path = temp_manager.get_temp_path(suffix=f"_{filename}")
                        shutil.copy2(src_path, temp_path)
                        image_paths.append(temp_path)

        # 清理临时目录
        shutil.rmtree(docx_dir, ignore_errors=True)
        return image_paths

    except Exception as e:
        print(f"从DOCX提取图片时出错: {e}")
        return []


def extract_images_from_pdf(pdf_path: str, temp_manager: TempFileManager) -> List[str]:
    """
    从 PDF 文件中提取图片。
    返回提取的图片路径列表。
    """
    try:
        # 尝试使用 pdf2image 将PDF转换为图片
        from pdf2image import convert_from_path

        images = convert_from_path(pdf_path)
        image_paths = []

        for idx, img in enumerate(images):
            temp_path = temp_manager.get_temp_path(suffix=f"_page_{idx+1}.png")
            img.save(temp_path, 'PNG')
            image_paths.append(temp_path)

        return image_paths

    except ImportError:
        print("警告：需要安装 pdf2image 来处理PDF图片: pip install pdf2image")
        print("        还需要安装 Poppler: https://pdf2image.readthedocs.io/en/latest/installation.html")
        return []
    except Exception as e:
        print(f"从PDF提取图片时出错: {e}")
        return []


def extract_images_from_document(file_path: str, temp_manager: TempFileManager) -> List[str]:
    """
    从任何支持的文档中提取图片。
    """
    _, extension = os.path.splitext(file_path.lower())

    if extension == '.docx':
        return extract_images_from_docx(file_path, temp_manager)
    elif extension == '.pdf':
        return extract_images_from_pdf(file_path, temp_manager)
    else:
        return []


# --- 文档转Markdown功能 ---
def convert_docx_to_markdown_with_placeholders(docx_path: str, image_paths: List[str], temp_manager: TempFileManager) -> str:
    """
    将DOCX转换为带占位符的Markdown。
    改进：根据图片在文档中的实际位置插入占位符。
    策略：智能检测图片位置，如果无法精确检测则按段落间隔插入。
    """
    try:
        import zipfile
        import xml.etree.ElementTree as ET

        # 使用python-docx读取文档
        doc = docx.Document(docx_path)

        markdown_lines = []
        image_idx = 0

        # 方法1: 尝试通过XML解析来精确检测图片位置
        try:
            docx_zip = zipfile.ZipFile(docx_path)
            document_xml = docx_zip.read('word/document.xml')
            root = ET.fromstring(document_xml)

            # 定义命名空间 - 修复命名空间映射
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
            }

            # 查找所有图片及其位置
            image_positions = []
            for idx, para in enumerate(root.findall('.//w:p', ns)):
                # 检查此段落是否包含图片 - 使用多种方式检测
                # 方式1: 检查wp:docPr (drawing properties)
                has_image1 = para.find('.//wp:docPr', ns) is not None
                # 方式2: 检查a:blip (bitmap image)
                has_image2 = para.find('.//a:blip', ns) is not None
                # 方式3: 检查pic:pic (picture)
                has_image3 = para.find('.//pic:pic', ns) is not None

                has_image = has_image1 or has_image2 or has_image3

                if has_image:
                    image_positions.append(idx)

            docx_zip.close()

            # 如果检测到图片位置，使用精确插入
            if image_positions:
                for para_idx, para in enumerate(doc.paragraphs):
                    text = para.text.strip()
                    if text:
                        if para.style.name.startswith('Heading'):
                            level = para.style.name.replace('Heading ', '')
                            markdown_lines.append(f"{'#' * int(level)} {text}\n")
                        else:
                            markdown_lines.append(text + "\n")

                    # 如果当前段落有图片，插入占位符
                    if para_idx in image_positions and image_idx < len(image_paths):
                        markdown_lines.append(f"![placeholder]({image_paths[image_idx]})\n")
                        image_idx += 1

                # 如果还有剩余图片，追加到末尾
                while image_idx < len(image_paths):
                    markdown_lines.append(f"![placeholder]({image_paths[image_idx]})\n")
                    image_idx += 1

                return "\n".join(markdown_lines)

        except Exception as xml_error:
            print(f"      精确检测图片位置失败，使用fallback策略: {str(xml_error)[:80]}")

        # 方法2: Fallback - 按段落间隔插入
        paragraph_count = len([p for p in doc.paragraphs if p.text.strip()])
        if paragraph_count == 0:
            paragraph_count = 1

        # 计算间隔：尽量均匀分布
        interval = max(1, paragraph_count // max(1, len(image_paths)))

        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                if para.style.name.startswith('Heading'):
                    level = para.style.name.replace('Heading ', '')
                    markdown_lines.append(f"{'#' * int(level)} {text}\n")
                else:
                    markdown_lines.append(text + "\n")

                # 每隔一定段落数插入一张图片
                if image_idx < len(image_paths) and (len([l for l in markdown_lines if l.strip() and not l.startswith('#')]) % interval == 0):
                    markdown_lines.append(f"![placeholder]({image_paths[image_idx]})\n")
                    image_idx += 1

        # 追加剩余图片
        while image_idx < len(image_paths):
            markdown_lines.append(f"![placeholder]({image_paths[image_idx]})\n")
            image_idx += 1

        return "\n".join(markdown_lines)

    except Exception as e:
        return f"转换DOCX时出错: {e}"


def convert_pdf_to_markdown_with_placeholders(pdf_path: str, image_paths: List[str]) -> str:
    """
    将PDF转换为带占位符的Markdown。
    改进：智能检测页面中的图片位置，如果无法检测则按合理间隔插入。
    策略：优先使用页面图片检测，失败时按文本长度和页面数量分配。
    """
    try:
        # 读取PDF文本
        import pdfplumber

        # 抑制PDF字体警告
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            markdown_lines = []
            image_idx = 0

            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                page_texts = []
                page_image_counts = []

                # 第一遍：提取所有页面的文本和图片信息
                for page_num, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text()
                    page_texts.append(page_text if page_text else "")

                    # 尝试检测页面中的图片数量
                    image_count = 0
                    try:
                        if hasattr(page, 'images') and page.images:
                            image_count = len(page.images)
                        else:
                            # 备用方案：使用正则检测图片相关文本
                            if page_text:
                                # 简单 heuristic：如果页面有"图"、"image"、"图表"等词汇，认为可能包含图片
                                image_indicators = ['图', 'image', '图表', '示意图', '截图', 'Image']
                                if any(indicator.lower() in page_text.lower() for indicator in image_indicators):
                                    image_count = 1
                    except:
                        pass

                    page_image_counts.append(image_count)

                # 第二遍：生成Markdown，按检测到的图片位置插入
                for page_num, (page_text, image_count) in enumerate(zip(page_texts, page_image_counts), 1):
                    markdown_lines.append(f"--- 第 {page_num} 页 ---\n")
                    if page_text:
                        markdown_lines.append(page_text)

                    # 如果检测到页面有图片，插入相应数量的占位符
                    if image_count > 0 and image_idx < len(image_paths):
                        for _ in range(image_count):
                            if image_idx < len(image_paths):
                                markdown_lines.append(f"\n![placeholder]({image_paths[image_idx]})\n")
                                image_idx += 1
                    # 如果页面有文本但没有检测到图片，按比例插入一张
                    elif page_text and not image_count and image_idx < len(image_paths) and len(image_paths) > total_pages:
                        # 如果图片数量超过页面数，每个有文本的页面至少放一张
                        markdown_lines.append(f"\n![placeholder]({image_paths[image_idx]})\n")
                        image_idx += 1

                # 如果还有剩余图片，追加到最后一页
                while image_idx < len(image_paths):
                    markdown_lines.append(f"\n![placeholder]({image_paths[image_idx]})\n")
                    image_idx += 1

        return "\n\n".join(markdown_lines)

    except Exception as e:
        return f"转换PDF时出错: {e}"


def convert_to_markdown_with_placeholders(file_path: str, image_paths: List[str], temp_manager: TempFileManager) -> str:
    """
    将文档转换为带占位符的Markdown。
    """
    _, extension = os.path.splitext(file_path.lower())

    if extension == '.docx':
        return convert_docx_to_markdown_with_placeholders(file_path, image_paths, temp_manager)
    elif extension == '.pdf':
        return convert_pdf_to_markdown_with_placeholders(file_path, image_paths)
    else:
        # 对于其他类型，使用原始文本（暂时不支持图片占位符）
        return get_content_from_file(file_path)

# 这是分发字典，它将文件扩展名映射到正确的读取函数。
FILE_READERS = {
    '.txt': read_txt_content,
    '.docx': read_docx_content,
    '.xlsx': read_xlsx_content,
    '.pdf': read_pdf_content,
    # 在这里添加新的读取函数，例如: '.pdf': read_pdf_content
}


# --- 多模态LLM调用功能 ---
def encode_image_to_base64(image_path: str) -> str:
    """
    将图片文件编码为base64字符串。
    """
    try:
        with open(image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
            return encoded_string
    except Exception as e:
        print(f"编码图片时出错 {image_path}: {e}")
        return ""


def analyze_images_with_qwen_vl(image_paths: List[str]) -> Dict[str, str]:
    """
    使用qwen-vl模型分析图片并返回描述结果。
    返回字典: {image_path: description}
    策略：为每张图片单独调用LLM，确保每张图片都能正确解析
    """
    try:
        # 检查API配置
        if QWEN_VL_CONFIG["api_key"] == "YOUR_API_KEY_HERE" or not QWEN_VL_CONFIG["api_key"]:
            print("警告：请先配置QWEN_VL_CONFIG中的API密钥")
            return {}

        # 初始化OpenAI客户端（使用通义千问的base_url）
        client = OpenAI(
            api_key=QWEN_VL_CONFIG["api_key"],
            base_url=QWEN_VL_CONFIG["base_url"]
        )

        image_descriptions = {}

        print(f"开始分析 {len(image_paths)} 张图片...")

        # 为每张图片单独调用LLM，确保准确性
        for idx, img_path in enumerate(image_paths, 1):
            print(f"  📸 正在分析图片 {idx}/{len(image_paths)}: {os.path.basename(img_path)}")

            try:
                # 编码图片
                base64_img = encode_image_to_base64(img_path)
                if not base64_img:
                    print(f"    ❌ 编码失败")
                    image_descriptions[img_path] = "[图片编码失败]"
                    continue

                # 构建单张图片的分析请求
                content = [
                    {
                        "type": "text",
                        "text": "请详细描述这张图片的内容，包括文字、图表、布局等所有可见信息。请用中文回答。"
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_img}"
                        }
                    }
                ]

                # 调用qwen-vl模型
                response = client.chat.completions.create(
                    model=QWEN_VL_CONFIG["model"],
                    messages=[
                        {
                            "role": "user",
                            "content": content
                        }
                    ],
                    max_tokens=1500
                )

                # 获取响应
                response_text = response.choices[0].message.content
                image_descriptions[img_path] = response_text.strip()

                # 显示描述长度作为成功标志
                desc_len = len(response_text)
                print(f"    ✅ 分析完成 (描述长度: {desc_len} 字符)")

            except Exception as e:
                error_msg = f"[图片分析失败: {str(e)}]"
                print(f"    ❌ 分析失败: {str(e)[:50]}...")
                image_descriptions[img_path] = error_msg

        print(f"图片分析完成！成功分析 {len([v for v in image_descriptions.values() if not v.startswith('[')])} / {len(image_paths)} 张图片")
        return image_descriptions

    except Exception as e:
        print(f"❌ 分析图片时出错: {e}")
        return {}


# --- 占位符替换功能 ---
def replace_placeholders(markdown_text: str, image_descriptions: Dict[str, str]) -> str:
    """
    将Markdown中的图片占位符替换为实际的图片描述。
    """
    try:
        # 使用正则表达式匹配 ![placeholder](image_path) 格式
        placeholder_pattern = r'!\[placeholder\]\(([^)]+)\)'

        def replace_match(match):
            image_path = match.group(1)
            # 查找对应的描述
            if image_path in image_descriptions:
                description = image_descriptions[image_path]
                # 格式化为Markdown代码块，添加长横线分隔符
                return f"\n================\n**图片描述:**\n{description}\n================\n"
            else:
                return f"\n================\n[未找到图片 {image_path} 的描述]\n================\n"

        # 执行替换
        result = re.sub(placeholder_pattern, replace_match, markdown_text)
        return result

    except Exception as e:
        print(f"替换占位符时出错: {e}")
        return markdown_text

def get_content_from_file(file_path: str) -> str:
    """
    从文件中获取内容的通用函数。
    它使用 FILE_READERS 字典来查找并调用正确的读取器。
    """
    if not os.path.exists(file_path):
        return f"错误：链接的文件 '{file_path}' 不存在"
    
    # 获取文件的扩展名
    _, extension = os.path.splitext(file_path)
    
    # 在我们的字典中查找对应的读取函数
    reader_func = FILE_READERS.get(extension.lower())
    
    if reader_func:
        # 如果找到了读取函数，就调用它
        return reader_func(file_path)
    else:
        # 否则，返回不支持的类型错误
        return f"错误：文件 '{file_path}' 的类型 ({extension}) 不受支持"

def format_as_markdown(content: str, file_extension: str) -> str:
    """
    将提取的文本内容格式化为 Markdown 代码块。
    :param content: 从文件中读取的原始文本内容。
    :param file_extension: 文件的扩展名（例如 '.txt'），用于代码块的语言标识。
    :return: 格式化后的 Markdown 字符串。
    """
    # 移除扩展名前的点，使其成为一个更干净的语言标识符
    lang_identifier = file_extension.lstrip('.')
    
    # 对于已知不支持的标识符或空标识符，使用 'text' 作为默认
    if not lang_identifier or lang_identifier in ['docx']:
        lang_identifier = 'text'
        
    return f"```{lang_identifier}\n{content}\n```"

# --- 主 Excel 处理逻辑 ---

def process_excel_in_place(excel_path: str):
    """
    自动查找链接列，在其后插入一个新列，
    用链接文档的内容填充它，并直接在原文件上保存更改。
    新版本支持图片提取和多模态LLM分析。
    """
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook.active
        print(f"成功加载文件: '{excel_path}'")
    except FileNotFoundError:
        print(f"错误：Excel 文件 '{excel_path}' 不存在。请检查路径是否正确。")
        return
    except Exception as e:
        print(f"加载 Excel 文件 '{excel_path}' 时出错: {e}")
        return

    # 获取Excel文件所在的绝对目录
    excel_base_dir = os.path.dirname(os.path.abspath(excel_path))
    print(f"将基于此目录解析相对路径: '{excel_base_dir}'")

    all_links = [{'cell': cell, 'target': cell.hyperlink.target}
                 for row in sheet.iter_rows() for cell in row if cell.hyperlink]

    if not all_links:
        print("在此文件中未找到任何超链接。未做任何更改。")
        return

    print(f"找到了 {len(all_links)} 个超链接。")

    first_link_col_idx = all_links[0]['cell'].column
    content_col_idx = first_link_col_idx + 1

    print(f"检测到链接列为 {get_column_letter(first_link_col_idx)} 列。 "
          f"将在 {get_column_letter(content_col_idx)} 列插入新内容。")

    sheet.insert_cols(content_col_idx)

    header_cell = sheet.cell(row=1, column=content_col_idx)
    header_cell.value = "链接文档内容"
    header_cell.font = openpyxl.styles.Font(bold=True)

    # 使用临时文件管理器来管理提取的图片
    with TempFileManager() as temp_manager:
        for link_info in all_links:
            link_cell = link_info['cell']
            # 这是从Excel中读取的原始路径，可能是相对的
            relative_or_absolute_path = link_info['target']

            # 解析路径，将相对路径转换为绝对路径
            if os.path.isabs(relative_or_absolute_path):
                # 如果路径已经是绝对路径 (例如 "C:\...")，则直接使用
                full_path = relative_or_absolute_path
            else:
                # 如果是相对路径，则与Excel文件所在目录进行拼接
                full_path = os.path.join(excel_base_dir, relative_or_absolute_path)

            print(f"  - 正在处理 {link_cell.coordinate}: '{relative_or_absolute_path}' -> 解析为 '{full_path}'")

            try:
                # 步骤1: 从文档中提取图片
                print(f"    提取图片中...")
                image_paths = extract_images_from_document(full_path, temp_manager)

                if image_paths:
                    print(f"    提取到 {len(image_paths)} 张图片")
                else:
                    print(f"    未检测到图片")

                # 步骤2: 转换为带占位符的Markdown
                print(f"    转换为Markdown格式...")
                markdown_with_placeholders = convert_to_markdown_with_placeholders(
                    full_path, image_paths, temp_manager
                )

                # 步骤3: 使用LLM分析图片
                final_markdown = markdown_with_placeholders
                if image_paths:
                    print(f"    使用多模态LLM分析图片...")
                    image_descriptions = analyze_images_with_qwen_vl(image_paths)

                    if image_descriptions:
                        print(f"    替换占位符...")
                        # 步骤4: 替换占位符
                        final_markdown = replace_placeholders(
                            markdown_with_placeholders, image_descriptions
                        )
                    else:
                        print(f"    图片分析失败，使用原始内容")

                # 步骤5: 插入到Excel单元格
                content_cell = sheet.cell(row=link_cell.row, column=content_col_idx)
                content_cell.value = final_markdown

                print(f"    完成")

            except Exception as e:
                print(f"    处理出错: {e}")
                # 出错时使用原始文本
                raw_content = get_content_from_file(full_path)
                _, extension = os.path.splitext(full_path)
                md_content = format_as_markdown(raw_content, extension)
                content_cell = sheet.cell(row=link_cell.row, column=content_col_idx)
                content_cell.value = md_content

    try:
        print(f"\n正在将更改保存到原始文件: '{excel_path}'...")
        workbook.save(excel_path)
        print("处理完成！原始文件已更新。")
    except PermissionError:
        print(f"\n错误：无法保存文件。请确保 '{excel_path}' 没有被其他程序（如Excel）打开。")
    except Exception as e:
        print(f"\n保存文件 '{excel_path}' 时发生未知错误: {e}")

# --- 脚本主入口 ---
if __name__ == "__main__":
    # --- 警告 ---
    # 此脚本将直接修改您的原始文件。
    # 强烈建议在运行前对您的 Excel 文件进行备份。
    
    # --- 请在这里提供您的 Excel 文件的完整路径 ---
    excel_file_path = "C:\\Users\\Admin\\Desktop\\text\\任务管理.xlsx"
    
    process_excel_in_place(excel_file_path)