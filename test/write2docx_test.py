from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docx.oxml.ns import qn
from unicodedata import category
from datetime import datetime
import os
import re

def save_to_docx(content: str, output_file: str) -> None:
    """保存内容到docx文档，支持Markdown格式"""
    if not content.strip():
        print("错误: 要保存的内容为空")
        return

    doc = Document()

    # 设置文档默认字体
    set_font(doc.styles['Normal'], 'Times New Roman', '宋体')

    # 添加文档标题
    title = doc.add_heading("融合重写文档", level=1)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    set_font(title.style, 'Times New Roman', '宋体')

    # 移除可能导致显示问题的不可见字符
    cleaned_content = ''.join(c for c in content if category(c)[0] != 'C')

    # 添加时间戳到文件名
    timestamp = datetime.now().strftime("%m%d_%H%M")
    base, ext = os.path.splitext(output_file)
    output_file_with_timestamp = f"{base}_{timestamp}{ext}"

    doc.save(output_file_with_timestamp)
    print(f"已保存融合文档到: {output_file_with_timestamp}，内容长度: {len(content)} 字符")

def set_font(style, latin_font, east_asian_font):
    """设置字体样式"""
    style.font.name = latin_font
    style.element.rPr.rFonts.set(qn('w:eastAsia'), east_asian_font)

def read_txt(input_file: str) -> str:
    """读取文本文件"""
    with open(input_file, 'r', encoding='utf-8') as f:
        return f.read()

def main():
    context = read_txt("context.txt")
    save_to_docx(context, r"D:\AIProject\DocMerge\test\testdoc.docx")

if __name__ == "__main__":
    main()