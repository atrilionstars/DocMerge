import win32com.client
import os
import time

def html_to_docx(html_path, docx_path=None):
    """
    使用Microsoft Word将HTML文件转换为DOCX格式

    参数:
    html_path (str): HTML文件的路径
    docx_path (str, 可选): 输出的DOCX文件路径。如果未指定，将使用与HTML相同的名称和位置
    """
    # 将路径转换为绝对路径
    html_path = os.path.abspath(html_path)

    # 检查HTML文件是否存在
    if not os.path.exists(html_path):
        raise FileNotFoundError(f"HTML文件不存在: {html_path}")

    # 如果未指定docx路径，使用与HTML相同的名称
    if docx_path is None:
        file_name = os.path.splitext(os.path.basename(html_path))[0]
        docx_path = os.path.join(os.path.dirname(html_path), f"{file_name}.docx")
    else:
        # 将输出路径也转换为绝对路径
        docx_path = os.path.abspath(docx_path)

    # 创建Word应用实例
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 设置为True可以看到Word操作过程

    try:
        # 打开HTML文件，使用绝对路径
        doc = word.Documents.Open(FileName=html_path, Format=8)  # Format=8表示HTML格式

        # 等待文件加载
        time.sleep(2)

        # 保存为DOCX格式
        doc.SaveAs2(FileName=docx_path, FileFormat=16)  # FileFormat=16表示DOCX格式

        print(f"成功转换: {html_path} -> {docx_path}")

    except Exception as e:
        print(f"转换过程中出错: {str(e)}")
        raise
    finally:
        # 关闭文档和Word应用
        if 'doc' in locals():
            doc.Close()
        word.Quit()

def find_first_html_file():
    """查找当前目录下的第一个HTML文件"""
    # 获取当前目录下的所有文件
    files = [f for f in os.listdir('.') if os.path.isfile(f)]
    # 筛选出HTML文件并按名称排序
    html_files = [f for f in files if f.lower().endswith('.html') or f.lower().endswith('.htm')]
    html_files.sort()
    # 返回第一个HTML文件，如果有的话
    return html_files[0] if html_files else None

def main():
    # 示例用法
    import sys

    if len(sys.argv) > 1:
        html_file = sys.argv[1]
        docx_file = sys.argv[2] if len(sys.argv) > 2 else None
    else:
        # 没有提供命令行参数，查找当前目录下的第一个HTML文件
        html_file = find_first_html_file()
        if html_file:
            print(f"未指定参数，使用当前目录下的第一个HTML文件: {html_file}")
            docx_file = os.path.splitext(html_file)[0] + ".docx"
        else:
            print("错误: 当前目录未找到任何HTML文件")
            sys.exit(1)

    try:
        html_to_docx(html_file, docx_file)
    except Exception as e:
        print(f"操作失败: {e}")


if __name__ == "__main__":
    main()
