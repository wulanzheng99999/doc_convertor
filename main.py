import os
import sys

# 添加项目路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from service.converter import DocumentConverter

def main():
    # 指定文件路径
    source_file = r"docx\test.docx"
    
    # 创建result目录
    result_dir = os.path.join(current_dir, "result")
    os.makedirs(result_dir, exist_ok=True)
    
    # 输出文件路径
    output_file = os.path.join(result_dir, "formatted_document.docx")
    
    # 使用默认模板（template/reference.docx）
    
    # 执行转换
    with DocumentConverter() as converter:
        success = converter.convert_document(
            source_file=source_file,
            output_file=output_file,
            header_text="数字总师可行性报告",
            toc_title="目 录",
            save_intermediate=False
        )

if __name__ == "__main__":
    main()