from docx import Document

def print_headers(doc_path):
    print(doc_path)
    doc = Document(doc_path)
    for i, section in enumerate(doc.sections):
        # 打印页眉内容
        header = section.header
        print(f"Section {i+1} Header:")
        for paragraph in header.paragraphs:
            print(paragraph.text)
        
        # 打印页脚内容
        footer = section.footer
        print(f"Section {i+1} Footer:")
        for paragraph in footer.paragraphs:
            print(paragraph.text)

# 比较两篇文档
print_headers('test.docx')
print("--------------------------------------------------")
print_headers('test6.docx')