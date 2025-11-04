"""
在Word文档的标准目录结束之后插入分页符
提供三种实现方法：
1. 使用COM库（精确在目录后插入分页符）
2. 修改XML（直接操作document.xml）
3. 使用python-docx（在文档末尾添加分页符）
"""

import os
import sys
import shutil
import tempfile
import zipfile

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

from utils.document_page_settings import load_page_settings, set_document_page_settings


def insert_section_break_after_toc_com(doc_path, save_path=None, break_type="nextpage"):
    """
    使用COM库在目录后插入分节符（目录之后、正文之前）
    
    Args:
        doc_path (str): 输入的docx文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        break_type (str): 分节符类型，"nextpage"=下一页分节符(默认)，"continuous"=连续分节符
    Returns:
        bool: 操作是否成功
    """
    word = None
    doc = None
    pythoncom = None
    
    try:
        import win32com.client as win32
        import pythoncom
        import os

        pythoncom.CoInitialize()
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(os.path.abspath(doc_path))

        if doc.TablesOfContents.Count > 0:
            toc = doc.TablesOfContents(1)
            toc_range = toc.Range

            # 找到目录之后的第一个非空白段落
            first_para = None
            for para in doc.Paragraphs:
                if para.Range.Start > toc_range.End and para.Range.Text.strip():
                    first_para = para
                    break

            if first_para:
                # 选择分节符类型
                if break_type.lower() == "continuous":
                    wdSectionBreak = 3  # 连续分节符
                else:
                    wdSectionBreak = 2  # 下一页分节符

                first_para.Range.InsertBreak(wdSectionBreak)
                print(f"✅ 已在正文段落前插入{'连续' if wdSectionBreak==3 else '下一页'}分节符。")
            else:
                print("⚠️ 没找到目录后的正文段落，未插入分节符。")
        else:
            print("⚠️ 文档中没有自动生成的目录。")

        # ---------- 保存 ----------
        if save_path:
            save_abspath = os.path.abspath(save_path)
            doc_abspath = os.path.abspath(doc_path)

            if save_abspath.lower() == doc_abspath.lower():
                doc.Save()
                print(f"💾 已覆盖保存到: {doc_abspath}")
            else:
                output_dir = os.path.dirname(save_abspath)
                if output_dir and not os.path.exists(output_dir):
                    os.makedirs(output_dir, exist_ok=True)
                doc.SaveAs(save_abspath)
                print(f"💾 已另存为: {save_abspath}")
        else:
            doc.Save()
            print(f"💾 已覆盖保存到: {os.path.abspath(doc_path)}")
        # --------------------------

        return True

    except Exception as e:
        print(f"❌ COM方法失败: {e}")
        return False

    finally:
        try:
            if doc:
                doc.Close()
        except:
            pass
        try:
            if word:
                word.Quit()
        except:
            pass
        try:
            if pythoncom:
                pythoncom.CoUninitialize()
        except:
            pass


def insert_section_break_after_toc_xml(doc_path, save_path=None):
    """
    方法2: 通过修改XML在正文开始前插入分节符
    在正文开始位置（目录结束后）插入分节符，并自动设置页面大小、页边距和页眉页脚距离
    
    Args:
        doc_path (str): 输入的docx文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        
    Returns:
        bool: 操作是否成功
    """
    try:
        # 如果没有指定保存路径，则覆盖原文件
        output_path = save_path if save_path else doc_path
        
        # 创建临时目录
        temp_dir = tempfile.mkdtemp()
        
        try:
            # 解压docx文件
            with zipfile.ZipFile(doc_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # 读取document.xml
            document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
            if not os.path.exists(document_xml_path):
                raise FileNotFoundError("document.xml not found")
            
            # 读取原始XML内容
            with open(document_xml_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 查找目录结束标记
            # 查找 </w:sdt> 标签，这是目录结构的结束标记
            toc_end_index = content.find('</w:sdt>')
            
            section_break_inserted = False
            
            if toc_end_index != -1:
                # 找到目录结束位置
                toc_end_position = toc_end_index + len('</w:sdt>')
                
                # 在目录结束后、正文开始前插入分节符
                # 使用完全空的连续分节符，不包含任何页面设置信息
                # 这样可以确保分节符不会改变页面设置
                section_break_xml = '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>'
                
                # 查找目录后第一个段落的位置
                first_body_paragraph_start = content.find('<w:p ', toc_end_position)
                
                if first_body_paragraph_start != -1:
                    # 在正文第一个段落前插入分节符
                    new_content = content[:first_body_paragraph_start] + section_break_xml + content[first_body_paragraph_start:]
                else:
                    # 如果找不到正文段落，则在目录结束后插入
                    new_content = content[:toc_end_position] + section_break_xml + content[toc_end_position:]
                
                section_break_inserted = True
                print("✅ XML方法：已在正文开始前插入分节符（保持原有页面设置）")
            else:
                print("⚠️ XML方法：文档中未找到目录结构")
                # 如果没有找到目录，仍然保存文件
                new_content = content
            
            # 加载页面设置
            page_settings = load_page_settings()
            paper_size = page_settings.get('paper_size', {})
            margins = page_settings.get('margins', {})
            
            # 获取单位信息
            paper_unit = paper_size.get("unit", "cm")
            margin_unit = margins.get("unit", "cm")
            
            # 转换为twips值
            if paper_unit == "cm":
                paper_width = paper_size.get("width", 21.0) * 567
                paper_height = paper_size.get("height", 29.7) * 567
            else:
                # 如果已经是twips单位
                paper_width = paper_size.get("width", 11906)
                paper_height = paper_size.get("height", 16838)
            
            if margin_unit == "cm":
                top_margin = margins.get("top", 3.1) * 567
                bottom_margin = margins.get("bottom", 2.8) * 567
                left_margin = margins.get("left", 2.8) * 567
                right_margin = margins.get("right", 2.8) * 567
                header_margin = margins.get("header", 2.4) * 567
                footer_margin = margins.get("footer", 2.4) * 567
            else:
                # 如果已经是twips单位
                top_margin = margins.get("top", 1758)
                bottom_margin = margins.get("bottom", 1588)
                left_margin = margins.get("left", 1588)
                right_margin = margins.get("right", 1588)
                header_margin = margins.get("header", 1361)
                footer_margin = margins.get("footer", 1361)
            
            gutter_margin = margins.get("gutter", 0)
            
            # 查找文档末尾的分节符
            # 查找最后一个 </w:sectPr> 标签
            last_sect_pr_end = new_content.rfind('</w:sectPr>')
            
            if last_sect_pr_end != -1:
                # 找到分节符开始位置
                last_sect_pr_start = new_content.rfind('<w:sectPr', 0, last_sect_pr_end)
                
                if last_sect_pr_start != -1:
                    # 提取原有的分节符内容
                    original_sect_pr = new_content[last_sect_pr_start:last_sect_pr_end + len('</w:sectPr>')]
                    
                    # 创建新的分节符，包含指定的页面设置
                    new_sect_pr = (
                        f'<w:sectPr>'
                        f'<w:pgSz w:w="{int(paper_width)}" w:h="{int(paper_height)}"/>'
                        f'<w:pgMar w:top="{int(top_margin)}" '
                        f'w:right="{int(right_margin)}" '
                        f'w:bottom="{int(bottom_margin)}" '
                        f'w:left="{int(left_margin)}" '
                        f'w:header="{int(header_margin)}" '
                        f'w:footer="{int(footer_margin)}" '
                        f'w:gutter="{int(gutter_margin)}"/>'
                        f'</w:sectPr>'
                    )
                    
                    # 替换分节符
                    new_content = new_content.replace(original_sect_pr, new_sect_pr)
            else:
                # 如果没有找到分节符，则在文档末尾添加一个
                # 查找文档的结束标签
                body_end_index = new_content.rfind('</w:body>')
                
                if body_end_index != -1:
                    # 创建新的分节符
                    new_sect_pr = (
                        f'<w:sectPr>'
                        f'<w:pgSz w:w="{int(paper_width)}" w:h="{int(paper_height)}"/>'
                        f'<w:pgMar w:top="{int(top_margin)}" '
                        f'w:right="{int(right_margin)}" '
                        f'w:bottom="{int(bottom_margin)}" '
                        f'w:left="{int(left_margin)}" '
                        f'w:header="{int(header_margin)}" '
                        f'w:footer="{int(footer_margin)}" '
                        f'w:gutter="{int(gutter_margin)}"/>'
                        f'</w:sectPr>'
                    )
                    
                    # 在body结束标签前插入分节符
                    new_content = new_content[:body_end_index] + new_sect_pr + new_content[body_end_index:]
            
            # 写入修改后的XML内容
            with open(document_xml_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            
            # 重新打包docx文件，保持与原文件相同的压缩方式
            with zipfile.ZipFile(doc_path, 'r') as original_zip:
                with zipfile.ZipFile(output_path, 'w') as new_zip:
                    # 复制所有文件，除了修改过的document.xml
                    for item in original_zip.infolist():
                        if item.filename != 'word/document.xml':
                            # 保持原有文件的压缩方式
                            new_zip.writestr(item, original_zip.read(item.filename))
                    
                    # 写入修改后的document.xml，保持原有压缩方式
                    document_info = None
                    for item in original_zip.infolist():
                        if item.filename == 'word/document.xml':
                            document_info = item
                            break
                    
                    if document_info:
                        # 使用原有的压缩方式
                        new_zip.writestr(document_info, new_content)
                    else:
                        # 如果找不到原始的document.xml信息，则使用默认方式
                        new_zip.writestr('word/document.xml', new_content)
            
            if section_break_inserted:
                print(f"✅ 已设置文档页面大小为: {paper_size.get('description', 'A4')}")
                print(f"   纸张尺寸: {int(paper_width)} x {int(paper_height)} twips ({paper_size.get('width', 21.0)}cm x {paper_size.get('height', 29.7)}cm)")
                print(f"   页边距: 上{int(top_margin)} twips, 下{int(bottom_margin)} twips, "
                      f"左{int(left_margin)} twips, 右{int(right_margin)} twips")
                print(f"   页眉距顶端: {int(header_margin)} twips ({margins.get('header', 2.4)}cm), "
                      f"页脚距底端: {int(footer_margin)} twips ({margins.get('footer', 2.4)}cm)")
            
            return True
                
        finally:
            # 清理临时目录
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except Exception as e:
        print(f"❌ XML方法失败: {e}")
        return False


# def insert_section_break_after_toc_python_docx(doc_path, save_path=None):
#     """
#     方法3: 使用python-docx在文档末尾添加分页符
#     注意：此方法无法精确在目录后插入，只能在文档末尾添加
    
#     Args:
#         doc_path (str): 输入的docx文件路径
#         save_path (str, optional): 保存路径，默认覆盖原文件
        
#     Returns:
#         bool: 操作是否成功
#     """
#     try:
#         # 动态导入避免静态检查错误
#         import importlib
#         docx_module = importlib.import_module('docx')
#         Document = docx_module.Document
        
#         section_module = importlib.import_module('docx.enum.section')
#         WD_SECTION = getattr(section_module, 'WD_SECTION')
        
#         # 打开文档
#         doc = Document(doc_path)
        
#         # 添加一个新节（这会在文档末尾添加分页符）
#         doc.add_section(WD_SECTION.NEW_PAGE)
#         print("⚠️ python-docx方法：在文档末尾添加了分页符")
#         print("💡 注意：此方法无法精确在目录后插入分页符")
        
#         # 保存文档
#         output_path = save_path if save_path else doc_path
#         output_dir = os.path.dirname(output_path)
#         if output_dir and not os.path.exists(output_dir):
#         os.makedirs(output_dir, exist_ok=True)
#         doc.save(output_path)
        
#         return True
        
#     except Exception as e:
#         print(f"❌ python-docx方法失败: {e}")
#         return False
def insert_section_break_after_toc_python_docx(doc_path, save_path=None):
    """
    在目录之后，正文内容的第一个段落前插入分节符
    注意：这里逻辑是：跳过所有带 'TOC' 样式的段落，找到正文第一个段落，插入分节符
    
    Args:
        doc_path (str): 输入的docx文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        
    Returns:
        bool: 操作是否成功
    """
    try:
        import importlib, os
        docx_module = importlib.import_module('docx')
        Document = docx_module.Document

        section_module = importlib.import_module('docx.enum.section')
        WD_SECTION = getattr(section_module, 'WD_SECTION')

        doc = Document(doc_path)

        # 遍历段落，跳过目录（一般目录段落样式是 "TOC Heading" 或者 "TOC 1" 等）
        first_body_paragraph = None
        for p in doc.paragraphs:
            style_name = p.style.name if p.style else ""
            if not style_name.startswith("TOC") and p.text.strip():
                first_body_paragraph = p
                break

        if not first_body_paragraph:
            print("⚠️ 没找到正文段落，无法插入分节符")
            return False

        # 在正文第一个段落前插入分节符
        # 方式：在该段落前新建一个段落，并设置分节符
        prior_paragraph = first_body_paragraph.insert_paragraph_before()
        prior_paragraph._p.addnext(doc._part.element.createElement("w:sectPr"))
        # 更规范的做法是使用 add_section
        doc.add_section(WD_SECTION.NEW_PAGE)
        # 但 add_section 总是在文档末尾，所以我们手动插入 sectPr

        output_path = save_path if save_path else doc_path
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        doc.save(output_path)

        print("✅ 已在正文第一个段落前插入分节符（避免落在目录内部）")
        return True

    except Exception as e:
        print(f"❌ 插入分节符失败: {e}")
        return False


# def insert_section_break_before_first_body_paragraph(doc_path, save_path=None):
#     """
#     在目录之后，正文内容的第一个段落前插入分节符
#     使用底层 XML (sectPr) 插入，避免分节符出现在目录内部
# 
#     Args:
#         doc_path (str): 输入的docx文件路径
#         save_path (str, optional): 保存路径，默认覆盖原文件
# 
#     Returns:
#         bool: 操作是否成功
#     """
#     # 这个函数使用python-docx库，如果有导入问题可以跳过
#     try:
#         import importlib, os
# 
#         docx_module = importlib.import_module('docx')
#         Document = docx_module.Document
# 
#         doc = Document(doc_path)
# 
#         # 找到第一个正文段落：跳过目录 (TOC) 段落
#         first_body_paragraph = None
#         for p in doc.paragraphs:
#             style_name = p.style.name if p.style else ""
#             if not style_name.startswith("TOC") and p.text.strip():
#                 first_body_paragraph = p
#                 break
# 
#         if not first_body_paragraph:
#             print("⚠️ 没找到正文段落，无法插入分节符")
#             return False
# 
#         # 在正文段落前插入一个新的段落 (容器)
#         prior_paragraph = first_body_paragraph.insert_paragraph_before()
# 
#         # 尝试导入OxmlElement，如果失败则跳过
#         try:
#             from docx.oxml import OxmlElement
#             sectPr = OxmlElement("w:sectPr")
#             pPr = OxmlElement("w:pPr")
#             pPr.append(sectPr)
#             prior_paragraph._p.append(pPr)
#         except ImportError:
#             # 如果导入失败，至少创建一个空段落
#             pass
# 
#         output_path = save_path if save_path else doc_path
#         output_dir = os.path.dirname(output_path)
#         if output_dir and not os.path.exists(output_dir):
#             os.makedirs(output_dir, exist_ok=True)
#         doc.save(output_path)
# 
#         print("✅ 已在正文第一个段落前插入分节符（精确避开目录）")
#         return True
# 
#     except Exception as e:
#         print(f"❌ 插入分节符失败: {e}")
#         return False


# 辅助函数
def _ensure_output_dir(file_path):
    """确保输出目录存在"""
    output_dir = os.path.dirname(file_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)


def _copy_if_needed(source_path, target_path):
    """如果路径不同则复制文件"""
    if os.path.abspath(source_path) != os.path.abspath(target_path):
        shutil.copy2(source_path, target_path)


if __name__ == "__main__":
    print("提供三种在目录后插入分页符的方法：")
    print("1. insert_section_break_after_toc_com() - 使用COM库（推荐）")
    print("2. insert_section_break_after_toc_xml() - 修改XML")
    print("3. insert_section_break_after_toc_python_docx() - 使用python-docx（功能有限）")