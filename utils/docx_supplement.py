"""
DOCX补充处理工具
提供修改文档中特定文本格式的功能，如将"库号：xxxxxxxxxx"信息靠右对齐
"""

import os
import sys
import importlib
import re
import threading
import time
import contextlib
from pathlib import Path

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(__file__)))

COM_RETRY_MAX = 4
COM_LOCK = threading.RLock()


RPC_RETRY_CODES = {-2147418111, -2147417846, -2147417836}

def _extract_hresult(exc):
    if hasattr(exc, "hresult") and exc.hresult is not None:
        return exc.hresult
    args = getattr(exc, "args", ())
    if args:
        first = args[0]
        if isinstance(first, tuple):
            return first[0]
        return first
    return None


def _is_rpc_retry_error(hr):
    return hr in RPC_RETRY_CODES


def _pump_com_messages(pythoncom_module, attempt, base_delay=0.4, max_delay=2.0):
    delay = min(base_delay * attempt, max_delay)
    if pythoncom_module:
        with contextlib.suppress(Exception):
            pythoncom_module.PumpWaitingMessages()
    time.sleep(delay)


def _wait_file_release(file_path, timeout=8, interval=0.3):
    if not file_path:
        return False
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with open(file_path, "rb"):
                return True
        except OSError:
            time.sleep(interval)
    return False


def _ensure_output_dir(file_path):
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)



def format_library_number_alignment(doc_path, save_path=None):
    """
    修改DOCX文档中"库号：xxxxxxxxxx"信息的格式，将其设置为靠右对齐
    
    Args:
        doc_path (str): 输入的DOCX文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        
    Returns:
        bool: 操作是否成功
    """
    try:
        # 动态导入docx模块
        docx = importlib.import_module('docx')
        Document = docx.Document
        WD_ALIGN_PARAGRAPH = importlib.import_module('docx.enum.text').WD_ALIGN_PARAGRAPH
        
        # 打开文档
        doc = Document(doc_path)
        print(f"📄 成功加载文档: {doc_path}")
        
        # 处理文档中的段落，查找"库号："开头的文本
        formatted_count = 0
        
        # 遍历文档中的所有段落（主要检查前几页的段落）
        for i, paragraph in enumerate(doc.paragraphs):
            # 限制只检查前50个段落，因为库号信息通常在文档开头
            if i > 50:
                break
                
            # 获取段落文本
            text = paragraph.text.strip()
            
            # 检查是否包含"库号："且符合格式（后面跟数字）
            if text.startswith("库号：") and len(text) > 3:
                # 检查库号后是否为数字
                library_number = text[3:]  # 获取"库号："之后的内容
                if library_number.isdigit() or (library_number.replace("-", "").isdigit()):
                    # 设置段落为右对齐
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    formatted_count += 1
                    print(f"✅ 已将段落设置为右对齐: {text}")
                    
                    # 如果只需要处理一个库号信息，可以在这里break
                    # break
        
        # 保存文档
        output_path = save_path if save_path else doc_path
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        doc.save(output_path)
        print(f"💾 文档已保存到: {output_path}")
        print(f"🎉 成功格式化了 {formatted_count} 个库号信息段落")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理文档时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


def format_library_number_in_first_pages(doc_path, save_path=None, max_pages=2):
    """
    修改DOCX文档第一页或第二页中"库号：xxxxxxxxxx"信息的格式，将其设置为靠右对齐
    
    Args:
        doc_path (str): 输入的DOCX文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        max_pages (int): 检查的最大页数，默认为2页
        
    Returns:
        bool: 操作是否成功
    """
    try:
        # 动态导入docx模块
        docx = importlib.import_module('docx')
        Document = docx.Document
        WD_ALIGN_PARAGRAPH = importlib.import_module('docx.enum.text').WD_ALIGN_PARAGRAPH
        
        # 打开文档
        doc = Document(doc_path)
        print(f"📄 成功加载文档: {doc_path}")
        
        # 处理文档中的段落，查找"库号："开头的文本
        formatted_count = 0
        
        # 遍历文档中的所有段落（主要检查前几页的段落）
        for i, paragraph in enumerate(doc.paragraphs):
            # 限制只检查前100个段落，因为库号信息通常在文档开头
            if i > 100:
                break
                
            # 获取段落文本
            text = paragraph.text.strip()
            
            # 打印前20个段落的内容用于调试
            if i < 20:
                print(f"🔍 段落 {i+1}: '{text}'")
            
            # 检查是否包含"库号："且符合格式（后面跟数字或数字加横线）
            if text.startswith("库号：") and len(text) > 3:
                # 检查库号后是否为数字或数字加横线格式
                library_number = text[3:]  # 获取"库号："之后的内容
                if library_number.isdigit() or (library_number.replace("-", "").isdigit()):
                    # 设置段落为右对齐
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    formatted_count += 1
                    print(f"✅ 已将段落设置为右对齐: {text}")
                    
                    # 如果只需要处理一个库号信息，可以在这里break
                    # break
        
        # 保存文档
        output_path = save_path if save_path else doc_path
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        doc.save(output_path)
        print(f"💾 文档已保存到: {output_path}")
        print(f"🎉 成功格式化了 {formatted_count} 个库号信息段落")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理文档时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


def find_library_numbers_in_document(doc_path, max_pages=2):
    """
    查找DOCX文档中的库号信息，用于调试
    
    Args:
        doc_path (str): 输入的DOCX文件路径
        max_pages (int): 检查的最大页数，默认为2页
        
    Returns:
        list: 找到的库号信息列表
    """
    try:
        # 动态导入docx模块
        docx = importlib.import_module('docx')
        Document = docx.Document
        
        # 打开文档
        doc = Document(doc_path)
        print(f"📄 成功加载文档: {doc_path}")
        
        # 存储找到的库号信息
        library_numbers = []
        
        # 遍历文档中的所有段落
        for i, paragraph in enumerate(doc.paragraphs):
            # 限制只检查前200个段落
            if i > 200:
                break
                
            # 获取段落文本
            text = paragraph.text.strip()
            
            # 使用正则表达式查找库号信息
            # 匹配"库号："后跟数字或数字加横线的格式
            pattern = r"[库号库号]{2}[：:]\s*([0-9\-]+)"
            match = re.search(pattern, text)
            
            if match:
                library_number = match.group(1)
                library_numbers.append({
                    'text': text,
                    'library_number': library_number,
                    'paragraph_index': i
                })
                print(f"🔍 找到库号信息: {text} (段落 {i+1})")
        
        print(f"📊 共找到 {len(library_numbers)} 个库号信息")
        return library_numbers
        
    except Exception as e:
        print(f"❌ 查找库号信息时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return []


def format_library_number_advanced(doc_path, save_path=None):
    """
    使用高级方法修改DOCX文档中库号信息的格式，将其设置为靠右对齐
    
    Args:
        doc_path (str): 输入的DOCX文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        
    Returns:
        bool: 操作是否成功
    """
    try:
        # 动态导入docx模块
        docx = importlib.import_module('docx')
        Document = docx.Document
        WD_ALIGN_PARAGRAPH = importlib.import_module('docx.enum.text').WD_ALIGN_PARAGRAPH
        
        # 打开文档
        doc = Document(doc_path)
        print(f"📄 成功加载文档: {doc_path}")
        
        # 处理文档中的段落，查找库号信息
        formatted_count = 0
        
        # 遍历文档中的所有段落
        for i, paragraph in enumerate(doc.paragraphs):
            # 限制只检查前200个段落
            if i > 200:
                break
                
            # 获取段落文本
            text = paragraph.text.strip()
            
            # 使用正则表达式查找库号信息
            # 匹配"库号："后跟数字或数字加横线的格式
            pattern = r"[库号库号]{2}[：:]\s*([0-9\-]+)"
            match = re.search(pattern, text)
            
            if match:
                library_number = match.group(1)
                # 设置段落为右对齐
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                formatted_count += 1
                print(f"✅ 已将段落设置为右对齐: {text}")
        
        # 保存文档
        output_path = save_path if save_path else doc_path
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        doc.save(output_path)
        print(f"💾 文档已保存到: {output_path}")
        print(f"🎉 成功格式化了 {formatted_count} 个库号信息段落")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理文档时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


def insert_section_break_after_toc(doc_path, save_path=None, break_type="nextpage"):
    """
    使用COM库在目录后插入分节符（目录之后、正文之前）

    Args:
        doc_path (str): 输入的docx文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        break_type (str): 分节符类型，"nextpage"=下一页分节符(默认)，"continuous"=连续分节符
    Returns:
        bool: 操作是否成功
    """
    import win32com.client as win32
    import pythoncom
    from pywintypes import com_error

    doc_path = os.path.abspath(doc_path)
    output_path = os.path.abspath(save_path) if save_path else doc_path
    last_error = None

    for attempt in range(1, COM_RETRY_MAX + 1):
        word = None
        doc = None
        initialized = False
        try:
            pythoncom.CoInitialize()
            initialized = True
            with COM_LOCK:
                word = win32.DispatchEx('Word.Application')
                with contextlib.suppress(Exception):
                    word.Options.SaveNormalPrompt = False
                    word.Options.SavePropertiesPrompt = False
            word.Visible = False
            word.DisplayAlerts = False

            doc = word.Documents.Open(doc_path)

            if doc.TablesOfContents.Count > 0:
                toc = doc.TablesOfContents(1)
                toc_range = toc.Range

                first_para = None
                for para in doc.Paragraphs:
                    if para.Range.Start > toc_range.End and para.Range.Text.strip():
                        first_para = para
                        break

                if first_para:
                    if break_type.lower() == "continuous":
                        wd_section_break = 3
                    else:
                        wd_section_break = 2
                    first_para.Range.InsertBreak(wd_section_break)
                    print(f"✅ 已在正文段落前插入{'连续' if wd_section_break == 3 else '下一页'}分节符。")
                else:
                    print("⚠️ 没找到目录后的正文段落，未插入分节符。")
            else:
                print("⚠️ 文档中没有自动生成的目录。")

            if save_path:
                if output_path.lower() == doc_path.lower():
                    doc.Save()
                    print(f"💾 已覆盖保存到: {doc_path}")
                else:
                    _ensure_output_dir(output_path)
                    doc.SaveAs(output_path)
                    print(f"💾 已另存为: {output_path}")
            else:
                doc.Save()
                print(f"💾 已覆盖保存到: {doc_path}")

            return True

        except com_error as exc:
            last_error = exc
            hr = _extract_hresult(exc)
            if _is_rpc_retry_error(hr) and attempt < COM_RETRY_MAX:
                print(f"[warn] insert_section_break_after_toc retry {attempt}/{COM_RETRY_MAX}: {exc}")
                _pump_com_messages(pythoncom, attempt)
                continue
            print(f"❌ COM方法失败: {exc}")
            break
        except Exception as exc:
            last_error = exc
            print(f"❌ COM方法失败: {exc}")
            break
        finally:
            with contextlib.suppress(Exception):
                if doc:
                    doc.Close()
            with contextlib.suppress(Exception):
                if word:
                    with contextlib.suppress(Exception):
                        word.NormalTemplate.Saved = True
                    word.Quit()
            if initialized:
                with contextlib.suppress(Exception):
                    pythoncom.CoUninitialize()
            _wait_file_release(doc_path)
            if save_path:
                _wait_file_release(output_path)

    if last_error:
        print(f"❌ COM方法失败: {last_error}")
    return False

def cancel_section_link_com(doc_path, save_path=None, section_number=2):
    """
    Cancel linkage between the specified section and the previous one using Word COM.
    """
    import win32com.client as win32
    import pythoncom
    from pywintypes import com_error

    doc_path = os.path.abspath(doc_path)
    output_path = os.path.abspath(save_path) if save_path else doc_path
    last_error = None

    for attempt in range(1, COM_RETRY_MAX + 1):
        word = None
        doc = None
        initialized = False
        section = None
        try:
            with COM_LOCK:
                pythoncom.CoInitialize()
                initialized = True
                word = win32.DispatchEx('Word.Application')
                with contextlib.suppress(Exception):
                    word.Options.SaveNormalPrompt = False
                    word.Options.SavePropertiesPrompt = False
                word.Visible = False
                word.DisplayAlerts = False
                doc = word.Documents.Open(doc_path)

            section_index = section_number - 1
            if section_index >= doc.Sections.Count or section_index < 0:
                print(f'[warn] section {section_number} out of range, total {doc.Sections.Count}')
                return False

            section = doc.Sections(section_number)

            for header_type in (1, 2, 3):
                with contextlib.suppress(Exception):
                    section.Headers(header_type).LinkToPrevious = False
            for footer_type in (1, 2, 3):
                with contextlib.suppress(Exception):
                    section.Footers(footer_type).LinkToPrevious = False

            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            if save_path:
                doc.SaveAs(output_path)
                print(f'[ok] saved as {output_path}')
            else:
                doc.Save()
                print(f'[ok] saved to {doc_path}')

            print(f'[ok] section {section_number} unlinked from previous')
            return True

        except com_error as exc:
            last_error = exc
            hr = _extract_hresult(exc)
            if _is_rpc_retry_error(hr) and attempt < COM_RETRY_MAX:
                print(f'[warn] cancel_section_link_com retry {attempt}/{COM_RETRY_MAX}: {exc}')
                _pump_com_messages(pythoncom, attempt)
                continue
            print(f'[warn] cancel_section_link_com failed on attempt {attempt}: {exc}')
            time.sleep(min(1.5 * attempt, 5))
        except Exception as exc:
            last_error = exc
            print(f'[warn] cancel_section_link_com retry {attempt}/{COM_RETRY_MAX}: {exc}')
            time.sleep(min(1.5 * attempt, 5))
        finally:
            section = None
            with contextlib.suppress(Exception):
                if doc:
                    doc.Close()
            with contextlib.suppress(Exception):
                if word:
                    with contextlib.suppress(Exception):
                        word.NormalTemplate.Saved = True
                    word.Quit()
            if initialized:
                with contextlib.suppress(Exception):
                    pythoncom.CoUninitialize()
            _wait_file_release(doc_path)
            if save_path:
                _wait_file_release(output_path)

    print(f'[error] cancel_section_link_com failed after retries: {last_error}')
    return False

def process_section2_docx(docx_path, save_path, section_index=2):
    """
    处理 docx 页脚，删除指定节的 PAGE 页码域

    Args:
        docx_path (str): 输入的 docx 文件路径
        save_path (str): 输出的 docx 文件路径
        section_index (int): 要处理的节序号（从 1 开始，比如 2 表示第二节）
    """
    try:
        import zipfile
        from lxml import etree

        # 解压 docx 到内存
        with zipfile.ZipFile(docx_path, 'r') as zin:
            filelist = zin.namelist()
            files = {name: zin.read(name) for name in filelist}

        # 找到 document.xml，定位节与 footer 的对应关系
        root = etree.fromstring(files["word/document.xml"])
        nsmap = root.nsmap
        sects = root.xpath(".//w:sectPr", namespaces=nsmap)

        if len(sects) >= section_index:
            target_sect = sects[section_index - 1]
            # 找到该节绑定的 footerReference
            footer_refs = target_sect.xpath(".//w:footerReference", namespaces=nsmap)
            for fref in footer_refs:
                rid = fref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                # 根据关系文件找到 footerX.xml
                rels_name = "word/_rels/document.xml.rels"
                rels_root = etree.fromstring(files[rels_name])
                footer_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if footer_target:
                    footer_file = "word/" + footer_target[0].get("Target")
                    if footer_file in files:
                        froot = etree.fromstring(files[footer_file])
                        # 删除 PAGE 域
                        for instr in froot.xpath(".//w:instrText", namespaces=nsmap):
                            if instr.text and "PAGE" in instr.text:
                                parent = instr.getparent()
                                if parent is not None and parent.getparent() is not None:
                                    parent.getparent().remove(parent)
                        files[footer_file] = etree.tostring(froot, xml_declaration=True, encoding="utf-8", standalone="yes")

        # 写回新的 docx
        with zipfile.ZipFile(save_path, 'w') as zout:
            for name, data in files.items():
                zout.writestr(name, data)
        
        print(f"✅ 已处理第 {section_index} 节的页码域")
        
    except Exception as e:
        print(f"❌ 处理节 {section_index} 的页码域失败: {e}")
        raise


def process_section3_docx(docx_path, save_path):
    """
    处理第三节：重置页码为 1

    Args:
        docx_path (str): 输入的 docx 文件路径
        save_path (str): 输出的 docx 文件路径
    """
    try:
        import zipfile
        from lxml import etree

        # 解压 docx
        with zipfile.ZipFile(docx_path, 'r') as zin:
            filelist = zin.namelist()
            files = {name: zin.read(name) for name in filelist}

        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        # 读取 document.xml
        root = etree.fromstring(files["word/document.xml"])
        nsmap = root.nsmap
        sects = root.xpath(".//w:sectPr", namespaces=nsmap)

        # -----------------------
        # 第三节：重置页码为 1
        # -----------------------
        if len(sects) >= 3:
            sect3 = sects[2]
            pgNumType = sect3.find("w:pgNumType", namespaces=nsmap)
            if pgNumType is None:
                pgNumType = etree.Element("{%s}pgNumType" % ns_w)
                sect3.append(pgNumType)
            pgNumType.set("{%s}start" % ns_w, "1")

        # 保存新的 docx
        files["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes")
        with zipfile.ZipFile(save_path, 'w') as zout:
            for name, data in files.items():
                zout.writestr(name, data)
        
        print("✅ 已重置第三节的页码为 1")
        
    except Exception as e:
        print(f"❌ 重置第三节页码失败: {e}")
        raise


def modify_section_page_numbers(doc_path, save_path=None):
    """
    使用COM方法修改DOCX文档中各节的页码设置
    - 移除第二节中的页码，但保留页脚及其他内容和格式
    - 将第三节中的页码设置为从1开始，同样保留页脚及其他内容和格式
    
    Args:
        doc_path (str): 输入的DOCX文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        
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

        # 初始化COM
        pythoncom.CoInitialize()
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        # 打开文档
        doc = word.Documents.Open(os.path.abspath(doc_path))
        
        # 检查文档是否至少有3节
        if doc.Sections.Count < 3:
            print(f"⚠️ 文档只有 {doc.Sections.Count} 节，至少需要3节才能执行此操作")
            return False
        
        print(f"📊 文档共有 {doc.Sections.Count} 节")
        
        # 处理第二节：移除页码但保留页脚内容
        print("🔧 处理第二节：移除页码但保留页脚内容...")
        section_2 = doc.Sections(2)  # Word的索引从1开始
        
        # 处理第二节的页脚
        for footer_type in [1]:  # 主要处理首页页脚
            try:
                footer = section_2.Footers(footer_type)
                if footer.Exists:
                    print(f"   处理第二节页脚类型 {footer_type}")
                    print(f"     处理前内容: '{footer.Range.Text.strip()}'")
                    print(f"     处理前域数量: {footer.Range.Fields.Count}")
                    
                    # 取消与前一节的链接
                    footer.LinkToPrevious = False
                    
                    # 遍历页脚中的所有域，查找并删除页码域
                    for i in range(footer.Range.Fields.Count, 0, -1):
                        field = footer.Range.Fields(i)
                        # 如果是页码域则删除 (wdFieldPage = 33)
                        if field.Type == 33:
                            print(f"     删除页码域: '{field.Code.Text.strip()}'")
                            field.Delete()
                    
                    print(f"     处理后内容: '{footer.Range.Text.strip()}'")
                    print(f"     处理后域数量: {footer.Range.Fields.Count}")
                    print(f"   ✅ 第二节页脚类型 {footer_type} 中的页码已移除")
            except Exception as e:
                print(f"   ⚠️ 处理第二节页脚类型 {footer_type} 时出错: {e}")
        
        # 处理第三节：保留页码但确保从1开始
        print("🔧 处理第三节：保留页码但确保从1开始...")
        section_3 = doc.Sections(3)  # Word的索引从1开始
        
        # 处理第三节的页脚
        for footer_type in [1]:  # 主要处理首页页脚
            try:
                footer = section_3.Footers(footer_type)
                if footer.Exists:
                    print(f"   处理第三节页脚类型 {footer_type}")
                    print(f"     处理前链接状态: {footer.LinkToPrevious}")
                    print(f"     处理前内容: '{footer.Range.Text.strip()}'")
                    print(f"     处理前域数量: {footer.Range.Fields.Count}")
                    
                    # 取消与前一节的链接
                    original_link_status = footer.LinkToPrevious
                    footer.LinkToPrevious = False
                    print(f"     取消链接后链接状态: {footer.LinkToPrevious}")
                    
                    # 如果取消链接后没有页码域，但原本是链接的，说明需要恢复页码
                    if footer.Range.Fields.Count == 0 and original_link_status:
                        # 添加页码域
                        footer.Range.Collapse(0)  # 折叠到末尾
                        if footer.Range.Text.strip():  # 如果有内容，添加换行
                            footer.Range.InsertAfter("\n")
                        footer.Range.InsertAlignmentTab(1, 1)  # 插入右对齐制表符
                        footer.Range.Fields.Add(footer.Range, 33, "", False)  # 添加页码域
                        print(f"     添加了新的页码域")
                    
                    # 更新所有页码域以确保从1开始
                    for i in range(footer.Range.Fields.Count):
                        field = footer.Range.Fields(i+1)
                        if field.Type == 33:  # 页码域
                            print(f"     更新页码域: '{field.Code.Text.strip()}'")
                            field.Update()
                    
                    print(f"     处理后内容: '{footer.Range.Text.strip()}'")
                    print(f"     处理后域数量: {footer.Range.Fields.Count}")
                    print(f"   ✅ 第三节页脚类型 {footer_type} 处理完成")
            except Exception as e:
                print(f"   ⚠️ 处理第三节页脚类型 {footer_type} 时出错: {e}")
        
        # 保存文档
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

        return True

    except Exception as e:
        print(f"❌ 修改节页码设置失败: {e}")
        import traceback
        traceback.print_exc()
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


def main():
    """
    主函数 - 提供命令行接口
    """
    print("🚀 开始执行DOCX库号信息格式处理脚本...")
    print("=" * 50)
    
    # 示例用法
    # format_library_number_alignment("input.docx", "output.docx")
    
    print("💡 使用方法:")
    print("   format_library_number_alignment('input.docx', 'output.docx')")
    print("   format_library_number_in_first_pages('input.docx', 'output.docx')")
    print("   find_library_numbers_in_document('input.docx')")
    print("   format_library_number_advanced('input.docx', 'output.docx')")
    print("   insert_section_break_after_toc('input.docx', 'output.docx')")
    print("   modify_section_page_numbers('input.docx', 'output.docx')")
    print("=" * 50)
    print("✅ 脚本执行完毕。")





if __name__ == "__main__":
    main()
