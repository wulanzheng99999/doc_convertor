import os
import win32com.client as win32
import pythoncom
import zipfile
import xml.etree.ElementTree as ET
import tempfile
import shutil

def update_toc_title(docx_path, new_title, original_title=None):
    """
    将指定路径的docx文件的目录标题改为指定的内容
    
    参数:
        docx_path: Word文档路径
        new_title: 新的目录标题
        original_title: 原始目录标题（可选，默认为常见的中英文目录标题）
    """
    # 确保绝对路径
    docx_path = os.path.abspath(docx_path)
    
    # 检查文件是否存在
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"文件不存在: {docx_path}")
    
    # 初始化COM
    pythoncom.CoInitialize()
    
    word = None
    doc = None
    try:
        # 初始化Word COM
        word = win32.Dispatch('Word.Application')
        word.Visible = False  # 后台运行
        word.DisplayAlerts = False  # 禁用警告
        
        # 打开文档
        doc = word.Documents.Open(docx_path)
        
        # 定义要查找的目录标题
        if original_title:
            toc_titles = [original_title]
        else:
            # 默认的中英文目录标题
            toc_titles = ["Table of Contents", "目录"]
        
        found_and_replaced = False
        for find_text in toc_titles:
            # 使用Word的查找替换功能
            selection = word.Selection
            selection.HomeKey(Unit=6)  # 移动到文档开头
            
            # 查找文本
            found = word.Selection.Find.Execute(
                FindText=find_text,
                Forward=True,
                Wrap=1,  # 搜索整个文档
                Format=False,
                MatchCase=False,
                MatchWholeWord=True  # 全词匹配
            )
            
            if found:
                # 替换为新标题
                word.Selection.TypeText(new_title)
                found_and_replaced = True
                break
        
        # if not found_and_replaced:
        #     print(f"未找到目录标题: {toc_titles}")
        
        # 保存文档
        doc.Save()
        
    except Exception as e:
        print(f"操作失败: {e}")
        import traceback
        traceback.print_exc()
        raise  # 重新抛出异常以便调用者处理
    finally:
        # 关闭文档和Word应用
        try:
            if doc is not None:
                doc.Close()
        except:
            pass
            
        try:
            if word is not None:
                word.Quit()
        except:
            pass
            
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def update_toc_title_xml(docx_path, new_title):
    """
    通过直接操作XML来修改Word文档中的目录标题，保留原有格式
    先定位目录结构的XML结构，再在其中寻找标题内容进行替换

    参数:
        docx_path: Word文档路径
        new_title: 新的目录标题
    """
    # 确保绝对路径
    docx_path = os.path.abspath(docx_path)

    # 检查文件是否存在
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"文件不存在: {docx_path}")

    # 创建临时目录用于解压
    temp_dir = tempfile.mkdtemp()

    try:
        # 解压docx文件
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # 读取document.xml文件
        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        if not os.path.exists(document_xml_path):
            raise FileNotFoundError(f"未找到document.xml文件: {document_xml_path}")

        # 读取原始XML内容
        with open(document_xml_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 查找并替换目录标题
        found_and_updated = False

        # 查找标准目录结构中的标题
        if '<w:docPartGallery w:val="Table of Contents"/>' in content:
            # 查找目录标题"目录"
            if '<w:t>目录</w:t>' in content:
                content = content.replace('<w:t>目录</w:t>', f'<w:t>{new_title}</w:t>', 1)
                found_and_updated = True
            # 查找目录标题"Table of Contents"
            elif '<w:t>Table of Contents</w:t>' in content:
                content = content.replace('<w:t>Table of Contents</w:t>', f'<w:t>{new_title}</w:t>', 1)
                found_and_updated = True
        else:
            # print("未找到标准目录结构，尝试使用通用查找方式...")
            # 通用查找方式
            if '<w:t>目录</w:t>' in content:
                content = content.replace('<w:t>目录</w:t>', f'<w:t>{new_title}</w:t>', 1)
                found_and_updated = True
            elif '<w:t>Table of Contents</w:t>' in content:
                content = content.replace('<w:t>Table of Contents</w:t>', f'<w:t>{new_title}</w:t>', 1)
                found_and_updated = True

        # if not found_and_updated:
        #     print("未找到目录标题")

        # 保存修改后的XML内容
        with open(document_xml_path, 'w', encoding='utf-8') as f:
            f.write(content)

        # 创建新的临时文件路径
        new_docx_path = docx_path + ".new"

        # 重新打包为docx文件，保持原有文件顺序
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            with zipfile.ZipFile(new_docx_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                # 复制所有文件，除了document.xml
                for item in zip_ref.infolist():
                    if item.filename != 'word/document.xml':
                        new_zip.writestr(item, zip_ref.read(item.filename))

                # 添加修改后的document.xml
                new_zip.write(document_xml_path, 'word/document.xml')

        # 替换原文件
        os.replace(new_docx_path, docx_path)

    except Exception as e:
        print(f"操作失败: {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        # 清理临时目录
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)