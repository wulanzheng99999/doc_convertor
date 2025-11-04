#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
文档页面设置工具
用于设置DOCX文档的纸张大小和页边距
"""

import os
import json

def load_page_settings(config_path=None):
    """
    从配置文件加载页面设置
    
    Args:
        config_path (str, optional): 配置文件路径，默认使用项目配置文件
        
    Returns:
        dict: 页面设置信息
    """
    if config_path is None:
        # 默认配置文件路径
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config', 'document_settings.json')
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config.get('page_settings', {})
    except Exception as e:
        print(f"❌ 加载配置文件失败: {e}")
        # 返回默认设置（A4）
        return {
            "paper_size": {
                "width": 21.0,
                "height": 29.7,
                "unit": "cm",
                "description": "A4 (21cm x 29.7cm)"
            },
            "margins": {
                "top": 3.1,
                "bottom": 2.8,
                "left": 2.8,
                "right": 2.8,
                "header": 2.4,
                "footer": 2.4,
                "gutter": 0,
                "unit": "cm"
            }
        }

def set_document_page_settings_com(doc_path, save_path=None, config_path=None):
    """
    使用COM库设置文档的纸张大小和页边距
    
    Args:
        doc_path (str): 输入的docx文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        config_path (str, optional): 配置文件路径
        
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

        # 打开文档
        doc = word.Documents.Open(os.path.abspath(doc_path))
        
        # 加载页面设置
        page_settings = load_page_settings(config_path)
        paper_size = page_settings.get('paper_size', {})
        margins = page_settings.get('margins', {})
        
        # 获取单位信息
        paper_unit = paper_size.get("unit", "cm")
        margin_unit = margins.get("unit", "cm")
        
        # 转换为厘米值（Word COM库使用厘米作为单位）
        if paper_unit == "cm":
            paper_width_cm = paper_size.get("width", 21.0)
            paper_height_cm = paper_size.get("height", 29.7)
        else:
            # 如果是twips单位，转换为厘米
            paper_width_cm = paper_size.get("width", 11906) / 567.0
            paper_height_cm = paper_size.get("height", 16838) / 567.0
        
        if margin_unit == "cm":
            top_margin_cm = margins.get("top", 3.1)
            bottom_margin_cm = margins.get("bottom", 2.8)
            left_margin_cm = margins.get("left", 2.8)
            right_margin_cm = margins.get("right", 2.8)
            header_margin_cm = margins.get("header", 2.4)
            footer_margin_cm = margins.get("footer", 2.4)
        else:
            # 如果是twips单位，转换为厘米
            top_margin_cm = margins.get("top", 1758) / 567.0
            bottom_margin_cm = margins.get("bottom", 1588) / 567.0
            left_margin_cm = margins.get("left", 1588) / 567.0
            right_margin_cm = margins.get("right", 1588) / 567.0
            header_margin_cm = margins.get("header", 1361) / 567.0
            footer_margin_cm = margins.get("footer", 1361) / 567.0
        
        gutter_margin_cm = margins.get("gutter", 0) / 567.0
        
        # 设置页面大小和页边距
        # 获取文档的第一个节（通常整个文档使用相同的页面设置）
        # 如果需要设置所有节，可以遍历Sections集合
        page_setup = doc.Sections(1).PageSetup  # 获取第一个节的页面设置
        
        # 设置纸张大小（A4）
        #  wdPaperA4 = 9
        page_setup.PageWidth = paper_width_cm * 28.35  # 转换为点（1厘米 ≈ 28.35点）
        page_setup.PageHeight = paper_height_cm * 28.35  # 转换为点（1厘米 ≈ 28.35点）
        
        # 设置页边距（厘米转点）
        page_setup.TopMargin = top_margin_cm * 28.35
        page_setup.BottomMargin = bottom_margin_cm * 28.35
        page_setup.LeftMargin = left_margin_cm * 28.35
        page_setup.RightMargin = right_margin_cm * 28.35
        page_setup.HeaderDistance = header_margin_cm * 28.35
        page_setup.FooterDistance = footer_margin_cm * 28.35
        page_setup.Gutter = gutter_margin_cm * 28.35
        
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
        
        print(f"✅ 已设置文档页面大小为: {paper_size.get('description', 'A4')}")
        print(f"   纸张尺寸: {paper_width_cm}cm x {paper_height_cm}cm")
        print(f"   页边距: 上{top_margin_cm}cm, 下{bottom_margin_cm}cm, "
              f"左{left_margin_cm}cm, 右{right_margin_cm}cm")
        print(f"   页眉距顶端: {header_margin_cm}cm, 页脚距底端: {footer_margin_cm}cm")
        
        return True

    except Exception as e:
        print(f"❌ 使用COM库设置文档页面大小失败: {e}")
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

def set_document_page_settings(doc_path, save_path=None, config_path=None):
    """
    设置文档的纸张大小和页边距（保持原有XML处理逻辑以确保向后兼容）
    
    Args:
        doc_path (str): 输入的docx文件路径
        save_path (str, optional): 保存路径，默认覆盖原文件
        config_path (str, optional): 配置文件路径
        
    Returns:
        bool: 操作是否成功
    """
    try:
        # 加载页面设置
        page_settings = load_page_settings(config_path)
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
        
        # 如果没有指定保存路径，则覆盖原文件
        output_path = save_path if save_path else doc_path
        
        # 创建临时目录
        import tempfile
        import zipfile
        import shutil
        
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
            
            # 查找文档末尾的分节符
            # 查找最后一个 </w:sectPr> 标签
            last_sect_pr_end = content.rfind('</w:sectPr>')
            
            if last_sect_pr_end != -1:
                # 找到分节符开始位置
                last_sect_pr_start = content.rfind('<w:sectPr', 0, last_sect_pr_end)
                
                if last_sect_pr_start != -1:
                    # 提取原有的分节符内容
                    original_sect_pr = content[last_sect_pr_start:last_sect_pr_end + len('</w:sectPr>')]
                    
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
                    new_content = content.replace(original_sect_pr, new_sect_pr)
                    
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
                    
                    print(f"✅ 已设置文档页面大小为: {paper_size.get('description', 'A4')}")
                    print(f"   纸张尺寸: {int(paper_width)} x {int(paper_height)} twips ({paper_size.get('width', 21.0)}cm x {paper_size.get('height', 29.7)}cm)")
                    print(f"   页边距: 上{int(top_margin)} twips, 下{int(bottom_margin)} twips, "
                          f"左{int(left_margin)} twips, 右{int(right_margin)} twips")
                    print(f"   页眉距顶端: {int(header_margin)} twips ({margins.get('header', 2.4)}cm), "
                          f"页脚距底端: {int(footer_margin)} twips ({margins.get('footer', 2.4)}cm)")
                    return True
                else:
                    print("❌ 未找到文档分节符")
                    return False
            else:
                # 如果没有找到分节符，则在文档末尾添加一个
                # 查找文档的结束标签
                body_end_index = content.rfind('</w:body>')
                
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
                    new_content = content[:body_end_index] + new_sect_pr + content[body_end_index:]
                    
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
                    
                    print(f"✅ 已设置文档页面大小为: {paper_size.get('description', 'A4')}")
                    print(f"   纸张尺寸: {int(paper_width)} x {int(paper_height)} twips ({paper_size.get('width', 21.0)}cm x {paper_size.get('height', 29.7)}cm)")
                    print(f"   页边距: 上{int(top_margin)} twips, 下{int(bottom_margin)} twips, "
                          f"左{int(left_margin)} twips, 右{int(right_margin)} twips")
                    print(f"   页眉距顶端: {int(header_margin)} twips ({margins.get('header', 2.4)}cm), "
                          f"页脚距底端: {int(footer_margin)} twips ({margins.get('footer', 2.4)}cm)")
                    return True
                else:
                    print("❌ 未找到文档结束标签")
                    return False
                
        finally:
            # 清理临时目录
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except Exception as e:
        print(f"❌ 设置文档页面大小失败: {e}")
        return False

def convert_cm_to_twips(cm):
    """
    将厘米转换为twips（1英寸=1440 twips，1厘米≈567 twips）
    
    Args:
        cm (float): 厘米值
        
    Returns:
        int: twips值
    """
    return int(cm * 567)

def update_config_with_cm_values(config_path=None):
    """
    根据厘米值更新配置文件中的twips值
    
    Args:
        config_path (str, optional): 配置文件路径
    """
    if config_path is None:
        config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config', 'document_settings.json')
    
    try:
        # A4尺寸：21cm x 29.7cm
        a4_width_cm = 21.0
        a4_height_cm = 29.7
        
        # 页边距（厘米）
        top_margin_cm = 3.1
        bottom_margin_cm = 2.8
        left_margin_cm = 2.8
        right_margin_cm = 2.8
        header_margin_cm = 2.4
        footer_margin_cm = 2.4
        
        # 转换为twips
        a4_width_twips = convert_cm_to_twips(a4_width_cm)
        a4_height_twips = convert_cm_to_twips(a4_height_cm)
        top_margin_twips = convert_cm_to_twips(top_margin_cm)
        bottom_margin_twips = convert_cm_to_twips(bottom_margin_cm)
        left_margin_twips = convert_cm_to_twips(left_margin_cm)
        right_margin_twips = convert_cm_to_twips(right_margin_cm)
        header_margin_twips = convert_cm_to_twips(header_margin_cm)
        footer_margin_twips = convert_cm_to_twips(footer_margin_cm)
        
        # 更新配置
        new_config = {
            "page_settings": {
                "paper_size": {
                    "width": a4_width_twips,
                    "height": a4_height_twips,
                    "unit": "twips",
                    "description": f"A4 ({a4_width_cm}cm x {a4_height_cm}cm)"
                },
                "margins": {
                    "top": top_margin_twips,
                    "bottom": bottom_margin_twips,
                    "left": left_margin_twips,
                    "right": right_margin_twips,
                    "header": header_margin_twips,
                    "footer": footer_margin_twips,
                    "gutter": 0,
                    "unit": "twips"
                }
            },
            "conversion_factors": {
                "cm_to_twips": 567,
                "inch_to_twips": 1440
            }
        }
        
        # 保存配置文件
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(new_config, f, ensure_ascii=False, indent=4)
        
        print("✅ 配置文件已更新")
        print(f"   A4尺寸: {a4_width_cm}cm x {a4_height_cm}cm ({a4_width_twips} x {a4_height_twips} twips)")
        print(f"   页边距: 上{top_margin_cm}cm, 下{bottom_margin_cm}cm, "
              f"左{left_margin_cm}cm, 右{right_margin_cm}cm")
        print(f"   页眉页脚距离: 距顶端{header_margin_cm}cm, 距底端{footer_margin_cm}cm")
        
    except Exception as e:
        print(f"❌ 更新配置文件失败: {e}")

def main():
    """主函数"""
    print("文档页面设置工具")
    print("1. 设置文档页面大小和页边距")
    print("2. 更新配置文件（根据厘米值计算twips）")
    
    # 更新配置文件
    update_config_with_cm_values()
    
    print("\n配置文件已根据以下设置更新:")
    print("- 纸张大小: A4 (21cm x 29.7cm)")
    print("- 页边距: 上3.1cm, 下2.8cm, 左2.8cm, 右2.8cm")
    print("- 页眉页脚距离: 距顶端2.4cm, 距底端2.4cm")

if __name__ == "__main__":
    main()