"""
DOCX页眉页脚内容提取与替换工具

该模块提供了以下功能：
1. 提取DOCX文档中指定节的页眉页脚内容（不包括页码）
2. 将提取的页眉页脚内容替换到另一个DOCX文档的对应节中
3. 保持目标文档原有的页码、格式等内容不变

使用方法：
- extract_header_footer_content(docx_path, section_index): 提取指定节的页眉页脚内容
- replace_header_footer_content(source_docx_path, target_docx_path, source_section_index, target_section_index, save_path): 
  将源文档的页眉页脚内容替换到目标文档
"""

import os
import re
import zipfile
from lxml import etree
from typing import Dict, Any


def extract_header_footer_content(docx_path: str, section_index: int = 1) -> Dict[str, Any]:
    """
    提取DOCX文档中指定节的页眉页脚内容（不包括页码）
    
    Args:
        docx_path (str): DOCX文件路径
        section_index (int): 节索引（从1开始）
        
    Returns:
        Dict[str, Any]: 包含页眉和页脚内容的字典
    """
    print(f"🔍 开始提取文档页眉页脚内容: {docx_path}")
    print(f"📝 目标节号: {section_index}")
    
    try:
        # 解压 docx 到内存
        with zipfile.ZipFile(docx_path, 'r') as zin:
            filelist = zin.namelist()
            files = {name: zin.read(name) for name in filelist}

        # 找到 document.xml，定位节与 header/footer 的对应关系
        root = etree.fromstring(files["word/document.xml"])
        nsmap = root.nsmap
        sects = root.xpath(".//w:sectPr", namespaces=nsmap)
        
        print(f"📊 找到 {len(sects)} 个节")

        if len(sects) < section_index:
            raise ValueError(f"文档共有 {len(sects)} 节，不能操作第 {section_index} 节")

        target_sect = sects[section_index - 1]
        
        # 查找该节的 headerReference 和 footerReference
        header_refs = target_sect.xpath("./w:headerReference", namespaces=nsmap)
        footer_refs = target_sect.xpath("./w:footerReference", namespaces=nsmap)
        
        print(f"📋 找到 {len(header_refs)} 个页眉引用, {len(footer_refs)} 个页脚引用")
        
        result = {
            "headers": {},
            "footers": {}
        }
        
        # 提取页眉内容（保留格式信息）
        for i, href in enumerate(header_refs):
            rid = href.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            header_type = href.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 headerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                header_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if header_target:
                    header_file = "word/" + header_target[0].get("Target")
                    if header_file in files:
                        hroot = etree.fromstring(files[header_file])
                        print(f"  📄 header_fiel是: {header_file}")
                        print(f"  📄 hroot是: {hroot}")
                        # 提取文本内容和格式信息，排除页码域
                        header_content1 = extract_formatted_content(hroot, nsmap)
                        print(f"  📄 提取headrcontent ({header_type}): {header_content1}")
                        parts = header_content1.strip().split()
                        header_content = f"{parts[0]}\t{parts[1]}"
                        result["headers"][header_type] = header_content
                        print(f"  📄 提取页眉 ({header_type}): {header_content}")

        # 提取页脚内容（保留格式信息）
        for i, fref in enumerate(footer_refs):
            rid = fref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            footer_type = fref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 footerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                footer_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if footer_target:
                    footer_file = "word/" + footer_target[0].get("Target")
                    if footer_file in files:
                        froot = etree.fromstring(files[footer_file])
                        
                        # 提取文本内容和格式信息，排除页码域
                        footer_content = extract_formatted_content(froot, nsmap)
                        result["footers"][footer_type] = footer_content
                        print(f"  📄 提取页脚 ({footer_type}): {footer_content}")
        
        print(f"✅ 页眉页脚内容提取完成")
        return result
        
    except Exception as e:
        print(f"❌ 提取页眉页脚内容失败: {e}")
        raise


def extract_formatted_content(root, nsmap):
    """
    提取带有格式信息的内容（保留制表符、换行符等）
    
    Args:
        root: XML根节点
        nsmap: 命名空间映射
        
    Returns:
        str: 格式化的内容
    """
    content_parts = []
    
    # 遍历所有段落
    for p_elem in root.xpath(".//w:p", namespaces=nsmap):
        para_parts = []
        
        # 遍历段落中的所有运行（run）
        for r_elem in p_elem.xpath(".//w:r", namespaces=nsmap):
            # 检查是否有制表符
            tabs = r_elem.xpath(".//w:tab", namespaces=nsmap)
            for _ in tabs:
                para_parts.append("\t")
            
            # 检查是否有文本
            for t_elem in r_elem.xpath(".//w:t", namespaces=nsmap):
                if t_elem.text:
                    para_parts.append(t_elem.text)
            
            # 检查是否有换行符
            brs = r_elem.xpath(".//w:br", namespaces=nsmap)
            for _ in brs:
                para_parts.append("\n")
        
        # 将段落内容添加到结果中
        content_parts.append("".join(para_parts))
    print(f"  📄 为什么呢: {content_parts}")
    # 用换行符连接所有段落
    return "\n".join(content_parts)


def replace_formatted_content(root, new_content, nsmap):
    """
    替换带有格式的内容（保留原有格式结构）
    
    Args:
        root: XML根节点
        new_content (str): 新的内容
        nsmap: 命名空间映射
    """
    # 获取所有段落
    p_elems = root.xpath(".//w:p", namespaces=nsmap)
    
    # 按行分割新内容
    lines = new_content.split('\n')
    
    # 为每个段落处理内容
    for p_index, p_elem in enumerate(p_elems):
        if p_index >= len(lines):
            break
            
        line = lines[p_index]
        
        # 按制表符分割内容，但保留制表符本身
        # 我们需要特殊处理制表符，因为它们在XML中是独立的元素
        parts = []
        current_part = ""
        for char in line:
            if char == '\t':
                # 遇到制表符，保存当前部分并添加制表符标记
                parts.append(current_part)
                parts.append('\t')  # 制表符标记
                current_part = ""
            else:
                current_part += char
        parts.append(current_part)  # 添加最后一部分
        
        # 获取段落中的所有运行（run）
        r_elems = p_elem.xpath(".//w:r", namespaces=nsmap)
        
        # 记录当前处理到哪个部分
        part_index = 0
        
        # 遍历所有运行
        for r_elem in r_elems:
            # 检查这个运行是否包含制表符（普通制表符或位置制表符）
            tab_elems = r_elem.xpath(".//w:tab", namespaces=nsmap)
            ptab_elems = r_elem.xpath(".//w:ptab", namespaces=nsmap)
            
            # 检查这个运行是否在页码域内
            is_in_page_field = False
            for t_elem in r_elem.xpath(".//w:t", namespaces=nsmap):
                parent = t_elem.getparent()
                while parent is not None:
                    if parent.tag.endswith("instrText") and parent.text and "PAGE" in parent.text:
                        is_in_page_field = True
                        break
                    parent = parent.getparent()
            
            # 如果不是页码域
            if not is_in_page_field:
                if tab_elems or ptab_elems:
                    # 这是一个制表符运行（普通制表符或位置制表符）
                    # 查找parts中下一个制表符标记
                    while part_index < len(parts) and parts[part_index] != '\t':
                        part_index += 1
                    
                    if part_index < len(parts) and parts[part_index] == '\t':
                        # 这个运行应该保持为制表符运行
                        # 确保至少有一个制表符元素存在（保持原有类型）
                        if not tab_elems and not ptab_elems:
                            # 如果没有制表符元素，创建一个普通制表符
                            etree.SubElement(r_elem, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab")
                        # 清除其他可能的文本内容
                        t_elems = r_elem.xpath(".//w:t", namespaces=nsmap)
                        for t_elem in t_elems:
                            t_elem.getparent().remove(t_elem)
                        part_index += 1
                else:
                    # 这是一个文本运行，更新文本内容
                    t_elems = r_elem.xpath(".//w:t", namespaces=nsmap)
                    if t_elems and part_index < len(parts):
                        # 跳过制表符标记，找到下一个文本部分
                        while part_index < len(parts) and parts[part_index] == '\t':
                            part_index += 1
                        
                        if part_index < len(parts):
                            # 保留原有文本元素的属性，只更新文本内容
                            t_elems[0].text = parts[part_index]
                            part_index += 1
                    elif t_elems:
                        # 如果没有更多内容，清空文本
                        t_elems[0].text = ""


def replace_header_footer_content(source_docx_path: str, target_docx_path: str, 
                                 source_section_index: int = 1, target_section_index: int = 1,
                                 save_path: str = "") -> bool:
    """
    将源文档的页眉页脚内容替换到目标文档的指定节中，保持页码和格式不变
    
    Args:
        source_docx_path (str): 源DOCX文件路径（提供页眉页脚内容）
        target_docx_path (str): 目标DOCX文件路径（被替换页眉页脚内容）
        source_section_index (int): 源文档节索引（从1开始）
        target_section_index (int): 目标文档节索引（从1开始）
        save_path (str): 保存路径，如果为None则覆盖目标文件
        
    Returns:
        bool: 操作是否成功
    """
    print(f"🔄 开始替换页眉页脚内容")
    print(f"  源文档: {source_docx_path} (第{source_section_index}节)")
    print(f"  目标文档: {target_docx_path} (第{target_section_index}节)")
    
    try:
        # 提取源文档的页眉页脚内容
        source_content = extract_header_footer_content(source_docx_path, source_section_index)
        
        # 解压源文档到内存，获取段落属性
        with zipfile.ZipFile(source_docx_path, 'r') as zin:
            source_filelist = zin.namelist()
            source_files = {name: zin.read(name) for name in source_filelist}
        
        # 解压目标文档到内存
        with zipfile.ZipFile(target_docx_path, 'r') as zin:
            filelist = zin.namelist()
            files = {name: zin.read(name) for name in filelist}

        # 找到源文档的 document.xml，定位节与 header/footer 的对应关系
        source_document_root = etree.fromstring(source_files["word/document.xml"])
        source_nsmap = source_document_root.nsmap
        source_sects = source_document_root.xpath(".//w:sectPr", namespaces=source_nsmap)
        
        if len(source_sects) < source_section_index:
            raise ValueError(f"源文档共有 {len(source_sects)} 节，不能操作第 {source_section_index} 节")

        source_sect = source_sects[source_section_index - 1]
        source_header_refs = source_sect.xpath("./w:headerReference", namespaces=source_nsmap)
        source_footer_refs = source_sect.xpath("./w:footerReference", namespaces=source_nsmap)
        
        # 找到目标文档的 document.xml，定位节与 header/footer 的对应关系
        document_root = etree.fromstring(files["word/document.xml"])
        nsmap = document_root.nsmap
        sects = document_root.xpath(".//w:sectPr", namespaces=nsmap)
        
        print(f"📊 目标文档找到 {len(sects)} 个节")

        if len(sects) < target_section_index:
            raise ValueError(f"目标文档共有 {len(sects)} 节，不能操作第 {target_section_index} 节")

        target_sect = sects[target_section_index - 1]
        
        # 查找该节的 headerReference 和 footerReference
        header_refs = target_sect.xpath("./w:headerReference", namespaces=nsmap)
        footer_refs = target_sect.xpath("./w:footerReference", namespaces=nsmap)
        
        print(f"📋 找到 {len(header_refs)} 个页眉引用, {len(footer_refs)} 个页脚引用")
        
        # 创建一个字典来存储源文档的段落属性
        source_header_properties = {}
        source_footer_properties = {}
        
        # 提取源文档页眉的段落属性
        for i, href in enumerate(source_header_refs):
            rid = href.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            header_type = href.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 headerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in source_files:
                rels_root = etree.fromstring(source_files[rels_name])
                header_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if header_target:
                    header_file = "word/" + header_target[0].get("Target")
                    if header_file in source_files:
                        hroot = etree.fromstring(source_files[header_file])
                        header_nsmap = hroot.nsmap
                        
                        # 提取段落属性
                        p_elems = hroot.xpath(".//w:p", namespaces=header_nsmap)
                        properties = []
                        for j, p_elem in enumerate(p_elems):
                            p_pr_elems = p_elem.xpath("./w:pPr", namespaces=header_nsmap)
                            if p_pr_elems:
                                p_pr_elem = p_pr_elems[0]
                                # 提取对齐方式
                                jc_elems = p_pr_elem.xpath("./w:jc", namespaces=header_nsmap)
                                jc_val = jc_elems[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if jc_elems else None
                                # 提取制表符停止点
                                tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=header_nsmap)
                                tabs_xml = etree.tostring(tabs_elems[0], encoding="unicode") if tabs_elems else None
                                properties.append({
                                    "jc_val": jc_val,
                                    "tabs_xml": tabs_xml
                                })
                        source_header_properties[header_type] = properties
        
        # 提取源文档页脚的段落属性
        for i, fref in enumerate(source_footer_refs):
            rid = fref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            footer_type = fref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 footerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in source_files:
                rels_root = etree.fromstring(source_files[rels_name])
                footer_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if footer_target:
                    footer_file = "word/" + footer_target[0].get("Target")
                    if footer_file in source_files:
                        froot = etree.fromstring(source_files[footer_file])
                        footer_nsmap = froot.nsmap
                        
                        # 提取段落属性
                        p_elems = froot.xpath(".//w:p", namespaces=footer_nsmap)
                        properties = []
                        for j, p_elem in enumerate(p_elems):
                            p_pr_elems = p_elem.xpath("./w:pPr", namespaces=footer_nsmap)
                            if p_pr_elems:
                                p_pr_elem = p_pr_elems[0]
                                # 提取对齐方式
                                jc_elems = p_pr_elem.xpath("./w:jc", namespaces=footer_nsmap)
                                jc_val = jc_elems[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if jc_elems else None
                                # 提取制表符停止点
                                tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=footer_nsmap)
                                tabs_xml = etree.tostring(tabs_elems[0], encoding="unicode") if tabs_elems else None
                                properties.append({
                                    "jc_val": jc_val,
                                    "tabs_xml": tabs_xml
                                })
                        source_footer_properties[footer_type] = properties
        
        # 替换页眉内容
        for i, href in enumerate(header_refs):
            rid = href.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            header_type = href.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 headerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                header_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if header_target:
                    header_file = "word/" + header_target[0].get("Target")
                    if header_file in files:
                        hroot = etree.fromstring(files[header_file])
                        header_nsmap = hroot.nsmap  # 使用页眉文件的命名空间映射
                        
                        # 如果源文档有对应类型的页眉内容，则替换
                        if header_type in source_content["headers"]:
                            header_text = source_content["headers"][header_type]
                            if header_text:  # 只有当有内容时才替换
                                # 保留原有结构，只替换非页码域的文本内容
                                replace_formatted_content(hroot, header_text, header_nsmap)
                                print(f"  🔄 替换页眉 ({header_type}): {header_text}")
                        
                        # 应用源文档的段落属性
                        if header_type in source_header_properties:
                            p_elems = hroot.xpath(".//w:p", namespaces=header_nsmap)
                            source_properties = source_header_properties[header_type]
                            for j, p_elem in enumerate(p_elems):
                                if j < len(source_properties):
                                    p_pr_elems = p_elem.xpath("./w:pPr", namespaces=header_nsmap)
                                    if p_pr_elems:
                                        p_pr_elem = p_pr_elems[0]
                                        source_prop = source_properties[j]
                                        
                                        # 应用对齐方式
                                        if source_prop["jc_val"] is not None:
                                            jc_elems = p_pr_elem.xpath("./w:jc", namespaces=header_nsmap)
                                            if jc_elems:
                                                jc_elem = jc_elems[0]
                                            else:
                                                jc_elem = etree.SubElement(p_pr_elem, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
                                            jc_elem.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", source_prop["jc_val"])
                                        
                                        # 应用制表符停止点
                                        if source_prop["tabs_xml"] is not None:
                                            # 删除现有的制表符停止点
                                            tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=header_nsmap)
                                            for tabs_elem in tabs_elems:
                                                tabs_elem.getparent().remove(tabs_elem)
                                            # 添加源文档的制表符停止点
                                            try:
                                                tabs_elem = etree.fromstring(source_prop["tabs_xml"])
                                                p_pr_elem.append(tabs_elem)
                                            except Exception as e:
                                                print(f"  ⚠️ 应用制表符停止点时出错: {e}")
                        
                        files[header_file] = etree.tostring(hroot, xml_declaration=True, encoding="UTF-8", standalone="yes")

        # 替换页脚内容
        for i, fref in enumerate(footer_refs):
            rid = fref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            footer_type = fref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type", "default")
            
            # 根据关系文件找到 footerX.xml
            rels_name = "word/_rels/document.xml.rels"
            if rels_name in files:
                rels_root = etree.fromstring(files[rels_name])
                footer_target = rels_root.xpath(f".//rel:Relationship[@Id='{rid}']",
                                                namespaces={"rel": "http://schemas.openxmlformats.org/package/2006/relationships"})
                if footer_target:
                    footer_file = "word/" + footer_target[0].get("Target")
                    if footer_file in files:
                        froot = etree.fromstring(files[footer_file])
                        footer_nsmap = froot.nsmap  # 使用页脚文件的命名空间映射
                        
                        # 如果源文档有对应类型的页脚内容，则替换
                        if footer_type in source_content["footers"]:
                            footer_text = source_content["footers"][footer_type]
                            if footer_text:  # 只有当有内容时才替换
                                # 保留原有结构，只替换非页码域的文本内容
                                replace_formatted_content(froot, footer_text, footer_nsmap)
                                print(f"  🔄 替换页脚 ({footer_type}): {footer_text}")
                        
                        # 应用源文档的段落属性
                        if footer_type in source_footer_properties:
                            p_elems = froot.xpath(".//w:p", namespaces=footer_nsmap)
                            source_properties = source_footer_properties[footer_type]
                            for j, p_elem in enumerate(p_elems):
                                if j < len(source_properties):
                                    p_pr_elems = p_elem.xpath("./w:pPr", namespaces=footer_nsmap)
                                    if p_pr_elems:
                                        p_pr_elem = p_pr_elems[0]
                                        source_prop = source_properties[j]
                                        
                                        # 应用对齐方式
                                        if source_prop["jc_val"] is not None:
                                            jc_elems = p_pr_elem.xpath("./w:jc", namespaces=footer_nsmap)
                                            if jc_elems:
                                                jc_elem = jc_elems[0]
                                            else:
                                                jc_elem = etree.SubElement(p_pr_elem, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc")
                                            jc_elem.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", source_prop["jc_val"])
                                        
                                        # 应用制表符停止点
                                        if source_prop["tabs_xml"] is not None:
                                            # 删除现有的制表符停止点
                                            tabs_elems = p_pr_elem.xpath("./w:tabs", namespaces=footer_nsmap)
                                            for tabs_elem in tabs_elems:
                                                tabs_elem.getparent().remove(tabs_elem)
                                            # 添加源文档的制表符停止点
                                            try:
                                                tabs_elem = etree.fromstring(source_prop["tabs_xml"])
                                                p_pr_elem.append(tabs_elem)
                                            except Exception as e:
                                                print(f"  ⚠️ 应用制表符停止点时出错: {e}")
                        
                        files[footer_file] = etree.tostring(froot, xml_declaration=True, encoding="UTF-8", standalone="yes")

        # 保存修改后的文档
        output_path = save_path if save_path else target_docx_path
        output_dir = os.path.dirname(os.path.abspath(output_path))
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        # 先尝试删除目标文件
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception as e:
                print(f"  ⚠️ 删除原文件失败: {e}")
        
        # 写入新文件
        with zipfile.ZipFile(output_path, 'w') as zout:
            for name, data in files.items():
                zout.writestr(name, data)
                print(f"  📄 更新文件: {name}")

        print(f"✅ 页眉页脚内容替换完成，文件保存到: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ 替换页眉页脚内容失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主函数 - 提供命令行接口"""
    print("🚀 DOCX页眉页脚内容提取与替换工具")
    print("=" * 50)
    print("使用方法:")
    print("1. extract_header_footer_content(docx_path, section_index)")
    print("   - 提取指定节的页眉页脚内容")
    print("2. replace_header_footer_content(source_docx_path, target_docx_path, source_section_index, target_section_index, save_path)")
    print("   - 将源文档的页眉页脚内容替换到目标文档")
    print("=" * 50)


if __name__ == "__main__":
    main()