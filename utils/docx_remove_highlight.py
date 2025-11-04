# remove_highlight_lxml_simple.py
"""
功能：
    通过 lxml 删除 Word 文件中的所有高亮（w:highlight、w:shd、w:color）
用法：
    直接修改下方 INPUT_PATH = "你的Word文件路径"
    然后运行本脚本即可生成清除高亮后的副本
"""

import zipfile
from pathlib import Path
from lxml import etree

# ========= 用户配置部分 =========
INPUT_PATH = r"C:\Users\yanha\Desktop\新建 DOCX 文档 - 副本 - 副本 (4) - 副本.docx"  # ← 修改为你的文件路径
OUTPUT_PATH = None  # 默认为输入文件同目录下 example_no_highlight_lxml.docx
REMOVE_COLOR_NODE = True  # 若想更彻底移除颜色节点，改为 True
# =================================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}
XML_PARSER = etree.XMLParser(ns_clean=True, recover=True, remove_blank_text=False)

def process_xml_bytes(data: bytes, remove_color_node=False) -> bytes:
    """删除 w:highlight、w:shd，并处理 w:color"""
    try:
        root = etree.fromstring(data, parser=XML_PARSER)
    except Exception:
        return data

    changed = False
    # 删除 highlight
    for node in root.xpath('.//w:highlight', namespaces=NSMAP):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
            changed = True

    # 删除底纹
    for node in root.xpath('.//w:shd', namespaces=NSMAP):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
            changed = True

    # 处理颜色
    for color in root.xpath('.//w:color', namespaces=NSMAP):
        val = color.get("val")
        if val is not None and val.lower() != "auto":
            if remove_color_node:
                parent = color.getparent()
                if parent is not None:
                    parent.remove(color)
                    changed = True
            else:
                color.set("val", "auto")
                changed = True

    if not changed:
        return data
    return etree.tostring(root, encoding="utf-8", xml_declaration=True)

def remove_highlight_from_docx(src_path: str, dest_path: str = None, remove_color_node=False):
    src = Path(src_path)
    if not src.exists():
        raise FileNotFoundError(f"文件不存在: {src}")

    if dest_path is None:
        dest = src.with_name(src.stem + "_no_highlight_lxml" + src.suffix)
    else:
        dest = Path(dest_path)

    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dest, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for name in zin.namelist():
                data = zin.read(name)
                if name.startswith("word/") and name.endswith(".xml"):
                    try:
                        new_data = process_xml_bytes(data, remove_color_node=remove_color_node)
                        zout.writestr(name, new_data)
                    except Exception as e:
                        print(f"⚠ 处理 {name} 出错，保留原文件。错误：{e}")
                        zout.writestr(name, data)
                else:
                    zout.writestr(name, data)

    print("✅ 输出文件：", dest)
    print("提示：可在 Word 中打开后按 F9 更新目录。")


remove_highlight_from_docx(INPUT_PATH, OUTPUT_PATH, REMOVE_COLOR_NODE)
