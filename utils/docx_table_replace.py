# -*- coding: utf-8 -*-
"""
将“原文件”中的所有表格按在正文中的位置顺序，
替换到“被修改的文件”的相同位置上。

依赖：
    pip install python-docx

使用方法：
    1) 修改下方路径常量 ORIGINAL_DOC_PATH / EDITED_DOC_PATH / OUTPUT_DOC_PATH
    2) 运行脚本：python replace_tables_by_position.py
"""

import copy
from pathlib import Path
from typing import Any, Iterator

# 导入Document类（忽略静态分析工具的警告）
# pyright: reportAttributeAccessIssue=false
try:
    from docx import Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    DOCX_AVAILABLE = True
except ImportError:
    Document = None
    CT_Tbl = type(None)  # 使用一个实际的类型而不是None
    CT_P = type(None)    # 使用一个实际的类型而不是None
    DOCX_AVAILABLE = False

# ====== 路径配置（按需修改）======
ORIGINAL_DOC_PATH = "input/test(1).docx"   # “原文件”（想要抽取表格的文件）
EDITED_DOC_PATH   = "input/step2_pandoc_转换成功_20250917_211258.docx"     # “被修改的文件”（要被替换表格的文件）
OUTPUT_DOC_PATH   = "output/replaced.docx"  # 输出文件
# =================================


def iter_body_children(doc) -> Iterator[Any]:
    """
    迭代正文 body 的直接子节点，保持原有顺序。
    只区分段落(CT_P)与表格(CT_Tbl)，其余节点直接返回原 oxml 以免改动顺序。
    """
    if not DOCX_AVAILABLE or doc is None:
        return iter([])
    
    body = doc.element.body
    for child in body.iterchildren():
        yield child


def collect_tables_in_body(doc) -> list:
    """
    按正文顺序收集表格的底层 OOXML 节点 (CT_Tbl)。
    不进入页眉/页脚/文本框/形状等，仅处理 document.body。
    """
    if not DOCX_AVAILABLE or doc is None:
        return []
    
    tables = []
    for child in iter_body_children(doc):
        if DOCX_AVAILABLE and CT_Tbl is not type(None) and isinstance(child, CT_Tbl):
            tables.append(child)
    return tables


def replace_tables_by_index(original_path, edited_path, output_path):
    """
    将原文件中的所有表格按在正文中的位置顺序，替换到被修改的文件的相同位置上。

    Args:
        original_path: 原文件路径（提供表格内容）
        edited_path: 被修改的文件路径（被替换表格内容）
        output_path: 输出文件路径

    Returns:
        bool: 替换是否成功
    """
    # 检查docx库是否可用
    if not DOCX_AVAILABLE or Document is None:
        raise ImportError("python-docx库不可用，请安装python-docx: pip install python-docx")
    
    # 载入文档
    if not Path(original_path).exists() or not Path(edited_path).exists():
        raise FileNotFoundError("请确认 ORIGINAL_DOC_PATH 与 EDITED_DOC_PATH 文件存在。")

    print(f"📄 加载原文件: {original_path}")
    doc_original = Document(original_path)
    print(f"📄 加载被修改文件: {edited_path}")
    doc_edited = Document(edited_path)

    # 收集两者正文中的表格（按出现顺序）
    orig_tables = collect_tables_in_body(doc_original)
    edited_tables = collect_tables_in_body(doc_edited)

    print(f"🔎 原文件表格数量: {len(orig_tables)}")
    print(f"🔎 被修改文件表格数量: {len(edited_tables)}")

    if len(orig_tables) == 0 and len(edited_tables) == 0:
        print("🤷 两个文件里都没有表格，无需处理。")
        doc_edited.save(output_path)
        print(f"💾 已保存（无改动）到: {output_path}")
        return True

    if len(edited_tables) == 0:
        print("⚠️ 被修改文件没有任何表格，无法执行替换。")
        return False

    # 计算要替换的数量（以较小者为准）
    n = min(len(orig_tables), len(edited_tables))
    if len(orig_tables) != len(edited_tables):
        print(f"⚠️ 表格数量不一致，仅替换前 {n} 个。")

    # 执行按位置替换
    body = doc_edited.element.body
    replaced_count = 0

    # 我们需要在 body 层面找到每一个“第 i 个表格”的节点，并做原地替换
    # 做法：遍历 body 的直接子节点，遇到表格就计数，当计数 == i 时进行替换
    def find_i_th_table_and_replace(i, new_tbl_oxml):
        idx = -1
        for pos, child in enumerate(body.iterchildren()):
            if DOCX_AVAILABLE and CT_Tbl is not type(None) and isinstance(child, CT_Tbl):
                idx += 1
                if idx == i:
                    # 在当前位置插入新的表格节点，然后移除旧节点，达到“就地替换”的效果
                    insert_at = list(body).index(child)
                    body.insert(insert_at, copy.deepcopy(new_tbl_oxml))
                    body.remove(child)
                    return True
        return False

    for i in range(n):
        ok = find_i_th_table_and_replace(i, orig_tables[i])
        if ok:
            replaced_count += 1
            print(f"✅ 已替换第 {i+1} 个表格")
        else:
            print(f"❌ 未找到可替换的位置（第 {i+1} 个表格）")

    # 保存结果
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    doc_edited.save(output_path)
    print(f"\n🎉 完成！共替换 {replaced_count}/{n} 个表格。")
    print(f"📁 输出文件: {output_path}")
    return True


if __name__ == "__main__":
    replace_tables_by_index(ORIGINAL_DOC_PATH, EDITED_DOC_PATH, OUTPUT_DOC_PATH)