# -*- coding: utf-8 -*-
"""
Word 表格替换（优化版）：
- 不再使用 doc.CopyStylesFromTemplate()，避免覆盖目标文档的标题/目录/列表等全局样式。
- 仅用 rng.FormattedText = src_tbl.Range.FormattedText 复制表格内容与直接格式。
- 如需保留源表格依赖的“少数自定义样式”，可按需调用 OrganizerCopy 只拷那几项（可选）。

依赖：
    pip install pywin32
运行环境：
    Windows + Microsoft Word（桌面版）
"""

from __future__ import annotations

import contextlib
import shutil
import time
from copy import deepcopy
from pathlib import Path
from typing import Dict, List, Tuple

# --- 配置 (保持不变) ---
TEMPLATE_DOC_PATH = "../temp/step0_preprocess_标题预处理_20251104_235356.docx"
SOURCE_DOC_PATH   = "../temp/step11_library_number_高级格式化完成_20251104_235500.docx"
OUTPUT_DOC_PATH   = "../output/formatted_document——1.docx"
CLEANED_DOC_PATH  = "output/cleaned_document.docx"

try:
    from docx import Document
    from docx.oxml.ns import qn
except Exception:
    Document = None
    qn = None

STYLE_COPY_ENABLED = True
STYLE_TYPES_TO_COPY = {"paragraph", "character", "table"}
STYLE_CUSTOM_ONLY_TYPES = {"paragraph", "character"}
STYLE_ID_EXCLUDE_PREFIXES = ("Heading", "TOC")
STYLE_ID_EXCLUDE: set[str] = set()
STYLE_SOURCE_OVERRIDE: str | None = None

try:
    import pythoncom
except Exception:
    pythoncom = None

import win32com.client as win32
from win32com.client import constants as C

try:
    import pywintypes
except Exception:
    print("⚠️ 未找到 pywintypes，pywin32 可能安装不完整。")
    pywintypes = None

RETRY_MAX = 3


def _abs(p: str) -> str:
    return str(Path(p).expanduser().resolve())


def _need_retry_com_error(e) -> bool:
    hresult = getattr(e, "hresult", None)
    if hresult in (-2147418111, -2147023174, -2147417836, -2147023170):
        return True
    msg = str(e)
    if "远程过程调用失败" in msg or "The remote procedure call failed" in msg:
        return True
    return any(k in msg for k in (
        "被呼叫方拒绝接收调用",
        "RPC 服务器不可用",
        "Call was rejected by callee",
        "The RPC server is unavailable",
    ))


def _pump_messages():
    if not pythoncom:
        return
    with contextlib.suppress(Exception):
        from pythoncom import PumpWaitingMessages

        PumpWaitingMessages()


def _retry_call(fn, *args, **kwargs):
    delay = 0.6
    for attempt in range(1, RETRY_MAX + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            if pywintypes and isinstance(e, pywintypes.com_error) and _need_retry_com_error(e):
                _pump_messages()
                time.sleep(delay)
                delay = min(delay * 1.4 + 0.2, 2.5)
                continue
            raise


def _get_tables_main_list(doc) -> List:
    rng = _retry_call(lambda: doc.StoryRanges(C.wdMainTextStory))
    if rng is None:
        return []
    tables_dispatch = _retry_call(lambda: getattr(rng, "Tables", None))
    if tables_dispatch is None:
        return []
    tables = [t for t in tables_dispatch]
    tables.sort(key=lambda t: t.Range.Start)
    return tables


def _backup_docx_styles(doc_path: str) -> Dict[str, object]:
    if not STYLE_COPY_ENABLED:
        return {}
    if Document is None or qn is None:
        print("⚠️ 未找到 python-docx，跳过样式备份。")
        return {}
    path = Path(doc_path)
    if not path.exists():
        print(f"⚠️ 样式备份失败：未找到文档 {doc_path}")
        return {}

    doc = Document(str(path))
    styles_element = doc.styles.element
    results: Dict[str, object] = {}
    style_nodes = styles_element.findall(qn("w:style"))
    for style in style_nodes:
        style_id = style.get(qn("w:styleId"))
        if not style_id:
            continue
        style_type = style.get(qn("w:type"))
        if STYLE_TYPES_TO_COPY and style_type not in STYLE_TYPES_TO_COPY:
            continue
        if any(style_id.startswith(prefix) for prefix in STYLE_ID_EXCLUDE_PREFIXES):
            continue
        if style_id in STYLE_ID_EXCLUDE:
            continue
        if (
            STYLE_CUSTOM_ONLY_TYPES
            and style_type in STYLE_CUSTOM_ONLY_TYPES
            and style.get(qn("w:customStyle")) != "1"
        ):
            continue
        results[style_id] = deepcopy(style)
    return results


def _restore_docx_styles(doc_path: str, styles: Dict[str, object]):
    if not styles:
        return
    if Document is None or qn is None:
        print("⚠️ 未找到 python-docx，无法回写样式。")
        return
    path = Path(doc_path)
    if not path.exists():
        print(f"⚠️ 样式回写失败：未找到文档 {doc_path}")
        return
    if not _wait_for_file_access(path, mode="read"):
        print(f"⚠️ 样式回写失败：暂时无法读取 {doc_path}")
        return

    doc = Document(str(path))
    styles_element = doc.styles.element
    existing = {
        node.get(qn("w:styleId")): node
        for node in styles_element.findall(qn("w:style"))
        if node.get(qn("w:styleId"))
    }
    for style_id, node in styles.items():
        current = existing.get(style_id)
        if current is not None:
            styles_element.remove(current)
        styles_element.append(deepcopy(node))
    if not _wait_for_file_access(path, mode="write"):
        print(f"⚠️ 样式回写失败：暂时无法写入 {doc_path}")
        return
    doc.save(str(path))


def _prepare_output_document(dst: str, outp: str, retries: int = 5):
    src_path = Path(dst)
    out_path = Path(outp)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    delay = 0.4
    last_error = None
    _wait_for_file_access(src_path, mode="read")

    for attempt in range(1, retries + 1):
        try:
            if out_path.exists():
                if not _wait_for_file_access(out_path, mode="write"):
                    raise PermissionError(f"timeout waiting to access {out_path}")
                with contextlib.suppress(Exception):
                    out_path.unlink()
                    _wait_for_file_access(out_path, mode="write")
            else:
                _wait_for_file_access(out_path, mode="write")
            shutil.copy2(src_path, out_path)
            _wait_for_file_access(out_path, mode="read")
            return
        except (PermissionError, OSError) as exc:
            last_error = exc
            if attempt >= retries:
                raise
            time.sleep(delay)
            delay = min(delay * 1.6, 2.0)
            continue

    if last_error:
        raise last_error


def _wait_for_file_access(path: Path, mode: str, timeout: float = 10.0, interval: float = 0.3) -> bool:
    deadline = time.time() + timeout
    while True:
        try:
            if mode == "read":
                if not path.exists():
                    return True
                with open(path, "rb"):
                    return True
            else:  # write
                with open(path, "a+b"):
                    return True
        except FileNotFoundError:
            if mode == "write":
                return True
        except OSError:
            pass

        if time.time() >= deadline:
            return False
        time.sleep(interval)


def _hard_replace_table(doc_out, src_tbl, dst_tbl) -> bool:
    try:
        src_fmt = _retry_call(lambda: src_tbl.Range.FormattedText)
        start_pos = _retry_call(lambda: dst_tbl.Range.Start)
        _retry_call(dst_tbl.Delete)
        rng = _retry_call(lambda: doc_out.Range(Start=start_pos, End=start_pos))

        def _assign():
            rng.FormattedText = src_fmt

        _retry_call(_assign)
        return True
    except Exception as e:
        print(f"    ❌ 硬替换失败：{e}")
        return False


def _create_word_app():
    try:
        app = win32.Dispatch("Word.Application")
    except Exception:
        app = win32.DispatchEx("Word.Application")

    with contextlib.suppress(Exception):
        app.Visible = False
    with contextlib.suppress(Exception):
        app.DisplayAlerts = 0
    with contextlib.suppress(Exception):
        app.AutomationSecurity = 3
        app.Options.CheckSpellingAsYouType = False
        app.Options.CheckGrammarAsYouType = False
    with contextlib.suppress(Exception):
        app.Options.SaveNormalPrompt = False
        app.Options.SavePropertiesPrompt = False
    return app


def _open_document(app, path: str, *, readonly: bool):
    documents = getattr(app, 'Documents', None)
    if documents is None:  # Word instance not ready
        raise AttributeError('Word.Application.Documents 属性不可用')
    return documents.Open(
        path,
        ReadOnly=readonly,
        AddToRecentFiles=False,
        ConfirmConversions=False,
        Revert=False,
        OpenAndRepair=False,
        NoEncodingDialog=True,
    )


def _close_doc(doc, save_changes):
    if not doc:
        return
    with contextlib.suppress(Exception):
        _retry_call(doc.Close, SaveChanges=save_changes)


def _quit_app(app):
    if not app:
        return
    with contextlib.suppress(Exception):
        _retry_call(app.Quit)

def _find_document_by_path(app, expected_path: str):
    try:
        expected = Path(expected_path).resolve()
    except Exception:
        return None
    with contextlib.suppress(Exception):
        count = app.Documents.Count
        for idx in range(1, count + 1):
            doc = app.Documents(idx)
            with contextlib.suppress(Exception):
                if Path(doc.FullName).resolve() == expected:
                    return doc
    return None



def _open_documents_with_restart(src: str, dst: str, outp: str) -> Tuple:
    """
    Open source and target documents with retries; returns (app, doc_src, doc_out).
    """
    last_error = None
    for attempt in range(1, RETRY_MAX + 1):
        app = _create_word_app()
        doc_src = None
        doc_out = None
        try:
            doc_src = _retry_call(lambda: _open_document(app, src, readonly=True))
            if doc_src is None:
                doc_src = _find_document_by_path(app, src)
            if doc_src is None:
                raise ValueError("DOC_OPEN_NONE")

            _prepare_output_document(dst, outp)
            doc_out = _retry_call(lambda: _open_document(app, outp, readonly=False))
            if doc_out is None:
                doc_out = _find_document_by_path(app, outp)
            if doc_out is None:
                raise ValueError("DOC_OPEN_NONE")

            # Skip copying styles from the template so target document keeps its own heading/toc styles.

            return app, doc_src, doc_out

        except Exception as e:
            last_error = e
            should_retry = (
                pywintypes and isinstance(e, pywintypes.com_error) and _need_retry_com_error(e)
            ) or (isinstance(e, ValueError) and str(e) == "DOC_OPEN_NONE")

            _close_doc(doc_out, save_changes=False)
            _close_doc(doc_src, save_changes=False)
            _quit_app(app)

            if should_retry and attempt < RETRY_MAX:
                time.sleep(0.5 * attempt)
                continue
            raise
    raise last_error


def replace_tables_in_mainstory_all(original_path: str, edited_path: str, output_path: str) -> Tuple[bool, int]:
    src = _abs(original_path)
    dst = _abs(edited_path)
    outp = _abs(output_path)

    if not Path(src).exists():
        print(f'⚠️ 原始文档不存在：{src}')
        return False, 0
    if not Path(dst).exists():
        print(f'⚠️ 目标文档不存在：{dst}')
        return False, 0
    Path(outp).parent.mkdir(parents=True, exist_ok=True)

    style_backup: Dict[str, object] = {}
    if STYLE_COPY_ENABLED:
        style_source_path = _abs(STYLE_SOURCE_OVERRIDE) if STYLE_SOURCE_OVERRIDE else src
        style_backup = _backup_docx_styles(style_source_path)

    pythoncom_initialized = False
    if pythoncom:
        pythoncom.CoInitialize()
        pythoncom_initialized = True

    app = None
    doc_src = None
    doc_out = None
    replaced = 0
    success = False
    should_restore_styles = False

    try:
        app, doc_src, doc_out = _open_documents_with_restart(src, dst, outp)

        src_tbls = _get_tables_main_list(doc_src)
        dst_tbls = _get_tables_main_list(doc_out)
        n = min(len(src_tbls), len(dst_tbls))

        print('📋 扫描主文流中的表格…')
        print(f'  · 源（主文流）：{len(src_tbls)} 张表')
        print(f'  · 目标（主文流）：{len(dst_tbls)} 张表')

        if n == 0:
            print('⚠️ 主文流一侧无表格，无需替换。')
            _retry_call(doc_out.Save)
            success = True
            should_restore_styles = True
        else:
            jobs = [(src_tbls[i], dst_tbls[i]) for i in range(n)]
            jobs.sort(key=lambda pair: pair[1].Range.Start, reverse=True)

            print(f'  · 计划替换：{n} 项（按目标表在文中的起点倒序执行）')
            for idx, (s, d) in enumerate(jobs[:8], 1):
                print(f'    - #{idx} 源Start={s.Range.Start} → 目标Start={d.Range.Start}')
            if n > 8:
                print(f'    - …其余 {n - 8} 项')

            for s_tbl, d_tbl in jobs:
                if _hard_replace_table(doc_out, s_tbl, d_tbl):
                    replaced += 1
                    print('      ✓ 成功')
                else:
                    print('      ✗ 失败')

            try:
                _retry_call(doc_out.Save)
            except Exception:
                _retry_call(doc_out.SaveAs, outp, FileFormat=12)

            print(f'[OK] 已保存：{outp}')
            print(f'   - 完成替换：{replaced}/{n}')
            success = True
            should_restore_styles = True

    except Exception as e:
        print(f'❌ 执行出错：{e}')
        import traceback

        traceback.print_exc()
    finally:
        _close_doc(doc_out, save_changes=True)
        _close_doc(doc_src, save_changes=False)
        _quit_app(app)
        if pythoncom_initialized:
            with contextlib.suppress(Exception):
                pythoncom.CoUninitialize()

    if should_restore_styles and style_backup:
        try:
            _restore_docx_styles(outp, style_backup)
            print('🔁 已回写原文自定义样式（python-docx）。')
        except Exception as style_err:
            print(f'⚠️ 样式回写失败：{style_err}')

    return success, replaced

def replace_tables(src_path: str, dst_path: str, out_path: str):
    replace_tables_in_mainstory_all(src_path, dst_path, out_path)


if __name__ == "__main__":
    replace_tables(
        TEMPLATE_DOC_PATH,
        SOURCE_DOC_PATH,
        OUTPUT_DOC_PATH,
    )
