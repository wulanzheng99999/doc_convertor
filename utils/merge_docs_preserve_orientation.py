import gc
import os
import time
import contextlib


def append_file_with_formatting(base_file, file_to_append, output_file):
    """
    【最终版逻辑 - V4 修复】
    打开 base_file (基础文件), 在其末尾追加 file_to_append (待追加文件)。

    此版本修复了 V3 中会将基础文件最后一节也设置为横向的 bug。
    现在它【只】遍历新插入的节，并强制将它们的页面方向(Orientation)设置为横向。
    """
    base_file_path = os.path.abspath(base_file)
    file_to_append_path = os.path.abspath(file_to_append)
    output_file_path = os.path.abspath(output_file)

    if not os.path.exists(base_file_path):
        raise FileNotFoundError(f"基础文件不存在: {base_file_path}")
    if not os.path.exists(file_to_append_path):
        raise FileNotFoundError(f"待追加文件不存在: {file_to_append_path}")

    word = None
    doc_base = None  # 基础文档
    pythoncom = None
    pythoncom_initialized = False

    try:
        import pythoncom
        import win32com.client as win32
        from win32com.client import VARIANT, constants, gencache
        from pywintypes import com_error

        # 预加载类型库，减少 late binding 风险
        gencache.EnsureDispatch("Word.Application")

        # --- Win32 COM 常量 ---
        wdSectionBreakNextPage = getattr(constants, "wdSectionBreakNextPage", 2)
        wdDoNotSaveChanges = getattr(constants, "wdDoNotSaveChanges", 0)
        msoAutomationSecurityForceDisable = getattr(constants, "msoAutomationSecurityForceDisable", 3)
        wdCollapseEnd = getattr(constants, "wdCollapseEnd", 0)

        # 页面方向常量
        wdOrientLandscape = getattr(constants, "wdOrientLandscape", 1)  # 1 = 横向

        # 页眉页脚类型常量
        wdHeaderFooterPrimary = getattr(constants, "wdHeaderFooterPrimary", 1)
        wdHeaderFooterFirstPage = getattr(constants, "wdHeaderFooterFirstPage", 2)
        wdHeaderFooterEvenPages = getattr(constants, "wdHeaderFooterEvenPages", 3)

        pythoncom.CoInitialize()
        pythoncom_initialized = True

        def _create_word_app():
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                app.AutomationSecurity = msoAutomationSecurityForceDisable
            except Exception:
                pass
            with contextlib.suppress(Exception):
                app.Options.SaveNormalPrompt = False
                app.Options.SavePropertiesPrompt = False
            return app

        def _variant_i4(value):
            """包装为 I4 Variant，降低 late-binding 失败概率。"""
            try:
                return VARIANT(pythoncom.VT_I4, value)
            except Exception:
                return value

        word = _create_word_app()

        base_open_kwargs = dict(
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            PasswordDocument="",
            PasswordTemplate="",
            Revert=False,
            WritePasswordDocument="",
            WritePasswordTemplate="",
            Format=0,
            Encoding=0,
            Visible=False,
            NoEncodingDialog=True,
        )

        repair_open_kwargs = dict(base_open_kwargs)
        repair_open_kwargs["OpenAndRepair"] = True

        def _open_with_retry(path, label):
            nonlocal word
            last_err = None
            for attempt in range(2):
                restart_needed = False
                documents = getattr(word, "Documents", None)
                if documents is None:
                    raise RuntimeError("Word COM \u5bf9\u8c61\u7f3a\u5c11 Documents \u96c6\u5408")
                try:
                    open_no_repair = getattr(documents, "OpenNoRepairDialog", None)
                    strategies = [
                        ("normal", documents.Open, base_open_kwargs),
                        ("repair", documents.Open, repair_open_kwargs),
                    ]
                    if open_no_repair:
                        strategies.append(("open_no_repair_dialog", open_no_repair, repair_open_kwargs))

                    for mode_name, opener, kwargs in strategies:
                        try:
                            return opener(FileName=path, **kwargs)
                        except com_error as err:
                            last_err = err
                            err_code = getattr(err, "hresult", None) or err.args[0]
                            print(f"\u26a0\ufe0f Word \u6253\u5f00{label}\u5931\u8d25\uff08\u6a21\u5f0f {mode_name}, \u5c1d\u8bd5 {attempt + 1}/2\uff09: {err}")

                            if err_code == -2147417836 and attempt < 1:
                                restart_needed = True
                                try:
                                    word.Quit()
                                except Exception:
                                    pass
                                time.sleep(0.5)
                                word = _create_word_app()
                                break
                            continue
                finally:
                    documents = None
                if restart_needed:
                    continue
                break
            if last_err:
                raise last_err
            raise RuntimeError(f"未能打开文档 {path}")

        # 1. 打开基础文件 (竖版)
        print(f"正在打开基础文档: {base_file_path}")
        doc_base = _open_with_retry(base_file_path, "基础文档(竖版)")

        # 2. 记录插入前的节(Section)数量
        sections_collection = getattr(doc_base, "Sections", None)
        if sections_collection is None:
            raise RuntimeError("无法通过 COM 访问基础文档的 Sections 集合")
        sections_before = sections_collection.Count
        print(f"基础文档包含 {sections_before} 个节。")

        # 3. 移动到文档末尾并插入分节符
        print("移动到文档末尾并插入分节符...")
        end_range = doc_base.Content
        end_range.Collapse(wdCollapseEnd)
        end_range.InsertBreak(Type=wdSectionBreakNextPage)

        # 此时，文档的节数量变为 sections_before + 1

        # 4. 在新节的开头插入横版文件
        print(f"正在末尾追加文件(横版): {file_to_append_path}")
        end_range.InsertFile(
            FileName=file_to_append_path,
            Link=False,
            ConfirmConversions=False
        )
        end_range = None

        # 5. 记录插入后的节(Section)数量
        sections_after = sections_collection.Count
        print(f"插入后文档变为 {sections_after} 个节。")

        # -------------------【关键修复逻辑 V4】-------------------
        # 6. 遍历所有新添加的节 (从 sections_before + 1 到末尾)
        #    注意：COM 集合的索引从 1 开始

        # 【修正点】
        # 我们只从插入分节符后创建的第一个新节开始
        start_section_index = sections_before + 1

        print(f"将从第 {start_section_index} 节开始强制设置为横向...")

        for i in range(start_section_index, sections_after + 1):
            new_section = None
            try:
                try:
                    new_section = sections_collection(_variant_i4(i))
                except Exception:
                    new_section = sections_collection(i)

                page_setup = getattr(new_section, "PageSetup", None)
                if page_setup is not None:
                    page_setup.Orientation = wdOrientLandscape
                print(f"  > 已强制第 {i} 节为横向。")
                page_setup = None

                headers = getattr(new_section, "Headers", None)
                if headers is not None:
                    for header_type in (wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages):
                        try:
                            header = headers(header_type)
                            header.LinkToPrevious = False
                        except Exception:
                            continue
                        finally:
                            header = None
                headers = None
                footers = getattr(new_section, "Footers", None)
                if footers is not None:
                    for footer_type in (wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages):
                        try:
                            footer = footers(footer_type)
                            footer.LinkToPrevious = False
                        except Exception:
                            continue
                        finally:
                            footer = None
                footers = None
                print(f"  > 已断开第 {i} 节的页眉/页脚链接。")

            except Exception as e:
                print(f"  > 处理第 {i} 节时出现非致命错误(已忽略): {e}")
            finally:
                new_section = None
        sections_collection = None
        # -------------------【修复结束】-------------------

        print(f"正在保存到: {output_file_path}")
        doc_base.SaveAs(output_file_path)
        print("合并成功！")
        return True

    except Exception as e:
        print(f"操作失败: {e}")
        try:
            size_a = os.path.getsize(base_file_path)
        except Exception:
            size_a = "unknown"
        try:
            size_b = os.path.getsize(file_to_append_path)
        except Exception:
            size_b = "unknown"
        print(
            f"调试信息 -> 基础文件: {base_file_path} (size: {size_a}), 追加文件: {file_to_append_path} (size: {size_b})")
        import traceback
        traceback.print_exc()
        return False

    finally:
        if doc_base is not None:
            try:
                doc_base.Close(SaveChanges=wdDoNotSaveChanges)
            except Exception as e_close:
                print(f"关闭文档时出错: {e_close}")
            finally:
                doc_base = None
        if word is not None:
            try:
                with contextlib.suppress(Exception):
                    word.NormalTemplate.Saved = True
                word.Quit()
            except Exception as e_quit:
                print(f"退出Word时出错: {e_quit}")
            finally:
                word = None
        if pythoncom_initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception as e_co:
                print(f"CoUninitialize 时出错: {e_co}")
            finally:
                pythoncom_initialized = False
                pythoncom = None
        try:
            gc.collect()
        except Exception:
            pass


if __name__ == "__main__":
    # --- 您可以修改这里的路径 ---

    # 横版文件
    LANDSCAPE_FILE = r"D:\storge\input\step1_5_split_横版内容_20251027_190928.docx"

    # 竖版文件
    PORTRAIT_FILE = r"D:\storge\input\step4_final_最终文档_20251027_204601.docx"

    # 输出文件
    FINAL_OUTPUT = r"D:\storge\input\merged_竖版在先_横版在后_V4_FIXED.docx"  # 使用新名字

    print("开始合并 (V4 逻辑：竖版 + 横版 + 仅新增部分强制横向)...")
    append_file_with_formatting(
        base_file=PORTRAIT_FILE,  # 基础文件 (竖)
        file_to_append=LANDSCAPE_FILE,  # 追加文件 (横)
        output_file=FINAL_OUTPUT
    )
    print("操作完成。")
