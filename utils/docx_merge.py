import os
import time


def copy_all_to_beginning(file_a, file_b, output_file):
    """
    复制文件A的全部内容到文件B的开头，并在两份内容之间添加分节符。

    Args:
        file_a: 需要复制的DOCX路径
        file_b: 目标DOCX路径
        output_file: 合并后的输出路径
    """
    file_a = os.path.abspath(file_a)
    file_b = os.path.abspath(file_b)
    output_file = os.path.abspath(output_file)

    if not os.path.exists(file_a):
        raise FileNotFoundError(f"源文件不存在: {file_a}")
    if not os.path.exists(file_b):
        raise FileNotFoundError(f"目标文件不存在: {file_b}")

    word = None
    doc_a = None
    doc_b = None
    pythoncom = None
    pythoncom_initialized = False

    try:
        import pythoncom
        import win32com.client as win32
        from pywintypes import com_error

        pythoncom.CoInitialize()
        pythoncom_initialized = True

        def _create_word_app():
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                # 禁用宏提示，避免隐藏弹窗导致命令失败
                app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass
            return app

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
                documents = word.Documents
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
                        print(f"⚠️ Word 打开{label}失败（模式 {mode_name}, 尝试 {attempt + 1}/2）: {err}")

                        # Word 进程断联，重新拉起实例后整体重试
                        if err_code == -2147417836 and attempt < 1:
                            restart_needed = True
                            try:
                                word.Quit()
                            except Exception:
                                pass
                            time.sleep(0.5)
                            word = _create_word_app()
                            break

                        # 其他错误，继续尝试下一个模式
                        continue

                if restart_needed:
                    continue
                break
            if last_err:
                raise last_err
            raise RuntimeError(f"未能打开文档 {path}")

        doc_a = _open_with_retry(file_a, "封面/目录文档")
        doc_b = _open_with_retry(file_b, "正文文档")

        doc_a.Range().Copy()
        doc_b.Range(0, 0).Paste()

        doc_b.Activate()
        selection = word.Selection
        selection.HomeKey(Unit=6)  # Word 常量: wdStory
        file_a_chars = doc_a.Range().Characters.Count
        if file_a_chars > 0:
            selection.MoveRight(Unit=1, Count=file_a_chars - 1)
        selection.InsertBreak(Type=7)  # wdSectionBreakNextPage

        doc_b.SaveAs(output_file)
        return True

    except Exception as e:
        print(f"操作失败: {e}")
        try:
            size_a = os.path.getsize(file_a)
        except Exception:
            size_a = "unknown"
        try:
            size_b = os.path.getsize(file_b)
        except Exception:
            size_b = "unknown"
        print(f"调试信息 -> file_a: {file_a} (size: {size_a}), file_b: {file_b} (size: {size_b})")
        import traceback

        traceback.print_exc()
        return False

    finally:
        for doc in (doc_a, doc_b):
            try:
                if doc:
                    doc.Close()
            except Exception:
                pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass
        try:
            time.sleep(2)  # 给 Word 一点时间完全退出
        except Exception:
            pass
        try:
            if pythoncom_initialized:
                pythoncom.CoUninitialize()
        except Exception:
            pass


# if __name__ == "__main__":
#     copy_all_to_beginning("cover.docx", "body.docx", "merged.docx")
