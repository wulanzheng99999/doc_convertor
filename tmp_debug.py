from pathlib import Path
text = Path(r'utils/merge_docs_preserve_orientation.py').read_text(encoding='utf-8')
old = '''    finally:\n        try:\n            if doc_base is not None:\n                doc_base.Close(SaveChanges=wdDoNotSaveChanges)\n        except Exception as e_close:\n            print(f"\u5173\u95ed\u6587\u6863\u65f6\u51fa\u9519: {e_close}")\n        try:\n            if word is not None:\n                word.Quit()\n        except Exception as e_quit:\n            print(f"\u9000\u51faWord\u65f6\u51fa\u9519: {e_quit}")\n        try:\n            if pythoncom_initialized and pythoncom is not None:\n                pythoncom.CoUninitialize()\n        except Exception as e_co:\n            print(f"CoUninitialize \u65f6\u51fa\u9519: {e_co}")\n'''
print('old in text?', old in text)
print('old unicode:', old.encode('unicode_escape'))
start = text.find('    finally:')
slice_text = text[start:start+200]
print('slice unicode:', slice_text.encode('unicode_escape'))
