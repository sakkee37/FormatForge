"""Word COM 引擎：docx ↔ PDF 转换。"""

import os
import pythoncom
import win32com.client


def docx_to_pdf(input_path, output_path=None):
    """将 Word 文档转为 PDF。"""
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(input_path)
        doc.SaveAs2(output_path, FileFormat=17)
        doc.Close()
        return True, f"✓ 完成：{os.path.basename(output_path)}"
    except Exception as e:
        return False, f"✗ 转换失败：{str(e)}"
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def pdf_to_docx(input_path, output_path=None):
    """将 PDF 转为 Word 文档（需要 Word 2013+）。"""
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".docx"
    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        # Word 打开 PDF 时可能需要指定格式或忽略转换提示
        doc = word.Documents.Open(
            input_path,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        if doc is None:
            return False, "✗ 无法打开 PDF（可能需要 Word 2013+）"
        doc.SaveAs2(output_path, FileFormat=16)  # 16 = wdFormatDocument
        doc.Close()
        return True, f"✓ 完成：{os.path.basename(output_path)}"
    except Exception as e:
        return False, f"✗ 转换失败：{str(e)}"
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
