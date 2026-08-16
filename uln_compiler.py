import os
import sys
from typing import Dict, Any, Optional
from uln_parser import ULNParser
from uln_renderer import ULNWordRenderer

try:
    import pythoncom
    import win32com.client
    pywin32_available = True
except ImportError:
    pywin32_available = False

class ULNCompiler:
    """
    Main Compiler module for Universal Layout Notation (ULN) to DOCX using pywin32 COM automation.
    """
    def __init__(self, settings: Optional[Dict[str, Any]] = None):
        self.settings = settings or {}
        self.renderer = ULNWordRenderer(self.settings)

    def compile(self, uln_text: str, output_filepath: Optional[str] = None, visible: bool = False, keep_open: bool = False, background_mode: bool = False) -> str:
        """
        Compiles raw ULN plain text into a formatted MS Word (.docx) file.
        Saves directly to the specified output_filepath.
        """
        if not pywin32_available:
            raise RuntimeError("pywin32 package is not installed or available on this Python runtime.")

        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if not output_filepath:
            output_filepath = os.path.abspath(f"uln_document_{timestamp}.docx")
        else:
            output_filepath = os.path.abspath(output_filepath)

        abs_output_path = output_filepath
        parent_dir = os.path.dirname(abs_output_path)
        if parent_dir:
            os.makedirs(parent_dir, exist_ok=True)

        blocks = ULNParser.parse(uln_text)
        word = None
        doc = None

        try:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

            word = win32com.client.Dispatch("Word.Application")
            is_live_view = (visible or keep_open) and not background_mode
            word.Visible = True if is_live_view else False
            word.DisplayAlerts = 0  # wdAlertsNone
            
            try:
                word.ScreenUpdating = True if is_live_view else False
            except Exception:
                pass

            doc = word.Documents.Add()
            if is_live_view:
                try:
                    word.Activate()
                except Exception:
                    pass
            
            # Execute pywin32 rendering
            self.renderer.render(blocks, doc, word)

            # Save as DOCX (FileFormat = 16)
            try:
                doc.SaveAs2(abs_output_path, FileFormat=16)
            except Exception:
                # If target filename is currently open/locked, save with a timestamp suffix
                import time
                base, ext = os.path.splitext(abs_output_path)
                fallback_path = f"{base}_{int(time.time())}{ext}"
                try:
                    doc.SaveAs2(fallback_path, FileFormat=16)
                    abs_output_path = fallback_path
                except Exception as save_err:
                    print(f"[ULNCompiler] Could not save DOCX file: {save_err}")

            if visible or keep_open:
                try:
                    doc.Saved = True
                    word.ScreenUpdating = True
                    word.Visible = True
                    try:
                        word.Selection.Collapse(0)
                    except Exception:
                        pass
                    try:
                        word.ActiveWindow.View.Type = 3  # wdPrintView = 3
                    except Exception:
                        pass
                    word.Activate()
                except Exception:
                    pass

            return abs_output_path

        finally:
            if not keep_open:
                if doc:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                if word:
                    try:
                        word.Quit()
                    except Exception:
                        pass
            else:
                # Fully detach COM object references so Word operates as a standalone user window
                try:
                    doc.Saved = True
                except Exception:
                    pass
                doc = None
                word = None
                word = None

            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def compile_file(self, input_filepath: str, output_filepath: str, visible: bool = False) -> str:
        """Loads a .txt ULN file and compiles it to DOCX."""
        with open(input_filepath, "r", encoding="utf-8") as f:
            uln_text = f.read()
        return self.compile(uln_text, output_filepath, visible=visible)

    def convert_docx_to_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """Converts DOCX to PDF using pywin32 COM (Word native engine)."""
        if not pywin32_available:
            return False

        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(pdf_path)

        word = None
        doc = None
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(abs_docx)
            doc.SaveAs2(abs_pdf, FileFormat=17)  # wdFormatPDF = 17
            return os.path.exists(abs_pdf)
        except Exception as e:
            print(f"[ULNCompiler] PDF export failed: {e}")
            return False
        finally:
            if doc:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
