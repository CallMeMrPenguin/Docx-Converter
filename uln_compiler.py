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

    def compile(self, uln_text: str, output_filepath: str, visible: bool = False, keep_open: bool = False) -> str:
        """
        Compiles raw ULN plain text into a formatted MS Word (.docx) file.
        Returns absolute path to generated file.
        If keep_open is True, MS Word remains open on screen after compilation.
        """
        if not pywin32_available:
            raise RuntimeError("pywin32 package is not installed or available on this Python runtime.")

        blocks = ULNParser.parse(uln_text)

        abs_output_path = os.path.abspath(output_filepath)
        os.makedirs(os.path.dirname(abs_output_path), exist_ok=True)

        word = None
        doc = None

        try:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True if keep_open else visible
            word.DisplayAlerts = 0  # wdAlertsNone

            doc = word.Documents.Add()
            
            # Execute pywin32 rendering
            self.renderer.render(blocks, doc, word)

            # Save as DOCX (FileFormat = 16)
            doc.SaveAs2(abs_output_path, FileFormat=16)
            
            if keep_open:
                word.Activate()

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
