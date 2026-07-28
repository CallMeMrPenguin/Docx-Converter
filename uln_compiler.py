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
        self.output_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "output"))
        os.makedirs(self.output_dir, exist_ok=True)

    def compile(self, uln_text: str, output_filepath: Optional[str] = None, visible: bool = False, keep_open: bool = False) -> str:
        """
        Compiles raw ULN plain text into a formatted MS Word (.docx) file.
        Returns absolute path to generated file in output/ directory by default.
        """
        if not pywin32_available:
            raise RuntimeError("pywin32 package is not installed or available on this Python runtime.")

        blocks = ULNParser.parse(uln_text)

        if not output_filepath:
            output_filepath = os.path.join(self.output_dir, "uln_document.docx")
        elif not os.path.isabs(output_filepath):
            if not output_filepath.startswith("output"):
                output_filepath = os.path.join(self.output_dir, output_filepath)
            else:
                output_filepath = os.path.abspath(output_filepath)

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
            word.Visible = True if (visible or keep_open) else False
            word.DisplayAlerts = 0  # wdAlertsNone

            doc = word.Documents.Add()
            
            # Execute pywin32 rendering in real-time
            self.renderer.render(blocks, doc, word)

            # Save as DOCX (FileFormat = 16)
            doc.SaveAs2(abs_output_path, FileFormat=16)

            if keep_open:
                try:
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
                # Detach COM references so Word runs independently as a user window
                doc = None
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
