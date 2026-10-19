"""
server/app.py - High-Performance Localhost Web Server & REST API.
Serves the modern Flat Minimal Light Web UI and orchestrates background Word COM automation.
"""

import os
import sys
import time
import socket
import tempfile
import threading
import base64
import io
import webbrowser
from typing import List, Dict, Any, Optional

# Add project root to sys.path so core, renderers, and services import seamlessly
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from core.compiler import ULNCompiler, extract_raw_uln, embed_raw_uln_docx, scan_folder_for_uln_docx
from services.image_service import ImageService, SUPPORTED_IMAGE_EXTENSIONS
from services.updater_service import get_current_version, check_for_updates


app = FastAPI(title="ULN to DOCX Web Converter", version=get_current_version())

# Enable CORS for local development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Global State for Background Tasks ────────────────────────────
class CompileManager:
    def __init__(self):
        self.lock = threading.Lock()
        self.is_compiling = False
        self.progress = 0.0
        self.message = "Sẵn sàng"
        self.error: Optional[str] = None
        self.output_path: Optional[str] = None
        self.cancel_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None

    def reset(self):
        self.is_compiling = True
        self.progress = 0.05
        self.message = "Đang khởi tạo tài liệu..."
        self.error = None
        self.output_path = None
        self.cancel_event.clear()

compile_manager = CompileManager()


# ── Request / Response Models ────────────────────────────────────
class CompileRequest(BaseModel):
    uln_text: str
    settings: Dict[str, Any] = {}
    output_filename: Optional[str] = None
    open_word: bool = True
    background_mode: bool = False


class EmbedRequest(BaseModel):
    filepath: str
    uln_text: str


class PromptRequest(BaseModel):
    prompt_text: str


class OpenFileRequest(BaseModel):
    filepath: str


# ── REST API Endpoints ───────────────────────────────────────────
@app.get("/api/status")
def get_status():
    """Returns application version and environment status."""
    return {
        "status": "online",
        "version": get_current_version(),
        "default_folder": os.path.expanduser("~\\Documents") if os.name == "nt" else os.getcwd()
    }


@app.get("/api/sample")
def get_sample_uln():
    """Returns default sample ULN text from uln_test.txt."""
    sample_path = os.path.join(project_root, "uln_test.txt")
    sample_content = ""
    if os.path.exists(sample_path):
        try:
            with open(sample_path, "r", encoding="utf-8") as f:
                sample_content = f.read()
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    else:
        sample_content = (
            "[H1] ĐỀ KIỂM TRA MÔN TIẾNG ANH\n"
            "[P0] Mark the letter A, B, C, or D to indicate the word whose underlined part differs:\n"
            "[P1] 1. A. [c]{u}at \t B. [c]{u}ity \t C. [c]{u}ar \t D. [c]{u}up\n"
            "[P1] 2. A. teach[er]{u} \t B. fath[er]{u} \t C. t[er]{u}m \t D. moth[er]{u}\n"
        )
    return {"sample_text": sample_content, "content": sample_content}


@app.get("/api/prompt")
def get_prompt():
    """Reads AI system prompt from prompt.txt."""
    prompt_path = os.path.join(project_root, "prompt.txt")
    prompt_content = ""
    if os.path.exists(prompt_path):
        try:
            with open(prompt_path, "r", encoding="utf-8") as f:
                prompt_content = f.read()
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    else:
        prompt_content = "You are an expert OCR and Universal Layout Notation (ULN) extraction engine."
    return {"prompt_text": prompt_content, "prompt": prompt_content, "default_prompt": prompt_content}


@app.post("/api/prompt")
def save_prompt(req: PromptRequest):
    """Saves updated AI system prompt to prompt.txt."""
    prompt_path = os.path.join(project_root, "prompt.txt")
    try:
        with open(prompt_path, "w", encoding="utf-8") as f:
            f.write(req.prompt_text)
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/compile")
def start_compilation(req: CompileRequest):
    """Starts background Word COM compilation thread."""
    if compile_manager.is_compiling:
        raise HTTPException(status_code=409, detail="Một tác vụ biên dịch đang chạy. Vui lòng chờ hoặc hủy trước.")

    if not req.uln_text or not req.uln_text.strip():
        raise HTTPException(status_code=400, detail="Mã ULN không được để trống.")

    # Determine output path
    docs_dir = os.path.expanduser("~\\Documents") if os.name == "nt" else os.getcwd()
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    fname = req.output_filename or f"uln_document_{timestamp}.docx"
    if not fname.lower().endswith(".docx"):
        fname += ".docx"
    output_path = os.path.abspath(os.path.join(docs_dir, fname))

    compile_manager.reset()

    def _worker():
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

        def _progress_cb(msg: str, pct: float):
            compile_manager.progress = pct
            compile_manager.message = msg

        try:
            settings = dict(req.settings)
            settings["cancel_event"] = compile_manager.cancel_event
            compiler = ULNCompiler(settings)

            compiled = compiler.compile(
                uln_text=req.uln_text,
                output_filepath=output_path,
                keep_open=req.open_word,
                background_mode=req.background_mode,
                cancel_check=lambda: compile_manager.cancel_event.is_set(),
                progress_callback=_progress_cb
            )
            compile_manager.output_path = compiled
            compile_manager.progress = 1.0
            compile_manager.message = "Xuất file DOCX thành công!"
        except KeyboardInterrupt:
            compile_manager.error = "Tác vụ đã bị người dùng hủy."
            compile_manager.message = "Đã dừng biên dịch."
        except Exception as e:
            compile_manager.error = str(e)
            compile_manager.message = f"Lỗi biên dịch: {e}"
        finally:
            compile_manager.is_compiling = False
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    compile_manager.worker_thread = threading.Thread(target=_worker, daemon=True)
    compile_manager.worker_thread.start()

    return {"success": True, "message": "Đã bắt đầu tác vụ biên dịch DOCX."}


@app.get("/api/compile/status")
def get_compile_status():
    """Polls active compilation progress and results."""
    return {
        "is_compiling": compile_manager.is_compiling,
        "progress": compile_manager.progress,
        "message": compile_manager.message,
        "error": compile_manager.error,
        "output_path": compile_manager.output_path
    }


@app.post("/api/compile/cancel")
def cancel_compilation():
    """Signals current compilation to stop."""
    if compile_manager.is_compiling:
        compile_manager.cancel_event.set()
        compile_manager.message = "Đang yêu cầu hủy..."
        return {"success": True, "message": "Đã gửi tín hiệu hủy tác vụ."}
    return {"success": False, "message": "Không có tác vụ nào đang chạy."}


@app.post("/api/upload-images")
async def upload_images(files: List[UploadFile] = File(...)):
    """Uploads multiple images into a temporary cache and generates thumbnails."""
    upload_dir = os.path.join(tempfile.gettempdir(), "docx_web_uploads")
    os.makedirs(upload_dir, exist_ok=True)

    results = []
    for file in files:
        if not file.filename:
            continue
        dest_path = os.path.join(upload_dir, f"{int(time.time() * 1000)}_{file.filename}")
        content = await file.read()
        with open(dest_path, "wb") as f_out:
            f_out.write(content)

        # Convert to Word-compatible format if WebP/HEIC/etc.
        real_path = ImageService.ensure_word_compatible_image(dest_path)
        sz_kb = os.path.getsize(real_path) / 1024.0

        # Generate base64 thumbnail
        thumb_b64 = ""
        try:
            pil_thumb = ImageService.create_thumbnail(real_path, max_size=(160, 100))
            if pil_thumb:
                buf = io.BytesIO()
                pil_thumb.save(buf, format="PNG")
                thumb_b64 = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")
        except Exception:
            pass

        results.append({
            "filename": file.filename,
            "filepath": real_path,
            "size_kb": round(sz_kb, 1),
            "thumbnail_url": thumb_b64
        })

    return {"success": True, "images": results}


@app.get("/api/docx-list")
def list_docx_files(folder: Optional[str] = Query(None)):
    """Scans directory for DOCX files containing embedded ULN."""
    target_dir = folder or os.path.expanduser("~\\Documents") if os.name == "nt" else os.getcwd()
    if not os.path.exists(target_dir):
        return {"items": [], "folder": target_dir}

    items = scan_folder_for_uln_docx(target_dir)
    return {"items": items, "folder": target_dir}


@app.post("/api/embed-uln")
def embed_uln(req: EmbedRequest):
    """Embeds raw ULN source text into an existing DOCX file."""
    if not os.path.exists(req.filepath):
        raise HTTPException(status_code=404, detail="Tệp DOCX không tồn tại.")
    ok = embed_raw_uln_docx(req.filepath, req.uln_text)
    if not ok:
        raise HTTPException(status_code=500, detail="Không thể nhúng mã ULN (vui lòng đóng file Word trước).")
    return {"success": True, "message": "Đã nhúng mã ULN thành công."}


@app.post("/api/open-docx")
def open_docx_file(req: OpenFileRequest):
    """Opens a compiled DOCX file with MS Word in Windows."""
    if not os.path.exists(req.filepath):
        raise HTTPException(status_code=404, detail="Tệp không tồn tại.")
    try:
        if os.name == "nt":
            os.startfile(req.filepath)
        else:
            webbrowser.open(f"file://{os.path.abspath(req.filepath)}")
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/check-update")
def check_update():
    """Checks for latest GitHub release updates."""
    return check_for_updates()


# ── Mount Frontend Static Web Assets ─────────────────────────────
web_dir = os.path.join(project_root, "web")
if os.path.exists(web_dir):
    app.mount("/static", StaticFiles(directory=web_dir), name="static")

    @app.get("/")
    def serve_index():
        return FileResponse(os.path.join(web_dir, "index.html"))


def find_free_port(start_port: int = 8765, max_port: int = 8800) -> int:
    """Finds an available TCP port for the local server."""
    for p in range(start_port, max_port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start_port


def start_server():
    """Entrypoint to start the server and open localhost in default web browser."""
    import uvicorn
    port = find_free_port(8765)
    url = f"http://127.0.0.1:{port}"

    print("=" * 60)
    print(f"🚀 ULN to DOCX Web Application Server running at: {url}")
    print("=" * 60)

    # Open user's browser after brief initialization delay
    def _open_browser():
        time.sleep(1.2)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=_open_browser, daemon=True).start()

    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")


if __name__ == "__main__":
    start_server()
