"""
server/app.py - High-Performance Localhost Web Server & REST API.
Serves the modern Flat Minimal Light Web UI and orchestrates background Word COM automation.
Includes automated idle watchdog: closes terminal process when web browser tab is closed.
"""

import os
import sys
import time
import socket
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

# ── Heartbeat & Idle Auto-Shutdown Watchdog ───────────────────────
last_heartbeat_time = time.time()
has_received_first_heartbeat = False


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


class ReadFileRequest(BaseModel):
    filepath: str


class SaveFileRequest(BaseModel):
    filename: str
    content: str
    folder: Optional[str] = None


class DeleteFileRequest(BaseModel):
    filepath: str


class OpenFolderRequest(BaseModel):
    filepath: Optional[str] = None
    folder: Optional[str] = None


# ── Heartbeat & Shutdown Endpoints ───────────────────────────────
@app.post("/api/heartbeat")
def record_heartbeat():
    """Receives ping from browser tab. Keeps local server alive."""
    global last_heartbeat_time, has_received_first_heartbeat
    last_heartbeat_time = time.time()
    has_received_first_heartbeat = True
    return {"status": "ok"}


@app.post("/api/shutdown")
def shutdown_server():
    """Immediately stops server and exits terminal when browser tab is closed."""
    def _delayed_exit():
        time.sleep(0.3)
        print("\n[Server] Tab trình duyệt đã đóng. Đang tắt tiến trình máy chủ...")
        os._exit(0)
    threading.Thread(target=_delayed_exit, daemon=True).start()
    return {"status": "shutting down"}


# ── REST API Endpoints ───────────────────────────────────────────
@app.get("/api/status")
def get_status():
    """Returns application version and environment status."""
    output_dir = os.path.join(project_root, "output")
    os.makedirs(output_dir, exist_ok=True)
    return {
        "status": "online",
        "version": get_current_version(),
        "output_folder": output_dir,
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

    output_dir = os.path.join(project_root, "output")
    os.makedirs(output_dir, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    fname = req.output_filename or f"uln_document_{timestamp}.docx"
    if not fname.lower().endswith(".docx"):
        fname += ".docx"
    output_path = os.path.abspath(os.path.join(output_dir, fname))

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

    compile_manager.worker_thread = threading.Thread(target=_worker, daemon=True)
    compile_manager.worker_thread.start()

    return {"status": "started", "target": output_path}


@app.get("/api/compile/status")
def get_compile_status():
    """Polls real-time compilation progress."""
    return {
        "is_compiling": compile_manager.is_compiling,
        "progress": compile_manager.progress,
        "message": compile_manager.message,
        "error": compile_manager.error,
        "output_path": compile_manager.output_path
    }


@app.post("/api/compile/cancel")
def cancel_compilation():
    """Signals cancellation to the compilation thread."""
    if compile_manager.is_compiling:
        compile_manager.cancel_event.set()
        compile_manager.message = "Đang hủy tác vụ..."
        return {"status": "cancelling"}
    return {"status": "idle"}


@app.post("/api/upload-images")
async def upload_images(files: List[UploadFile] = File(...)):
    """Receives uploaded images, converts WebP/HEIC/AVIF if needed, returns thumbnail base64."""
    uploaded_items = []
    temp_dir = os.path.join(project_root, "output", "temp_images")
    os.makedirs(temp_dir, exist_ok=True)

    for file in files:
        filename = file.filename
        content = await file.read()
        target_path = os.path.join(temp_dir, filename)

        try:
            with open(target_path, "wb") as f:
                f.write(content)

            # Process image format conversion via ImageService
            actual_path = ImageService.process_image_for_word(target_path)
            thumb_b64 = ImageService.create_thumbnail_base64(actual_path, max_size=(100, 100))

            uploaded_items.append({
                "filename": filename,
                "temp_path": actual_path,
                "size": len(content),
                "thumbnail": thumb_b64
            })
        except Exception as e:
            uploaded_items.append({
                "filename": filename,
                "error": str(e)
            })

    return {"uploaded": uploaded_items}


# ── File Management REST APIs ─────────────────────────────────────
@app.get("/api/files")
def list_managed_files(folder: Optional[str] = None, category: Optional[str] = "all"):
    """Lists DOCX and ULN/TXT files with metadata in the target folder."""
    target_dir = folder
    output_dir = os.path.join(project_root, "output")
    os.makedirs(output_dir, exist_ok=True)

    if not target_dir or not os.path.exists(target_dir):
        target_dir = output_dir

    files_list = []
    try:
        for entry in os.scandir(target_dir):
            if entry.is_file():
                if entry.name.startswith("~$"):
                    continue
                ext = os.path.splitext(entry.name)[1].lower()
                is_docx = (ext == ".docx")
                is_text = (ext in [".uln", ".txt", ".md"])

                if category == "docx" and not is_docx:
                    continue
                if category == "uln" and not is_text:
                    continue
                if not (is_docx or is_text):
                    continue

                stat = entry.stat()
                has_embedded = False
                embedded_preview = ""
                if is_docx:
                    embedded_text = extract_raw_uln(entry.path)
                    if embedded_text:
                        has_embedded = True
                        embedded_preview = embedded_text[:400]

                files_list.append({
                    "name": entry.name,
                    "path": entry.path,
                    "ext": ext,
                    "size": stat.st_size,
                    "mtime": stat.st_mtime,
                    "mtime_formatted": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(stat.st_mtime)),
                    "is_docx": is_docx,
                    "is_text": is_text,
                    "has_embedded_uln": has_embedded,
                    "embedded_preview": embedded_preview
                })
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    # Sort newest modified first
    files_list.sort(key=lambda x: x["mtime"], reverse=True)

    docs_dir = os.path.expanduser("~\\Documents") if os.name == "nt" else os.getcwd()
    return {
        "folder": target_dir,
        "files": files_list,
        "output_folder": output_dir,
        "documents_folder": docs_dir
    }


@app.post("/api/files/read")
def read_file_content(req: ReadFileRequest):
    """Reads raw text or extracted embedded ULN from a file."""
    if not os.path.exists(req.filepath):
        raise HTTPException(status_code=404, detail="Tệp không tồn tại.")
    ext = os.path.splitext(req.filepath)[1].lower()
    if ext == ".docx":
        content = extract_raw_uln(req.filepath)
        if not content:
            raise HTTPException(status_code=400, detail="File DOCX này không có nhúng mã nguồn ULN.")
        return {"content": content, "filepath": req.filepath, "type": "docx_embedded"}
    else:
        try:
            with open(req.filepath, "r", encoding="utf-8") as f:
                content = f.read()
            return {"content": content, "filepath": req.filepath, "type": "text"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Lỗi đọc file: {str(e)}")


@app.post("/api/files/save")
def save_uln_file(req: SaveFileRequest):
    """Saves ULN content to a file in output/ directory."""
    folder = req.folder
    output_dir = os.path.join(project_root, "output")
    os.makedirs(output_dir, exist_ok=True)

    if not folder or not os.path.exists(folder):
        folder = output_dir

    filename = req.filename.strip()
    if not filename.endswith(".uln") and not filename.endswith(".txt"):
        filename += ".uln"

    target_path = os.path.join(folder, filename)
    try:
        with open(target_path, "w", encoding="utf-8") as f:
            f.write(req.content)
        return {"success": True, "filepath": target_path, "filename": filename}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/files/delete")
def delete_file(req: DeleteFileRequest):
    """Deletes a file."""
    if not os.path.exists(req.filepath):
        raise HTTPException(status_code=404, detail="Tệp không tồn tại.")
    try:
        os.remove(req.filepath)
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/files/open-folder")
def open_folder_in_explorer(req: OpenFolderRequest):
    """Opens target folder in Windows Explorer."""
    path = req.folder or (os.path.dirname(req.filepath) if req.filepath else None)
    if not path or not os.path.exists(path):
        path = os.path.join(project_root, "output")
    try:
        if os.name == "nt":
            if req.filepath and os.path.exists(req.filepath):
                os.system(f'explorer /select,"{os.path.abspath(req.filepath)}"')
            else:
                os.startfile(path)
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/docx-list")
def list_docx_files(folder: Optional[str] = None, only_uln: bool = True):
    """Scans directory for DOCX files and checks for embedded ULN."""
    target_dir = folder or os.path.join(project_root, "output")
    if not os.path.exists(target_dir):
        target_dir = os.path.expanduser("~\\Documents") if os.name == "nt" else os.getcwd()
    try:
        return scan_folder_for_uln_docx(target_dir, only_embedded=only_uln)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/embed-uln")
def embed_uln_to_docx(req: EmbedRequest):
    """Embeds raw ULN source into an existing DOCX file."""
    if not os.path.exists(req.filepath):
        raise HTTPException(status_code=404, detail="File DOCX không tồn tại.")
    success = embed_raw_uln_docx(req.filepath, req.uln_text)
    if not success:
        raise HTTPException(status_code=500, detail="Nhúng mã ULN thất bại.")
    return {"success": True}


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
    print(f"ULN to DOCX Studio Localhost Web Server running at: {url}")
    print("=" * 60)

    # Browser opener thread
    def _open_browser():
        time.sleep(1.0)
        try:
            webbrowser.open(url)
        except Exception:
            pass

    threading.Thread(target=_open_browser, daemon=True).start()

    # Watchdog thread: kills terminal process when web tab is closed
    def _watchdog():
        time.sleep(25) # 25s initial startup grace period
        while True:
            time.sleep(2.0)
            if has_received_first_heartbeat and (time.time() - last_heartbeat_time > 8):
                print("\n[Server] Khong con ket noi tu trinh duyet (tab da dong). Tu dong tat may chu.")
                os._exit(0)

    threading.Thread(target=_watchdog, daemon=True).start()

    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")


if __name__ == "__main__":
    start_server()
