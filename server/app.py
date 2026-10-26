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
import json
import webbrowser
from typing import List, Dict, Any, Optional

# Add project root to sys.path so core, renderers, and services import seamlessly
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query, Request
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
        self.start_time: float = 0.0

    def reset(self):
        self.is_compiling = True
        self.progress = 0.05
        self.message = "Đang khởi tạo tài liệu..."
        self.error = None
        self.output_path = None
        self.cancel_event.clear()
        self.start_time = time.time()

compile_manager = CompileManager()


# ── Request / Response Models ────────────────────────────────────
class CompileRequest(BaseModel):
    uln_text: str
    settings: Dict[str, Any] = {}
    output_filename: Optional[str] = None
    open_word: bool = True
    background_mode: bool = False
    user_images: Optional[List[str]] = None


class SettingsRequest(BaseModel):
    settings: Dict[str, Any]


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


def sanitize_filepath(p: Optional[str]) -> str:
    """Strips leading/trailing quotes and extra whitespace from file paths."""
    if not p:
        return ""
    clean = p.strip().strip('"').strip("'").strip()
    return os.path.abspath(clean)


# ── Heartbeat & Shutdown Endpoints ───────────────────────────────
shutdown_in_progress = False


@app.post("/api/heartbeat")
def record_heartbeat():
    """Receives ping from browser tab to confirm connection."""
    global last_heartbeat_time, has_received_first_heartbeat
    last_heartbeat_time = time.time()
    has_received_first_heartbeat = True
    return {"status": "ok"}


@app.post("/api/shutdown")
def shutdown_server():
    """Terminates server when user explicitly clicks 'Thoát ứng dụng'."""
    global shutdown_in_progress
    if shutdown_in_progress:
        return {"status": "in_progress"}
    shutdown_in_progress = True

    def _exit():
        time.sleep(0.5)
        print("\n[Server] Người dùng đã dừng máy chủ. Đang thoát.")
        os._exit(0)

    threading.Thread(target=_exit, daemon=True).start()
    return {"status": "scheduled"}


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


SETTINGS_FILE = os.path.join(project_root, "settings.json")


@app.get("/api/settings")
def get_user_settings():
    """Reads saved user configuration from settings.json."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {"settings": data}
        except Exception:
            pass
    return {"settings": {}}


@app.post("/api/settings")
def save_user_settings(req: SettingsRequest):
    """Saves user configuration to settings.json."""
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(req.settings, f, ensure_ascii=False, indent=2)
        return {"success": True, "saved": req.settings}
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
            font_name = settings.get("font_name") or settings.get("font_family") or "Times New Roman"
            settings["font_name"] = font_name
            settings["font_family"] = font_name
            if req.user_images:
                settings["user_images"] = list(req.user_images)
            elif "user_images" not in settings:
                settings["user_images"] = []
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
    """Polls real-time compilation progress with estimated remaining time."""
    now = time.time()
    elapsed = max(0.0, now - compile_manager.start_time) if compile_manager.is_compiling else 0.0

    remaining_seconds = 0
    if compile_manager.is_compiling and compile_manager.progress > 0.05:
        estimated_total = elapsed / max(0.01, compile_manager.progress)
        remaining_seconds = max(1, int(round(estimated_total - elapsed)))

    return {
        "is_compiling": compile_manager.is_compiling,
        "progress": round(compile_manager.progress, 3),
        "message": compile_manager.message,
        "elapsed_seconds": round(elapsed, 1),
        "remaining_seconds": remaining_seconds,
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
    output_dir = os.path.join(project_root, "output")
    os.makedirs(output_dir, exist_ok=True)

    target_dir = sanitize_filepath(folder) if folder else output_dir
    if not target_dir:
        target_dir = output_dir

    # If user provided a path to a specific file, browse its parent directory
    if os.path.isfile(target_dir):
        target_dir = os.path.dirname(target_dir)

    if not os.path.exists(target_dir):
        raise HTTPException(status_code=404, detail=f"Thư mục không tồn tại: {target_dir}")

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
    filepath = sanitize_filepath(req.filepath)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail=f"Tệp không tồn tại: {filepath}")
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".docx":
        content = extract_raw_uln(filepath)
        if not content:
            raise HTTPException(status_code=400, detail="File DOCX này không có nhúng mã nguồn ULN.")
        return {"content": content, "filepath": filepath, "type": "docx_embedded"}
    else:
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()
            return {"content": content, "filepath": filepath, "type": "text"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Lỗi đọc file: {str(e)}")


@app.post("/api/files/save")
def save_uln_file(req: SaveFileRequest):
    """Saves ULN content to a file in output/ directory."""
    folder = sanitize_filepath(req.folder) if req.folder else os.path.join(project_root, "output")
    os.makedirs(folder, exist_ok=True)

    filename = req.filename.strip().strip('"').strip("'")
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
    filepath = sanitize_filepath(req.filepath)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="Tệp không tồn tại.")
    try:
        os.remove(filepath)
        return {"success": True}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.api_route("/api/files/open-folder", methods=["GET", "POST", "OPTIONS"])
@app.api_route("/api/files/open-folder/", methods=["GET", "POST", "OPTIONS"])
async def open_folder_in_explorer(
    request: Request,
    folder: Optional[str] = Query(None),
    filepath: Optional[str] = Query(None)
):
    """Opens target folder or selects target file in Windows Explorer."""
    if request.method == "OPTIONS":
        return {"status": "ok"}

    target_folder = folder
    target_filepath = filepath

    if request.method == "POST":
        try:
            body = await request.json()
            if isinstance(body, dict):
                target_folder = body.get("folder", target_folder)
                target_filepath = body.get("filepath", target_filepath)
        except Exception:
            pass

    clean_file = sanitize_filepath(target_filepath) if target_filepath else ""
    clean_folder = sanitize_filepath(target_folder) if target_folder else ""

    output_dir = os.path.join(project_root, "output")
    os.makedirs(output_dir, exist_ok=True)

    path = clean_file or clean_folder or output_dir
    if not os.path.exists(path):
        parent = os.path.dirname(path)
        if parent and os.path.exists(parent):
            path = parent
        else:
            path = output_dir

    try:
        abs_path = os.path.abspath(path)
        if os.name == "nt":
            import subprocess
            if clean_file and os.path.isfile(clean_file):
                try:
                    subprocess.Popen(['explorer.exe', f'/select,{os.path.abspath(clean_file)}'])
                except Exception:
                    target_dir = abs_path if os.path.isdir(abs_path) else os.path.dirname(abs_path)
                    os.startfile(target_dir)
            else:
                target_dir = abs_path if os.path.isdir(abs_path) else os.path.dirname(abs_path)
                try:
                    os.startfile(target_dir)
                except Exception:
                    subprocess.Popen(['explorer.exe', target_dir])
        else:
            webbrowser.open(f"file://{abs_path}")
        return {"success": True, "opened": abs_path}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/docx-list")
def list_docx_files(folder: Optional[str] = None, only_uln: bool = True):
    """Scans directory for DOCX files and checks for embedded ULN."""
    target_dir = sanitize_filepath(folder) if folder else os.path.join(project_root, "output")
    if not os.path.exists(target_dir):
        target_dir = os.path.expanduser("~\\Documents") if os.name == "nt" else os.getcwd()
    try:
        return scan_folder_for_uln_docx(target_dir, only_embedded=only_uln)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.api_route("/api/embed-uln", methods=["GET", "POST", "OPTIONS"])
@app.api_route("/api/embed-uln/", methods=["GET", "POST", "OPTIONS"])
async def embed_uln_to_docx(
    request: Request,
    filepath: Optional[str] = Query(None),
    uln_text: Optional[str] = Query(None)
):
    """Embeds raw ULN source into an existing DOCX file."""
    if request.method == "OPTIONS":
        return {"status": "ok"}

    target_filepath = filepath
    target_uln = uln_text

    if request.method == "POST":
        try:
            body = await request.json()
            if isinstance(body, dict):
                target_filepath = body.get("filepath", target_filepath)
                target_uln = body.get("uln_text", target_uln)
        except Exception:
            pass

    if not target_filepath:
        if request.method == "GET":
            return {
                "status": "ready",
                "info": "Gửi yêu cầu POST với JSON {'filepath': '...', 'uln_text': '...'} để nhúng mã ULN vào file Word."
            }
        raise HTTPException(status_code=400, detail="Vui lòng cung cấp đường dẫn tệp DOCX (filepath).")

    clean_path = sanitize_filepath(target_filepath)
    if not os.path.exists(clean_path):
        raise HTTPException(status_code=404, detail=f"Không tìm thấy file DOCX: {clean_path}")
    if not clean_path.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Tệp được chọn không phải là tệp định dạng .docx.")

    if target_uln is None:
        target_uln = ""

    success = embed_raw_uln_docx(clean_path, target_uln)
    if not success:
        raise HTTPException(status_code=500, detail="Nhúng mã ULN thất bại. Vui lòng đảm bảo file không bị khóa hoặc đang mở ở tiến trình khác.")
    return {"success": True, "filepath": clean_path}


@app.post("/api/open-docx")
def open_docx_file(req: OpenFileRequest):
    """Opens a compiled DOCX file with MS Word in Windows."""
    filepath = sanitize_filepath(req.filepath)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail=f"Tệp không tồn tại: {filepath}")
    try:
        if os.name == "nt":
            os.startfile(filepath)
        else:
            webbrowser.open(f"file://{os.path.abspath(filepath)}")
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


def find_free_port(start_port: int = 8899, max_port: int = 8999, host: str = "127.0.0.1") -> int:
    """Finds an available TCP port for the local server by testing socket bind capability."""
    for p in range(start_port, max_port):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                s.bind((host, p))
                return p
        except OSError:
            continue
    return start_port


def start_server(port: Optional[int] = None, host: str = "127.0.0.1", open_browser: bool = True):
    """Entrypoint to start the server and open localhost in default web browser."""
    import uvicorn

    # Priority: explicit arg > env DOCX_STUDIO_PORT / PORT > find_free_port(8899)
    if port is None:
        env_port = os.environ.get("DOCX_STUDIO_PORT") or os.environ.get("PORT")
        if env_port:
            try:
                port = int(env_port)
            except ValueError:
                port = None
    if port is None:
        port = find_free_port(8899, host=host)

    url = f"http://{host}:{port}"

    print("=" * 60)
    print(f"ULN to DOCX Studio Localhost Web Server running at: {url}")
    print("=" * 60)

    # Browser opener thread
    if open_browser:
        def _open_browser():
            time.sleep(1.0)
            try:
                webbrowser.open(url)
            except Exception:
                pass

        threading.Thread(target=_open_browser, daemon=True).start()

    uvicorn.run(app, host=host, port=port, log_level="warning")


if __name__ == "__main__":
    start_server()
