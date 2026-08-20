import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import updater

def get_prompt_storage_path() -> str:
    """Returns path where user-edited prompt is saved and loaded."""
    if getattr(sys, 'frozen', False):
        app_data_dir = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "DocxConverter")
        os.makedirs(app_data_dir, exist_ok=True)
        return os.path.join(app_data_dir, "prompt.txt")
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompt.txt")

def get_default_prompt_path() -> str:
    """Returns bundled default prompt path."""
    bundle_p = os.path.join(updater.get_bundle_dir(), "prompt.txt")
    if os.path.exists(bundle_p):
        return bundle_p
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompt.txt")

def load_prompt_text() -> str:
    """Loads prompt from storage path or fallback to default bundle."""
    storage_p = get_prompt_storage_path()
    if os.path.exists(storage_p):
        try:
            with open(storage_p, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            pass
    
    default_p = get_default_prompt_path()
    if os.path.exists(default_p):
        try:
            with open(default_p, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            pass
    return "You are an expert OCR and Universal Layout Notation (ULN) extraction engine."

def save_prompt_text(text: str) -> bool:
    """Saves prompt text to storage path (and project directory if dev mode)."""
    try:
        storage_p = get_prompt_storage_path()
        parent = os.path.dirname(storage_p)
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(storage_p, "w", encoding="utf-8") as f:
            f.write(text)
        
        # If in dev mode, also ensure local prompt.txt is synced
        if not getattr(sys, 'frozen', False):
            local_p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompt.txt")
            if local_p != storage_p:
                with open(local_p, "w", encoding="utf-8") as f:
                    f.write(text)
        return True
    except Exception as e:
        print(f"Error saving prompt: {e}")
        return False

def reset_prompt_text() -> str:
    """Resets prompt text from default bundled prompt and saves to storage path."""
    default_p = get_default_prompt_path()
    content = ""
    if os.path.exists(default_p):
        try:
            with open(default_p, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception:
            pass
    if content:
        save_prompt_text(content)
    return content

def open_prompt_editor_dialog(parent_root):
    """Opens a modern dark modal to view, edit, copy, and manage AI prompt."""
    win = tk.Toplevel(parent_root)
    win.title("📜 Quản Lý & Chỉnh Sửa AI Prompt (ULN Extraction)")
    win.geometry("960x720")
    win.configure(bg="#0f172a")
    win.transient(parent_root)
    win.grab_set()

    try:
        x = parent_root.winfo_x() + (parent_root.winfo_width() // 2) - 480
        y = parent_root.winfo_y() + (parent_root.winfo_height() // 2) - 360
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    # Top Header Bar
    top_bar = tk.Frame(win, bg="#1e293b", padx=20, pady=12)
    top_bar.pack(fill="x")

    tk.Label(
        top_bar,
        text="📜 AI System Prompt & Quy Tắc Trích Xuất ULN",
        font=("Segoe UI", 14, "bold"),
        bg="#1e293b",
        fg="#38bdf8"
    ).pack(anchor="w")

    tk.Label(
        top_bar,
        text="Prompt này được dùng để nạp vào ChatGPT, Claude, Gemini hoặc DeepSeek để nhận diện & chuyển đổi đề thi thành ULN.",
        font=("Segoe UI", 9),
        bg="#1e293b",
        fg="#94a3b8"
    ).pack(anchor="w", pady=(2, 0))

    # Toolbar
    tool_frame = tk.Frame(win, bg="#111827", padx=16, pady=8)
    tool_frame.pack(fill="x")

    def copy_all():
        text = p_text.get("1.0", tk.END).strip()
        parent_root.clipboard_clear()
        parent_root.clipboard_append(text)
        status_lbl.config(text="✓ Đã sao chép toàn bộ Prompt vào Clipboard!", fg="#4ade80")
        parent_root.after(3000, lambda: status_lbl.config(text=f"File: {get_prompt_storage_path()}", fg="#94a3b8"))

    def save_changes():
        text = p_text.get("1.0", tk.END).rstrip() + "\n"
        if save_prompt_text(text):
            status_lbl.config(text="✓ Đã lưu thay đổi Prompt thành công!", fg="#4ade80")
            messagebox.showinfo("Thành công", "Đã lưu nội dung Prompt thành công!")
        else:
            status_lbl.config(text="⚠ Lỗi khi lưu file prompt!", fg="#f43f5e")
            messagebox.showerror("Lỗi", "Không thể lưu file prompt.")
        parent_root.after(3000, lambda: status_lbl.config(text=f"File: {get_prompt_storage_path()}", fg="#94a3b8"))

    def reset_default():
        if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn khôi phục lại Prompt mặc định ban đầu không?"):
            default_content = reset_prompt_text()
            p_text.delete("1.0", tk.END)
            p_text.insert("1.0", default_content)
            status_lbl.config(text="✓ Đã khôi phục Prompt về mặc định!", fg="#38bdf8")
            parent_root.after(3000, lambda: status_lbl.config(text=f"File: {get_prompt_storage_path()}", fg="#94a3b8"))

    def open_folder():
        storage_path = get_prompt_storage_path()
        folder = os.path.dirname(storage_path)
        if not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)
        if not os.path.exists(storage_path):
            save_prompt_text(p_text.get("1.0", tk.END))
        try:
            subprocess.Popen(f'explorer /select,"{storage_path}"')
        except Exception:
            try:
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở thư mục: {e}")

    btn_copy = tk.Button(
        tool_frame,
        text="📋 Sao Chép Prompt (Copy)",
        command=copy_all,
        bg="#2563eb",
        fg="#ffffff",
        activebackground="#1d4ed8",
        activeforeground="#ffffff",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=12,
        pady=5,
        cursor="hand2"
    )
    btn_copy.pack(side="left", padx=(0, 6))

    btn_save = tk.Button(
        tool_frame,
        text="💾 Lưu Thay Đổi (Save)",
        command=save_changes,
        bg="#16a34a",
        fg="#ffffff",
        activebackground="#15803d",
        activeforeground="#ffffff",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=12,
        pady=5,
        cursor="hand2"
    )
    btn_save.pack(side="left", padx=6)

    btn_reset = tk.Button(
        tool_frame,
        text="🔄 Khôi Phục Mặc Định",
        command=reset_default,
        bg="#475569",
        fg="#ffffff",
        activebackground="#334155",
        activeforeground="#ffffff",
        font=("Segoe UI", 9),
        relief="flat",
        padx=10,
        pady=5,
        cursor="hand2"
    )
    btn_reset.pack(side="left", padx=6)

    btn_open_file = tk.Button(
        tool_frame,
        text="📁 Mở Thư Mục Chứa File",
        command=open_folder,
        bg="#334155",
        fg="#e2e8f0",
        activebackground="#1e293b",
        activeforeground="#ffffff",
        font=("Segoe UI", 9),
        relief="flat",
        padx=10,
        pady=5,
        cursor="hand2"
    )
    btn_open_file.pack(side="left", padx=6)

    # Editor Area
    editor_box = tk.Frame(win, bg="#0f172a", padx=16, pady=10)
    editor_box.pack(fill="both", expand=True)

    scroll_y = ttk.Scrollbar(editor_box, orient="vertical")
    scroll_y.pack(side="right", fill="y")

    p_text = tk.Text(
        editor_box,
        wrap="word",
        bg="#090d16",
        fg="#f8fafc",
        insertbackground="#ffffff",
        font=("Consolas", 10),
        yscrollcommand=scroll_y.set,
        padx=12,
        pady=12,
        relief="flat"
    )
    p_text.pack(fill="both", expand=True)
    scroll_y.config(command=p_text.yview)

    # Load current content
    current_content = load_prompt_text()
    p_text.insert("1.0", current_content)

    # Status Footer
    footer = tk.Frame(win, bg="#1e293b", padx=16, pady=8)
    footer.pack(fill="x", side="bottom")

    status_lbl = tk.Label(
        footer,
        text=f"File: {get_prompt_storage_path()}",
        font=("Segoe UI", 9),
        bg="#1e293b",
        fg="#94a3b8"
    )
    status_lbl.pack(side="left")

    btn_close = tk.Button(
        footer,
        text="Đóng",
        command=win.destroy,
        bg="#475569",
        fg="#ffffff",
        font=("Segoe UI", 9),
        relief="flat",
        padx=14,
        pady=4,
        cursor="hand2"
    )
    btn_close.pack(side="right")
