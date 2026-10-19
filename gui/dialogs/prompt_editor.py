"""
gui/dialogs/prompt_editor.py - Flat Minimal AI Prompt & Rules Editor Dialog.
Allows viewing, editing, copying, and saving ULN extraction rules with zero borders.
"""

import os
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import updater
from gui.theme import Theme


def get_prompt_storage_path() -> str:
    """Returns path where user-edited prompt is saved and loaded."""
    if getattr(sys, 'frozen', False):
        app_data_dir = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "DocxConverter")
        os.makedirs(app_data_dir, exist_ok=True)
        return os.path.join(app_data_dir, "prompt.txt")
    else:
        return os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "prompt.txt")


def get_default_prompt_path() -> str:
    """Returns bundled default prompt path."""
    bundle_p = os.path.join(updater.get_bundle_dir(), "prompt.txt")
    if os.path.exists(bundle_p):
        return bundle_p
    return os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "prompt.txt")


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
    """Saves prompt text to storage path."""
    try:
        storage_p = get_prompt_storage_path()
        parent = os.path.dirname(storage_p)
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(storage_p, "w", encoding="utf-8") as f:
            f.write(text)

        if not getattr(sys, 'frozen', False):
            local_p = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "prompt.txt")
            if local_p != storage_p and os.path.exists(os.path.dirname(local_p)):
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
    """Opens a clean Light UI modal to view, edit, copy, and manage AI prompt."""
    win = tk.Toplevel(parent_root)
    win.title("Quy tắc & Prompt trích xuất ULN")
    win.geometry("960x720")
    win.configure(bg=Theme.BG_CANVAS)
    win.transient(parent_root)
    win.grab_set()

    try:
        x = parent_root.winfo_x() + (parent_root.winfo_width() // 2) - 480
        y = parent_root.winfo_y() + (parent_root.winfo_height() // 2) - 360
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    # Top Header Bar
    top_bar = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    top_bar.pack(fill="x")

    tk.Label(
        top_bar,
        text="Quy tắc trích xuất ULN cho AI",
        font=Theme.FONT_TITLE,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_PRIMARY
    ).pack(anchor="w")

    tk.Label(
        top_bar,
        text="Copy prompt này gửi cho ChatGPT, Claude hoặc Gemini để chuyển đổi đề thi thành định dạng ULN chuẩn xác.",
        font=Theme.FONT_SUBTITLE,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_SECONDARY
    ).pack(anchor="w", pady=(2, 0))

    # Editor Container
    center_frame = tk.Frame(win, bg=Theme.BG_CANVAS, padx=16, pady=10)
    center_frame.pack(fill="both", expand=True)

    editor_card = tk.Frame(center_frame, bg=Theme.BG_SURFACE, padx=4, pady=4)
    editor_card.pack(fill="both", expand=True)

    txt_scroll = ttk.Scrollbar(editor_card)
    txt_scroll.pack(side="right", fill="y")

    prompt_editor = tk.Text(
        editor_card,
        wrap="word",
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_PRIMARY,
        insertbackground=Theme.PRIMARY,
        font=Theme.FONT_CODE,
        yscrollcommand=txt_scroll.set,
        padx=12,
        pady=10,
        bd=0,
        highlightthickness=0,
        relief="flat"
    )
    prompt_editor.pack(fill="both", expand=True)
    txt_scroll.config(command=prompt_editor.yview)

    current_text = load_prompt_text()
    prompt_editor.insert("1.0", current_text)

    # Bottom Actions Bar
    bot_bar = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    bot_bar.pack(fill="x", side="bottom")

    status_lbl = tk.Label(bot_bar, text="", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE, fg=Theme.SUCCESS)
    status_lbl.pack(side="left")

    def show_msg(msg: str, is_error: bool = False):
        status_lbl.config(text=msg, fg=Theme.DANGER if is_error else Theme.SUCCESS)
        win.after(3000, lambda: status_lbl.config(text=""))

    def copy_all():
        text = prompt_editor.get("1.0", tk.END).strip()
        win.clipboard_clear()
        win.clipboard_append(text)
        show_msg("✓ Đã sao chép toàn bộ prompt vào Clipboard!")

    def save_changes():
        text = prompt_editor.get("1.0", tk.END).strip()
        if save_prompt_text(text):
            show_msg("✓ Đã lưu thay đổi thành công!")
        else:
            show_msg("✕ Lỗi lưu prompt", is_error=True)

    def reset_default():
        if messagebox.askyesno("Khôi phục mặc định", "Bạn có chắc chắn muốn khôi phục về prompt mẫu ban đầu không?"):
            content = reset_prompt_text()
            prompt_editor.delete("1.0", tk.END)
            prompt_editor.insert("1.0", content)
            show_msg("✓ Đã khôi phục prompt gốc!")

    btn_close = tk.Button(
        bot_bar,
        text="Đóng",
        command=win.destroy,
        bg=Theme.BG_SURFACE_ALT,
        fg=Theme.TEXT_PRIMARY,
        activebackground=Theme.BG_HOVER,
        font=Theme.FONT_BODY,
        relief="flat",
        bd=0,
        padx=14,
        pady=5,
        cursor="hand2"
    )
    btn_close.pack(side="right", padx=(8, 0))

    btn_save = tk.Button(
        bot_bar,
        text="💾 Lưu Thay Đổi",
        command=save_changes,
        bg=Theme.PRIMARY,
        fg=Theme.TEXT_ON_PRIMARY,
        activebackground=Theme.PRIMARY_HOVER,
        activeforeground=Theme.TEXT_ON_PRIMARY,
        font=Theme.FONT_BODY_BOLD,
        relief="flat",
        bd=0,
        padx=14,
        pady=5,
        cursor="hand2"
    )
    btn_save.pack(side="right", padx=(8, 0))

    btn_copy = tk.Button(
        bot_bar,
        text="📋 Copy Toàn Bộ Prompt",
        command=copy_all,
        bg=Theme.SUCCESS,
        fg=Theme.TEXT_ON_PRIMARY,
        activebackground=Theme.SUCCESS_HOVER,
        activeforeground=Theme.TEXT_ON_PRIMARY,
        font=Theme.FONT_BODY_BOLD,
        relief="flat",
        bd=0,
        padx=14,
        pady=5,
        cursor="hand2"
    )
    btn_copy.pack(side="right", padx=(8, 0))

    btn_reset = tk.Button(
        bot_bar,
        text="↺ Khôi Phục Gốc",
        command=reset_default,
        bg=Theme.BG_SURFACE_ALT,
        fg=Theme.TEXT_SECONDARY,
        activebackground=Theme.BG_HOVER,
        font=Theme.FONT_BODY,
        relief="flat",
        bd=0,
        padx=12,
        pady=5,
        cursor="hand2"
    )
    btn_reset.pack(side="left")
