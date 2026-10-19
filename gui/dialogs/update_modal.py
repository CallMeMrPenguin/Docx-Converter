"""
gui/dialogs/update_modal.py - Flat Minimal Update Modal Dialog.
Clean release notification and background download progress with zero borders.
"""

import threading
import tkinter as tk
from tkinter import ttk
import updater
from gui.theme import Theme


def show_update_modal_dialog(parent_root, info: dict):
    """Modern flat light-themed update modal popup dialog."""
    win = tk.Toplevel(parent_root)
    win.title("Có bản cập nhật mới")
    win.geometry("560x420")
    win.configure(bg=Theme.BG_CANVAS)
    win.transient(parent_root)
    win.grab_set()
    win.resizable(False, False)

    try:
        x = parent_root.winfo_x() + (parent_root.winfo_width() // 2) - 280
        y = parent_root.winfo_y() + (parent_root.winfo_height() // 2) - 210
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    # Top Header
    top_bar = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=14)
    top_bar.pack(fill="x")

    tk.Label(
        top_bar,
        text="Đã có phiên bản mới!",
        font=Theme.FONT_TITLE,
        bg=Theme.BG_SURFACE,
        fg=Theme.PRIMARY
    ).pack(anchor="w")

    ver_text = f"Hiện tại: v{info.get('current_version', '')}  ➔  Mới nhất: v{info.get('latest_version', '')}"
    tk.Label(
        top_bar,
        text=ver_text,
        font=Theme.FONT_BODY_BOLD,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_SECONDARY
    ).pack(anchor="w", pady=(3, 0))

    # Content Area
    content = tk.Frame(win, bg=Theme.BG_CANVAS, padx=20, pady=12)
    content.pack(fill="both", expand=True)

    tk.Label(
        content,
        text="Thông tin cập nhật:",
        font=Theme.FONT_BODY_BOLD,
        bg=Theme.BG_CANVAS,
        fg=Theme.TEXT_PRIMARY
    ).pack(anchor="w", pady=(0, 6))

    notes_card = tk.Frame(content, bg=Theme.BG_SURFACE, padx=2, pady=2)
    notes_card.pack(fill="both", expand=True)

    notes_box = tk.Text(
        notes_card,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_PRIMARY,
        font=Theme.FONT_BODY,
        wrap="word",
        height=8,
        padx=10,
        pady=8,
        bd=0,
        highlightthickness=0,
        relief="flat"
    )
    notes_box.pack(fill="both", expand=True)
    notes_box.insert("1.0", f"{info.get('release_title', '')}\n\n{info.get('release_notes', '')}")
    notes_box.config(state="disabled")

    status_lbl = tk.Label(
        content,
        text="",
        font=Theme.FONT_SMALL,
        bg=Theme.BG_CANVAS,
        fg=Theme.PRIMARY
    )
    status_lbl.pack(fill="x", pady=(6, 0))

    progress_bar = ttk.Progressbar(content, mode="determinate", style="Horizontal.TProgressbar")

    # Bottom Buttons
    btn_frame = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    btn_frame.pack(fill="x", side="bottom")

    def start_update():
        btn_update_now.config(state="disabled")
        btn_cancel.config(state="disabled")
        progress_bar.pack(fill="x", pady=(4, 0))
        status_lbl.config(text="Đang kết nối tải bản cập nhật...")

        def progress_cb(downloaded, total):
            if total > 0:
                pct = int(downloaded / total * 100)
                def _update_ui():
                    try:
                        if win.winfo_exists():
                            progress_bar.config(value=pct)
                            status_lbl.config(text=f"Đang tải bản cập nhật: {pct}% ({downloaded // 1024} KB / {total // 1024} KB)")
                    except Exception:
                        pass
                parent_root.after(0, _update_ui)

        def run_dl():
            ok, msg = updater.download_and_install_update(
                download_url=info.get("download_url"),
                release_url=info.get("release_url", ""),
                progress_callback=progress_cb
            )
            def _handle_dl_result():
                try:
                    if not win.winfo_exists():
                        return
                    if not ok:
                        progress_bar.pack_forget()
                        status_lbl.config(text=msg, fg=Theme.DANGER)
                        btn_update_now.config(state="normal", text="Mở trang tải trên Web")
                        btn_update_now.config(command=lambda: updater.webbrowser.open(info.get("release_url", "")))
                        btn_cancel.config(state="normal")
                    else:
                        status_lbl.config(text=msg, fg=Theme.SUCCESS)
                        btn_cancel.config(state="normal", text="Đóng")
                except Exception:
                    pass
            parent_root.after(0, _handle_dl_result)

        threading.Thread(target=run_dl, daemon=True).start()

    btn_cancel = tk.Button(
        btn_frame,
        text="Để sau",
        command=win.destroy,
        bg=Theme.BG_SURFACE_ALT,
        fg=Theme.TEXT_SECONDARY,
        activebackground=Theme.BG_HOVER,
        font=Theme.FONT_BODY,
        relief="flat",
        bd=0,
        padx=14,
        pady=5,
        cursor="hand2"
    )
    btn_cancel.pack(side="right", padx=(8, 0))

    btn_update_now = tk.Button(
        btn_frame,
        text="Cập nhật ngay",
        command=start_update,
        bg=Theme.SUCCESS,
        fg=Theme.TEXT_ON_PRIMARY,
        activebackground=Theme.SUCCESS_HOVER,
        activeforeground=Theme.TEXT_ON_PRIMARY,
        font=Theme.FONT_BODY_BOLD,
        relief="flat",
        bd=0,
        padx=16,
        pady=5,
        cursor="hand2"
    )
    btn_update_now.pack(side="right")
