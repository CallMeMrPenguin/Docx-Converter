import threading
import tkinter as tk
from tkinter import ttk, messagebox
import updater

def show_update_modal_dialog(parent_root, info: dict):
    """Modern dark-themed update popup dialog."""
    win = tk.Toplevel(parent_root)
    win.title("Có Bản Cập Nhật Mới")
    win.geometry("560x420")
    win.configure(bg="#0f172a")
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
    top_bar = tk.Frame(win, bg="#1e293b", padx=16, pady=12)
    top_bar.pack(fill="x")

    tk.Label(top_bar, text="🎉 Đã Có Bản Cập Nhật Mới!", font=("Segoe UI", 14, "bold"), bg="#1e293b", fg="#38bdf8").pack(anchor="w")
    ver_text = f"Hiện tại: v{info['current_version']}  ➔  Mới nhất: v{info['latest_version']}"
    tk.Label(top_bar, text=ver_text, font=("Segoe UI", 10, "bold"), bg="#1e293b", fg="#a5f3fc").pack(anchor="w", pady=(2, 0))

    # Content Area
    content = tk.Frame(win, bg="#0f172a", padx=16, pady=12)
    content.pack(fill="both", expand=True)

    tk.Label(content, text="Thông tin bản phát hành:", font=("Segoe UI", 10, "bold"), bg="#0f172a", fg="#f8fafc").pack(anchor="w", pady=(0, 4))
    
    notes_box = tk.Text(content, bg="#090d16", fg="#e2e8f0", font=("Segoe UI", 9), wrap="word", height=8, padx=8, pady=8, relief="flat")
    notes_box.pack(fill="both", expand=True)
    notes_box.insert("1.0", f"{info.get('release_title', '')}\n\n{info.get('release_notes', '')}")
    notes_box.config(state="disabled")

    # Progress / Status label
    status_lbl = tk.Label(content, text="", font=("Segoe UI", 9, "italic"), bg="#0f172a", fg="#38bdf8")
    status_lbl.pack(fill="x", pady=(6, 0))

    progress_bar = ttk.Progressbar(content, mode="determinate")

    # Bottom Buttons
    btn_frame = tk.Frame(win, bg="#1e293b", padx=16, pady=10)
    btn_frame.pack(fill="x", side="bottom")

    def start_update():
        btn_update_now.config(state="disabled")
        btn_cancel.config(state="disabled")
        progress_bar.pack(fill="x", pady=(4, 0))
        status_lbl.config(text="Đang kết nối tải bản cập nhật...")

        def progress_cb(downloaded, total):
            if total > 0:
                pct = int(downloaded / total * 100)
                parent_root.after(0, lambda: [
                    progress_bar.config(value=pct),
                    status_lbl.config(text=f"Đang tải bản cập nhật: {pct}% ({downloaded // 1024} KB / {total // 1024} KB)")
                ])

        def run_dl():
            ok, msg = updater.download_and_install_update(
                download_url=info.get("download_url"),
                release_url=info.get("release_url", ""),
                progress_callback=progress_cb
            )
            if not ok:
                parent_root.after(0, lambda: [
                    progress_bar.pack_forget(),
                    status_lbl.config(text=msg, fg="#f43f5e"),
                    btn_update_now.config(state="normal", text="Mở trang tải trên Web"),
                    btn_update_now.config(command=lambda: updater.webbrowser.open(info.get("release_url", ""))),
                    btn_cancel.config(state="normal")
                ])
            else:
                parent_root.after(0, lambda: [
                    status_lbl.config(text=msg, fg="#4ade80"),
                    btn_cancel.config(state="normal", text="Đóng")
                ])

        threading.Thread(target=run_dl, daemon=True).start()

    btn_cancel = tk.Button(btn_frame, text="Để Sau", command=win.destroy, bg="#475569", fg="#ffffff", font=("Segoe UI", 9), relief="flat", padx=14, pady=6)
    btn_cancel.pack(side="right", padx=(6, 0))

    btn_update_now = tk.Button(btn_frame, text="🚀 Cập Nhật Ngay", command=start_update, bg="#16a34a", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=16, pady=6)
    btn_update_now.pack(side="right")
