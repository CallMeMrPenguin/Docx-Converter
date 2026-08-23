import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Tuple
from PIL import Image, ImageTk
from renderer_utils import ensure_word_compatible_image


def format_file_size(size_bytes: int) -> str:
    """Formats file size in bytes into human-readable B / KB / MB."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"


def load_and_scale_image(image_path: str, max_w: int = 240, max_h: int = 160) -> Optional[Tuple[ImageTk.PhotoImage, int, int, str, str]]:
    """
    Safely loads, converts (if SVG/HEIC/WebP/AVIF), and scales an image to fit max_w x max_h.
    Returns (PhotoImage, original_width, original_height, format_name, formatted_file_size) or None.
    """
    if not image_path or not os.path.exists(image_path):
        return None

    try:
        real_path = ensure_word_compatible_image(image_path) or image_path
        if not os.path.exists(real_path):
            return None

        stat = os.stat(image_path)
        size_str = format_file_size(stat.st_size)
        ext_str = os.path.splitext(image_path)[1].upper().lstrip('.')

        with Image.open(real_path) as img:
            orig_w, orig_h = img.size
            fmt = ext_str if ext_str else (img.format or "IMG")

            # Calculate scaled dimensions preserving aspect ratio
            ratio = min(max_w / max(1, orig_w), max_h / max(1, orig_h))
            new_w = max(1, int(orig_w * ratio))
            new_h = max(1, int(orig_h * ratio))

            resample_filter = getattr(Image, 'Resampling', Image).LANCZOS
            resized_img = img.resize((new_w, new_h), resample=resample_filter)
            photo = ImageTk.PhotoImage(resized_img)
            return photo, orig_w, orig_h, fmt, size_str
    except Exception as e:
        print(f"[ImagePreview] Error loading image {image_path}: {e}")
        return None


def open_image_preview_dialog(parent_app, initial_index: int = 0):
    """
    Opens an interactive dark modal dialog displaying a large preview of the queue images
    with navigation (Next/Prev), metadata, zoom, and open-in-viewer actions.
    """
    if not parent_app.user_image_paths:
        messagebox.showinfo("Thông báo", "Danh sách Image Queue hiện đang trống.")
        return

    root = parent_app.root
    win = tk.Toplevel(root)
    win.title("🖼️ Xem Chi Tiết Ảnh Queue ([PIC])")
    win.geometry("860x680")
    win.configure(bg="#0f172a")
    win.transient(root)
    win.grab_set()

    try:
        x = root.winfo_x() + (root.winfo_width() // 2) - 430
        y = root.winfo_y() + (root.winfo_height() // 2) - 340
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    curr_idx = [max(0, min(initial_index, len(parent_app.user_image_paths) - 1))]
    current_photo_ref = [None]

    # ── HEADER BAR ─────────────────────────────────────────────
    header = tk.Frame(win, bg="#1e293b", padx=16, pady=10)
    header.pack(fill="x")

    lbl_title = tk.Label(header, text="🖼️ Preview Ảnh", font=("Segoe UI", 12, "bold"), bg="#1e293b", fg="#38bdf8")
    lbl_title.pack(anchor="w")

    lbl_meta = tk.Label(header, text="", font=("Segoe UI", 9), bg="#1e293b", fg="#94a3b8")
    lbl_meta.pack(anchor="w", pady=(2, 0))

    # ── PREVIEW CANVAS ─────────────────────────────────────────
    center_frame = tk.Frame(win, bg="#090d16", padx=10, pady=10)
    center_frame.pack(fill="both", expand=True, padx=12, pady=10)

    canvas = tk.Canvas(center_frame, bg="#090d16", highlightthickness=1, highlightbackground="#334155")
    canvas.pack(fill="both", expand=True)

    # ── BOTTOM CONTROL BAR ─────────────────────────────────────
    bot_bar = tk.Frame(win, bg="#1e293b", padx=16, pady=10)
    bot_bar.pack(fill="x", side="bottom")

    btn_prev = tk.Button(bot_bar, text="◀ Ảnh Trước", bg="#334155", fg="#f8fafc", activebackground="#475569", activeforeground="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, cursor="hand2")
    btn_prev.pack(side="left", padx=(0, 6))

    btn_next = tk.Button(bot_bar, text="Ảnh Sau ▶", bg="#334155", fg="#f8fafc", activebackground="#475569", activeforeground="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=12, pady=5, cursor="hand2")
    btn_next.pack(side="left", padx=6)

    lbl_pos = tk.Label(bot_bar, text="", font=("Segoe UI", 10, "bold"), bg="#1e293b", fg="#38bdf8")
    lbl_pos.pack(side="left", padx=12)

    btn_close = tk.Button(bot_bar, text="Đóng (Esc)", command=win.destroy, bg="#475569", fg="#ffffff", activebackground="#64748b", activeforeground="#ffffff", font=("Segoe UI", 9), relief="flat", padx=14, pady=5, cursor="hand2")
    btn_close.pack(side="right", padx=(6, 0))

    btn_open_default = tk.Button(bot_bar, text="🖥 Mở Bằng App Ngoài", bg="#0284c7", fg="#ffffff", activebackground="#0369a1", activeforeground="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=5, cursor="hand2")
    btn_open_default.pack(side="right", padx=6)

    btn_open_dir = tk.Button(bot_bar, text="📂 Mở Thư Mục", bg="#475569", fg="#f8fafc", activebackground="#64748b", activeforeground="#ffffff", font=("Segoe UI", 9), relief="flat", padx=10, pady=5, cursor="hand2")
    btn_open_dir.pack(side="right", padx=6)

    def render_current_image():
        if not parent_app.user_image_paths:
            win.destroy()
            return

        idx = curr_idx[0]
        img_path = parent_app.user_image_paths[idx]
        fname = os.path.basename(img_path)
        total = len(parent_app.user_image_paths)

        lbl_title.config(text=f"🖼️ [PIC #{idx + 1}] — {fname}")
        lbl_pos.config(text=f"Ảnh {idx + 1} / {total}")

        # Update buttons state
        btn_prev.config(state="normal" if idx > 0 else "disabled")
        btn_next.config(state="normal" if idx < total - 1 else "disabled")

        # Determine canvas size
        canvas.update_idletasks()
        cw = max(200, canvas.winfo_width() - 20)
        ch = max(200, canvas.winfo_height() - 20)

        res = load_and_scale_image(img_path, max_w=cw, max_h=ch)
        canvas.delete("all")
        if res:
            photo, orig_w, orig_h, fmt, size_str = res
            current_photo_ref[0] = photo
            lbl_meta.config(text=f"Định dạng: {fmt}  |  Kích thước gốc: {orig_w} × {orig_h} px  |  Dung lượng: {size_str}  |  Vị trí: {img_path}")
            canvas.create_image(canvas.winfo_width() // 2, canvas.winfo_height() // 2, anchor="center", image=photo)
        else:
            lbl_meta.config(text=f"Không thể giải mã file ảnh: {img_path}")
            canvas.create_text(canvas.winfo_width() // 2, canvas.winfo_height() // 2, text=f"⚠️ Không thể tải ảnh:\n{fname}", fill="#ef4444", font=("Segoe UI", 11, "bold"), justify="center")

    def go_prev(event=None):
        if curr_idx[0] > 0:
            curr_idx[0] -= 1
            render_current_image()
            if hasattr(parent_app, 'img_listbox'):
                parent_app.img_listbox.selection_clear(0, tk.END)
                parent_app.img_listbox.selection_set(curr_idx[0])
                parent_app.img_listbox.see(curr_idx[0])
                parent_app.on_image_selected()

    def go_next(event=None):
        if curr_idx[0] < len(parent_app.user_image_paths) - 1:
            curr_idx[0] += 1
            render_current_image()
            if hasattr(parent_app, 'img_listbox'):
                parent_app.img_listbox.selection_clear(0, tk.END)
                parent_app.img_listbox.selection_set(curr_idx[0])
                parent_app.img_listbox.see(curr_idx[0])
                parent_app.on_image_selected()

    def open_external():
        if parent_app.user_image_paths:
            p = parent_app.user_image_paths[curr_idx[0]]
            try:
                os.startfile(p)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở file: {e}")

    def open_folder():
        if parent_app.user_image_paths:
            p = parent_app.user_image_paths[curr_idx[0]]
            try:
                subprocess.Popen(f'explorer /select,"{os.path.abspath(p)}"')
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở thư mục: {e}")

    btn_prev.config(command=go_prev)
    btn_next.config(command=go_next)
    btn_open_default.config(command=open_external)
    btn_open_dir.config(command=open_folder)

    win.bind("<Left>", go_prev)
    win.bind("<Right>", go_next)
    win.bind("<Escape>", lambda e: win.destroy())
    win.bind("<Configure>", lambda e: win.after_idle(render_current_image))

    win.after(100, render_current_image)
