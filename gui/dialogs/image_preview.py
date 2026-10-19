"""
gui/dialogs/image_preview.py - Flat Minimal Image Preview & Exercise Picture Matching Dialog.
Provides detailed large-image viewing and side-by-side exercise question-to-picture cross-checking
with zero component borders and authentic Flat Minimal Light UI styling.
"""

import os
import re
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, Tuple, List, Dict, Any
from PIL import Image, ImageTk

from gui.theme import Theme
from services.image_service import ensure_word_compatible_image
from core.utils import extract_question_prefix_and_body, natural_sort_key


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
    Opens an interactive Flat Minimal Light UI modal dialog displaying a large preview
    of the queue images with navigation, metadata, and zero borders.
    """
    if not parent_app.user_image_paths:
        messagebox.showinfo("Thông báo", "Danh sách Image Queue hiện đang trống.")
        return

    root = parent_app.root
    win = tk.Toplevel(root)
    win.title("Chi tiết hình ảnh ([PIC])")
    win.geometry("880x700")
    win.configure(bg=Theme.BG_CANVAS)
    win.transient(root)
    win.grab_set()

    try:
        x = root.winfo_x() + (root.winfo_width() // 2) - 440
        y = root.winfo_y() + (root.winfo_height() // 2) - 350
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    curr_idx = [max(0, min(initial_index, len(parent_app.user_image_paths) - 1))]
    current_photo_ref = [None]

    # ── Header Bar ───────────────────────────────────────────────
    header = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    header.pack(fill="x")

    lbl_title = tk.Label(header, text="Preview ảnh", font=Theme.FONT_TITLE, bg=Theme.BG_SURFACE, fg=Theme.TEXT_PRIMARY)
    lbl_title.pack(anchor="w")

    lbl_meta = tk.Label(header, text="", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY)
    lbl_meta.pack(anchor="w", pady=(2, 0))

    # ── Preview Canvas Card ──────────────────────────────────────
    center_frame = tk.Frame(win, bg=Theme.BG_CANVAS, padx=16, pady=10)
    center_frame.pack(fill="both", expand=True)

    card = tk.Frame(center_frame, bg=Theme.BG_SURFACE, padx=4, pady=4)
    card.pack(fill="both", expand=True)

    canvas = tk.Canvas(card, bg=Theme.BG_SURFACE_ALT, highlightthickness=0, bd=0)
    canvas.pack(fill="both", expand=True)

    # ── Bottom Control Bar ───────────────────────────────────────
    bot_bar = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    bot_bar.pack(fill="x", side="bottom")

    btn_prev = tk.Button(
        bot_bar,
        text="◀ Ảnh trước",
        bg=Theme.BG_SURFACE_ALT,
        fg=Theme.TEXT_PRIMARY,
        activebackground=Theme.BG_HOVER,
        font=Theme.FONT_BODY_BOLD,
        relief="flat",
        bd=0,
        padx=12,
        pady=5,
        cursor="hand2"
    )
    btn_prev.pack(side="left", padx=(0, 6))

    btn_next = tk.Button(
        bot_bar,
        text="Ảnh sau ▶",
        bg=Theme.BG_SURFACE_ALT,
        fg=Theme.TEXT_PRIMARY,
        activebackground=Theme.BG_HOVER,
        font=Theme.FONT_BODY_BOLD,
        relief="flat",
        bd=0,
        padx=12,
        pady=5,
        cursor="hand2"
    )
    btn_next.pack(side="left", padx=6)

    lbl_pos = tk.Label(bot_bar, text="", font=Theme.FONT_BODY_BOLD, bg=Theme.BG_SURFACE, fg=Theme.PRIMARY)
    lbl_pos.pack(side="left", padx=12)

    btn_close = tk.Button(
        bot_bar,
        text="Đóng (Esc)",
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
    btn_close.pack(side="right", padx=(6, 0))

    btn_open_default = tk.Button(
        bot_bar,
        text="Mở bằng App ngoài",
        bg=Theme.INFO,
        fg=Theme.TEXT_ON_PRIMARY,
        activebackground=Theme.INFO_HOVER,
        activeforeground=Theme.TEXT_ON_PRIMARY,
        font=Theme.FONT_BODY_BOLD,
        relief="flat",
        bd=0,
        padx=12,
        pady=5,
        cursor="hand2"
    )
    btn_open_default.pack(side="right", padx=6)

    def render_current_image():
        if not parent_app.user_image_paths:
            win.destroy()
            return

        idx = curr_idx[0]
        img_path = parent_app.user_image_paths[idx]
        fname = os.path.basename(img_path)
        total = len(parent_app.user_image_paths)

        lbl_title.config(text=f"[PIC #{idx + 1}] — {fname}")
        lbl_pos.config(text=f"Ảnh {idx + 1} / {total}")

        btn_prev.config(state="normal" if idx > 0 else "disabled")
        btn_next.config(state="normal" if idx < total - 1 else "disabled")

        canvas.update_idletasks()
        cw = max(200, canvas.winfo_width() - 20)
        ch = max(200, canvas.winfo_height() - 20)

        res = load_and_scale_image(img_path, max_w=cw, max_h=ch)
        canvas.delete("all")
        if res:
            photo, orig_w, orig_h, fmt, size_str = res
            current_photo_ref[0] = photo
            lbl_meta.config(text=f"Định dạng: {fmt}  |  Kích thước: {orig_w} × {orig_h} px  |  Dung lượng: {size_str}  |  Đường dẫn: {img_path}")
            canvas.create_image(canvas.winfo_width() // 2, canvas.winfo_height() // 2, anchor="center", image=photo)
        else:
            lbl_meta.config(text=f"Không thể giải mã file ảnh: {img_path}")
            canvas.create_text(
                canvas.winfo_width() // 2,
                canvas.winfo_height() // 2,
                text=f"⚠️ Không thể tải ảnh:\n{fname}",
                fill=Theme.DANGER,
                font=Theme.FONT_SECTION,
                justify="center"
            )

    def go_prev(event=None):
        if curr_idx[0] > 0:
            curr_idx[0] -= 1
            render_current_image()
            if hasattr(parent_app, 'image_queue'):
                parent_app.image_queue.img_listbox.selection_clear(0, tk.END)
                parent_app.image_queue.img_listbox.selection_set(curr_idx[0])
                parent_app.image_queue.img_listbox.see(curr_idx[0])
                parent_app.image_queue._render_thumbnail()

    def go_next(event=None):
        if curr_idx[0] < len(parent_app.user_image_paths) - 1:
            curr_idx[0] += 1
            render_current_image()
            if hasattr(parent_app, 'image_queue'):
                parent_app.image_queue.img_listbox.selection_clear(0, tk.END)
                parent_app.image_queue.img_listbox.selection_set(curr_idx[0])
                parent_app.image_queue.img_listbox.see(curr_idx[0])
                parent_app.image_queue._render_thumbnail()

    def open_external():
        if parent_app.user_image_paths:
            p = parent_app.user_image_paths[curr_idx[0]]
            try:
                os.startfile(p)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở file: {e}")

    btn_prev.config(command=go_prev)
    btn_next.config(command=go_next)
    btn_open_default.config(command=open_external)

    win.bind("<Left>", go_prev)
    win.bind("<Right>", go_next)
    win.bind("<Escape>", lambda e: win.destroy())

    win.after(50, render_current_image)


def extract_pic_questions_from_uln(uln_text: str) -> List[Dict[str, Any]]:
    """Extracts all questions and content blocks that contain [PIC] tags in document order."""
    from core.parser import ULNParser, parse_pic_tag
    blocks = ULNParser.parse(uln_text)
    pic_items = []
    current_section = ""

    def traverse(block_list):
        nonlocal current_section
        i = 0
        while i < len(block_list):
            b = block_list[i]
            tag = b.tag

            if tag in ["H1", "H2", "H3", "H4", "H5", "H6"]:
                current_section = b.content.strip()
                i += 1
                continue

            if tag == "PIC_GRID" and b.children:
                for idx_c, child in enumerate(b.children):
                    c_txt = child.content.strip()
                    pic_items.append({
                        "pic_index": len(pic_items) + 1,
                        "section": current_section,
                        "text": f"Lưới ảnh #{idx_c + 1}: {c_txt}",
                        "raw_line": c_txt
                    })
                i += 1
                continue

            if tag.startswith("TAB") and tag != "TABLE":
                if tag == "TAB2" or not b.cols or len(b.cols) == 2:
                    c1 = b.col1 or ""
                    c2 = b.col2 or ""
                    has_pic = ("[PIC" in c1.upper() or "[PIC" in c2.upper() or parse_pic_tag(c1) or parse_pic_tag(c2))
                    if has_pic:
                        c_clean = re.sub(r'\[PIC(?::[^\]]*)?\]', '', c1 + " " + c2, flags=re.IGNORECASE).strip()
                        pic_items.append({
                            "pic_index": len(pic_items) + 1,
                            "section": current_section,
                            "text": c_clean if c_clean else f"{c1} | {c2}",
                            "raw_line": f"{c1} | {c2}"
                        })
                    i += 1
                    continue

            if tag == "PIC":
                pic_items.append({
                    "pic_index": len(pic_items) + 1,
                    "section": current_section,
                    "text": b.content.strip() or "Hình ảnh độc lập",
                    "raw_line": b.content
                })
                i += 1
                continue

            if tag in ["P0", "P1", "P2", "P"]:
                m = re.search(r'\s*(\[PIC(?::[^\]]+)?\])\s*$', b.content, re.IGNORECASE)
                if m:
                    txt = b.content[:m.start()].strip()
                    pic_items.append({
                        "pic_index": len(pic_items) + 1,
                        "section": current_section,
                        "text": txt or b.content,
                        "raw_line": b.content
                    })
                    i += 1
                    continue

            i += 1

    traverse(blocks)
    return pic_items


def open_exercise_pic_match_dialog(parent_app):
    """Opens a streamlined Flat Minimal dialog mapping [PIC] tags side-by-side with images."""
    uln_text = parent_app.text_editor.get("1.0", tk.END).strip()
    pic_items = extract_pic_questions_from_uln(uln_text)

    root = parent_app.root
    win = tk.Toplevel(root)
    win.title("Đối chiếu câu hỏi & ảnh ([PIC])")
    win.geometry("980x720")
    win.configure(bg=Theme.BG_CANVAS)
    win.transient(root)
    win.grab_set()

    try:
        x = root.winfo_x() + (root.winfo_width() // 2) - 490
        y = root.winfo_y() + (root.winfo_height() // 2) - 360
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    header = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    header.pack(fill="x")

    tk.Label(header, text="Đối chiếu nhanh câu hỏi & ảnh ([PIC])", font=Theme.FONT_TITLE, bg=Theme.BG_SURFACE, fg=Theme.TEXT_PRIMARY).pack(anchor="w")

    num_pics = len(pic_items)
    num_imgs = len(parent_app.user_image_paths)
    sub_text = f"Số thẻ [PIC] trong đề: {num_pics}  |  Số ảnh trong Queue: {num_imgs}"
    lbl_sub = tk.Label(header, text=sub_text, font=Theme.FONT_SUBTITLE, bg=Theme.BG_SURFACE, fg=Theme.PRIMARY if num_pics == num_imgs else Theme.WARNING)
    lbl_sub.pack(anchor="w", pady=(3, 0))

    # Center Scrollable Cards
    center_container = tk.Frame(win, bg=Theme.BG_CANVAS, padx=16, pady=10)
    center_container.pack(fill="both", expand=True)

    canvas = tk.Canvas(center_container, bg=Theme.BG_CANVAS, highlightthickness=0, bd=0)
    scrollbar = ttk.Scrollbar(center_container, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg=Theme.BG_CANVAS)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas_win = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

    def _on_resize(e):
        canvas.itemconfig(canvas_win, width=e.width)
    canvas.bind("<Configure>", _on_resize)
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Render match cards
    photo_cache = []
    for idx, item in enumerate(pic_items):
        card = tk.Frame(scroll_frame, bg=Theme.BG_SURFACE, padx=12, pady=10)
        card.pack(fill="x", pady=4)

        # Left Info
        left_box = tk.Frame(card, bg=Theme.BG_SURFACE)
        left_box.pack(side="left", fill="both", expand=True, padx=(0, 12))

        q_title = f"Thẻ [PIC #{item['pic_index']}]"
        if item.get("section"):
            q_title += f" — {item['section']}"
        tk.Label(left_box, text=q_title, font=Theme.FONT_SECTION, bg=Theme.BG_SURFACE, fg=Theme.PRIMARY).pack(anchor="w")
        tk.Label(left_box, text=item["text"], font=Theme.FONT_BODY, bg=Theme.BG_SURFACE, fg=Theme.TEXT_PRIMARY, wraplength=540, justify="left").pack(anchor="w", pady=(4, 0))

        # Right Slot
        right_box = tk.Frame(card, bg=Theme.BG_SURFACE_ALT, padx=8, pady=8, width=220, height=130)
        right_box.pack(side="right")
        right_box.pack_propagate(False)

        if idx < len(parent_app.user_image_paths):
            img_p = parent_app.user_image_paths[idx]
            res = load_and_scale_image(img_p, max_w=200, max_h=90)
            if res:
                photo, w, h, fmt, sz = res
                photo_cache.append(photo)
                img_lbl = tk.Label(right_box, image=photo, bg=Theme.BG_SURFACE_ALT)
                img_lbl.pack(anchor="center")
                tk.Label(right_box, text=f"Ảnh #{idx + 1}: {os.path.basename(img_p)}", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).pack(anchor="center", pady=(2, 0))
            else:
                tk.Label(right_box, text=f"Lỗi nạp: {os.path.basename(img_p)}", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.DANGER).pack(anchor="center")
        else:
            tk.Label(right_box, text="[Chưa có ảnh trong Queue]", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.WARNING).pack(expand=True)

    # Keep photo references
    win._photo_cache = photo_cache

    bot_bar = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    bot_bar.pack(fill="x", side="bottom")

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
        padx=16,
        pady=5,
        cursor="hand2"
    )
    btn_close.pack(side="right")
