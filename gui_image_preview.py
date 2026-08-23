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


def extract_pic_questions_from_uln(uln_text: str):
    """
    Parses ULN text into AST blocks and extracts all [PIC] occurrences in the exact
    identical order that ULNWordRenderer processes them for MS Word document compilation.
    """
    import re
    from uln_parser import ULNParser, parse_pic_tag
    from renderer_utils import extract_question_prefix_and_body

    parser = ULNParser()
    blocks = parser.parse(uln_text)

    pic_items = []
    current_section = ""

    def traverse_blocks(block_list):
        nonlocal current_section
        i = 0
        while i < len(block_list):
            b = block_list[i]
            tag = b.tag.upper()

            # Track section instruction / headings
            if b.is_instruction or (b.spans and any(s.is_instruction for s in b.spans)) or tag.startswith("H"):
                clean_sec = re.sub(r'\[(?:ins|INS|P0|P1|P2|H[1-6])\]', '', b.content).replace('*', '').strip()
                if clean_sec:
                    current_section = clean_sec
            elif tag in ["P0", "P"] and re.match(r'^\s*(?:[IVXLCDM]+\.?|[A-Z]\.|\d+\.)\s+[A-Z]', b.content):
                clean_sec = b.content.replace('*', '').strip()
                if len(clean_sec) > 3:
                    current_section = clean_sec

            # 1. NUM Container
            if tag == "NUM":
                if b.children:
                    traverse_blocks(b.children)
                i += 1
                continue

            # 2. PIC_GRID
            if tag == "PIC_GRID":
                for idx_child, child in enumerate(b.children):
                    q_text = child.content.strip() if child.content else f"Grid Item #{idx_child + 1}"
                    pic_items.append({
                        "pic_index": len(pic_items) + 1,
                        "section": current_section,
                        "text": q_text,
                        "raw_line": child.content
                    })
                i += 1
                continue

            # 3. TAB2 group or Side-by-Side Sign MCQ
            if tag.startswith("TAB") and tag != "TABLE":
                num_cols = len(b.cols) if b.cols else (3 if tag == "TAB3" else (4 if tag == "TAB4" else 2))
                if num_cols == 2:
                    tab2_group = []
                    k = i
                    while k < len(block_list) and block_list[k].tag.startswith("TAB") and block_list[k].tag != "TABLE" and (not block_list[k].cols or len(block_list[k].cols) == 2):
                        tab2_group.append(block_list[k])
                        k += 1

                    # Check if single TAB2 block + OPT is a Side-by-Side Sign MCQ
                    if len(tab2_group) == 1:
                        single_tab = tab2_group[0]
                        c1_txt = single_tab.col1 or ""
                        c2_txt = single_tab.col2 or ""
                        has_pic = ("[PIC" in c1_txt.upper() or "[PIC" in c2_txt.upper() or parse_pic_tag(c1_txt) is not None or parse_pic_tag(c2_txt) is not None)
                        if has_pic and k < len(block_list) and block_list[k].tag == "OPT":
                            opt_block = block_list[k]
                            pref, delim, q_num, c_body = extract_question_prefix_and_body(c1_txt)
                            num_str = f"Câu #{q_num}: " if q_num is not None else ""
                            q_text = f"{num_str}{c_body.replace('[PIC]', '').replace('#', '').strip()}" if c_body else c1_txt.replace('[PIC]', '').replace('#', '').strip()
                            opt_text = " | ".join(opt_block.cols) if opt_block.cols else opt_block.content
                            opt_clean = opt_text.replace('*', '').strip()
                            display_text = f"{q_text}\n   -> {opt_clean}" if opt_clean else q_text
                            pic_items.append({
                                "pic_index": len(pic_items) + 1,
                                "section": current_section,
                                "text": display_text,
                                "raw_line": f"{single_tab.col1} | {single_tab.col2}"
                            })
                            i = k + 1
                            continue

                    # Process each block in tab2_group
                    for blk in tab2_group:
                        c1_txt = blk.col1 or ""
                        c2_txt = blk.col2 or ""
                        has_pic_c1 = ("[PIC" in c1_txt.upper() or parse_pic_tag(c1_txt) is not None)
                        has_pic_c2 = ("[PIC" in c2_txt.upper() or parse_pic_tag(c2_txt) is not None)

                        if has_pic_c1 or has_pic_c2:
                            if has_pic_c1:
                                pref, delim, q_num, c1_body = extract_question_prefix_and_body(c1_txt)
                                num_str = f"Câu #{q_num}: " if q_num is not None else ""
                                c2_clean = c2_txt.replace('<blank>', '___________').replace('[blank]', '___________').replace('#', '').strip()
                                display_text = f"{num_str}{c2_clean}" if c2_clean else c1_txt
                            else:
                                pref, delim, q_num, c1_body = extract_question_prefix_and_body(c1_txt)
                                num_str = f"Câu #{q_num}: " if q_num is not None else ""
                                display_text = f"{num_str}{c1_body.strip()}" if c1_body.strip() else c1_txt

                            pic_items.append({
                                "pic_index": len(pic_items) + 1,
                                "section": current_section,
                                "text": display_text,
                                "raw_line": f"{c1_txt} | {c2_txt}"
                            })

                    i = k
                    continue
                else:
                    # TAB3, TAB4
                    for c in (b.cols or []):
                        if "[PIC" in c.upper() or parse_pic_tag(c) is not None:
                            pic_items.append({
                                "pic_index": len(pic_items) + 1,
                                "section": current_section,
                                "text": c.replace('#', '').strip(),
                                "raw_line": c
                            })
                    i += 1
                    continue

            # 4. Standalone PIC block
            if tag == "PIC":
                pic_items.append({
                    "pic_index": len(pic_items) + 1,
                    "section": current_section,
                    "text": b.content.strip() if b.content else "Standalone Picture",
                    "raw_line": b.content
                })
                i += 1
                continue

            # 5. P0, P1, P2 blocks
            if tag in ["P0", "P1", "P2", "P"]:
                trailing_pic_match = re.search(r'\s*(\[PIC(?::[^\]]+)?\])\s*$', b.content, re.IGNORECASE)
                if trailing_pic_match:
                    text_part = b.content[:trailing_pic_match.start()].strip()
                    if i + 1 < len(block_list) and block_list[i + 1].tag == "OPT":
                        opt_blk = block_list[i + 1]
                        opt_text = " | ".join(opt_blk.cols) if opt_blk.cols else opt_blk.content
                        display_text = f"{text_part}\n   -> {opt_text.strip()}"
                        pic_items.append({
                            "pic_index": len(pic_items) + 1,
                            "section": current_section,
                            "text": display_text,
                            "raw_line": b.content
                        })
                        i += 2
                        continue
                    else:
                        pic_items.append({
                            "pic_index": len(pic_items) + 1,
                            "section": current_section,
                            "text": text_part if text_part else b.content,
                            "raw_line": b.content
                        })
                        i += 1
                        continue

                if b.spans:
                    for s in b.spans:
                        if s.text.startswith("[PIC:") or s.text.strip().upper() == "[PIC]":
                            pic_items.append({
                                "pic_index": len(pic_items) + 1,
                                "section": current_section,
                                "text": b.content.replace('#', '').strip(),
                                "raw_line": b.content
                            })

            # 6. TABLE
            if tag == "TABLE" and b.table_data:
                for row in b.table_data.rows:
                    for cell in row.cells:
                        if "[PIC" in cell.content.upper() or parse_pic_tag(cell.content) is not None:
                            pic_items.append({
                                "pic_index": len(pic_items) + 1,
                                "section": current_section,
                                "text": cell.content.replace('#', '').strip(),
                                "raw_line": cell.content
                            })

            i += 1

    traverse_blocks(blocks)
    return pic_items


def open_exercise_pic_match_dialog(parent_app):
    """
    Opens a streamlined, simple preview window showing all questions containing [PIC]
    mapped side-by-side with their queued image thumbnail to allow instant cross-checking.
    """
    import re
    from tkinter import filedialog

    uln_text = parent_app.text_editor.get("1.0", tk.END).strip()
    pic_items = extract_pic_questions_from_uln(uln_text)

    root = parent_app.root
    win = tk.Toplevel(root)
    win.title("🖼️ Đối Chiếu Nhanh Câu Hỏi & Ảnh ([PIC] Preview)")
    win.geometry("960x700")
    win.minsize(750, 500)
    win.configure(bg="#0f172a")
    win.transient(root)
    win.grab_set()

    try:
        x = root.winfo_x() + (root.winfo_width() // 2) - 480
        y = root.winfo_y() + (root.winfo_height() // 2) - 350
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    # Header Bar
    header = tk.Frame(win, bg="#1e293b", padx=16, pady=10)
    header.pack(fill="x")

    top_row = tk.Frame(header, bg="#1e293b")
    top_row.pack(fill="x")

    lbl_title = tk.Label(top_row, text="🖼️ Đối Chiếu Nhanh Câu Hỏi & Ảnh ([PIC])", font=("Segoe UI", 12, "bold"), bg="#1e293b", fg="#38bdf8")
    lbl_title.pack(side="left")

    btn_refresh = tk.Button(
        top_row,
        text="🔄 Làm Mới",
        command=lambda: refresh_list(),
        bg="#334155",
        fg="#f8fafc",
        activebackground="#475569",
        font=("Segoe UI", 8, "bold"),
        relief="flat",
        padx=10,
        pady=2,
        cursor="hand2"
    )
    btn_refresh.pack(side="right", padx=2)

    btn_sort = tk.Button(
        top_row,
        text="🔤 Sắp Xếp A-Z",
        command=lambda: [parent_app.sort_images(), refresh_list()],
        bg="#0284c7",
        fg="#ffffff",
        activebackground="#0369a1",
        font=("Segoe UI", 8, "bold"),
        relief="flat",
        padx=10,
        pady=2,
        cursor="hand2"
    )
    btn_sort.pack(side="right", padx=4)

    btn_add = tk.Button(
        top_row,
        text="➕ Thêm Ảnh...",
        command=lambda: [parent_app.add_images(), refresh_list()],
        bg="#2563eb",
        fg="#ffffff",
        activebackground="#1d4ed8",
        font=("Segoe UI", 8, "bold"),
        relief="flat",
        padx=10,
        pady=2,
        cursor="hand2"
    )
    btn_add.pack(side="right", padx=4)

    lbl_status = tk.Label(header, text="", font=("Segoe UI", 9), bg="#1e293b", fg="#94a3b8")
    lbl_status.pack(anchor="w", pady=(4, 0))

    # Scrollable Content Area
    center_container = tk.Frame(win, bg="#090d16")
    center_container.pack(fill="both", expand=True, padx=12, pady=8)

    canvas = tk.Canvas(center_container, bg="#090d16", highlightthickness=0)
    scrollbar = ttk.Scrollbar(center_container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg="#090d16")

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    def on_canvas_configure(event):
        canvas.itemconfig(canvas_window, width=event.width)

    canvas.bind("<Configure>", on_canvas_configure)
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Mousewheel scrolling
    def _on_mousewheel(event):
        if canvas.winfo_exists():
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    win.bind("<MouseWheel>", _on_mousewheel)

    # Cache for PhotoImages to prevent garbage collection
    photo_cache = []

    def refresh_list():
        photo_cache.clear()
        for widget in scrollable_frame.winfo_children():
            widget.destroy()

        current_uln = parent_app.text_editor.get("1.0", tk.END).strip()
        items = extract_pic_questions_from_uln(current_uln)
        images = parent_app.user_image_paths
        num_pics = len(items)
        num_imgs = len(images)

        # Update status header
        if num_pics == 0:
            lbl_status.config(text=f"ℹ️ Bài tập hiện tại không có thẻ [PIC] nào. (Queue có {num_imgs} ảnh)", fg="#94a3b8")
        elif num_imgs == num_pics:
            lbl_status.config(text=f"✅ Hoàn hảo! Khớp đủ {num_pics} / {num_pics} ảnh theo thứ tự câu hỏi.", fg="#4ade80")
        elif num_imgs < num_pics:
            lbl_status.config(text=f"⚠️ Có {num_pics} thẻ [PIC] nhưng Queue mới chỉ có {num_imgs} ảnh (thiếu {num_pics - num_imgs} ảnh).", fg="#fbbf24")
        else:
            lbl_status.config(text=f"ℹ️ Có {num_pics} thẻ [PIC] trong bài, Queue có {num_imgs} ảnh ({num_imgs - num_pics} ảnh sau sẽ không dùng).", fg="#38bdf8")

        if num_pics == 0:
            empty_lbl = tk.Label(
                scrollable_frame,
                text="Không tìm thấy thẻ [PIC] nào trong văn bản ULN.\nHãy thêm thẻ [PIC] vào bài tập (ví dụ: [TAB2] #1. [PIC] | You <blank> park here.)",
                font=("Segoe UI", 11),
                bg="#090d16",
                fg="#64748b",
                pady=40
            )
            empty_lbl.pack(fill="x")
            return

        for idx, it in enumerate(items):
            card = tk.Frame(scrollable_frame, bg="#1e293b", highlightthickness=1, highlightbackground="#334155", padx=10, pady=8)
            card.pack(fill="x", pady=5, padx=4)

            # Left Box: Image & Controls
            left_box = tk.Frame(card, bg="#1e293b", width=220)
            left_box.pack(side="left", fill="y", padx=(0, 12))

            lbl_badge = tk.Label(
                left_box,
                text=f"🖼️ [PIC #{it['pic_index']}]",
                font=("Segoe UI", 9, "bold"),
                bg="#1e293b",
                fg="#38bdf8"
            )
            lbl_badge.pack(anchor="w")

            has_img = idx < len(images)
            if has_img:
                img_path = images[idx]
                fname = os.path.basename(img_path)

                # Thumbnail Canvas
                thumb_res = load_and_scale_image(img_path, max_w=120, max_h=80)
                if thumb_res:
                    photo, ow, oh, fmt, sz = thumb_res
                    photo_cache.append(photo)
                    c_thumb = tk.Canvas(left_box, width=120, height=80, bg="#0f172a", highlightthickness=1, highlightbackground="#475569", cursor="hand2")
                    c_thumb.pack(pady=4)
                    c_thumb.create_image(60, 40, image=photo, anchor="center")
                    c_thumb.bind("<Button-1>", lambda e, p_idx=idx: open_image_preview_dialog(parent_app, initial_index=p_idx))

                    lbl_fname = tk.Label(left_box, text=f"{fname}\n({fmt}, {ow}×{oh})", font=("Segoe UI", 8), bg="#1e293b", fg="#94a3b8", justify="center", wraplength=130)
                    lbl_fname.pack()
                else:
                    lbl_err = tk.Label(left_box, text=f"⚠️ {fname}\n(Không đọc được ảnh)", font=("Segoe UI", 8), bg="#1e293b", fg="#ef4444")
                    lbl_err.pack(pady=4)

                # Mini Up/Down buttons for this image
                btn_row = tk.Frame(left_box, bg="#1e293b")
                btn_row.pack(pady=(4, 0))

                def make_move_cmd(from_i, to_i):
                    def cmd():
                        if 0 <= to_i < len(parent_app.user_image_paths):
                            parent_app.user_image_paths[from_i], parent_app.user_image_paths[to_i] = parent_app.user_image_paths[to_i], parent_app.user_image_paths[from_i]
                            if hasattr(parent_app, 'refresh_image_listbox'):
                                parent_app.refresh_image_listbox()
                            refresh_list()
                    return cmd

                if idx > 0:
                    btn_up = tk.Button(btn_row, text="▲ Lên", command=make_move_cmd(idx, idx - 1), bg="#334155", fg="#f8fafc", font=("Segoe UI", 7, "bold"), relief="flat", padx=4, pady=1, cursor="hand2")
                    btn_up.pack(side="left", padx=1)
                if idx < len(images) - 1:
                    btn_dn = tk.Button(btn_row, text="▼ Xuống", command=make_move_cmd(idx, idx + 1), bg="#334155", fg="#f8fafc", font=("Segoe UI", 7, "bold"), relief="flat", padx=4, pady=1, cursor="hand2")
                    btn_dn.pack(side="left", padx=1)

            else:
                lbl_missing = tk.Label(
                    left_box,
                    text="⚠️ CHƯA CÓ ẢNH\n(Sẽ dùng ảnh mẫu)",
                    font=("Segoe UI", 9, "bold"),
                    bg="#0f172a",
                    fg="#f59e0b",
                    padx=12,
                    pady=16,
                    highlightthickness=1,
                    highlightbackground="#f59e0b"
                )
                lbl_missing.pack(pady=4)

            # Right Box: Question Context
            right_box = tk.Frame(card, bg="#1e293b")
            right_box.pack(side="left", fill="both", expand=True)

            if it["section"]:
                lbl_sec = tk.Label(
                    right_box,
                    text=f"📌 {it['section']}",
                    font=("Segoe UI", 8, "italic"),
                    bg="#1e293b",
                    fg="#60a5fa",
                    anchor="w"
                )
                lbl_sec.pack(fill="x", pady=(0, 2))

            txt_q = tk.Text(
                right_box,
                height=3,
                bg="#090d16",
                fg="#f8fafc",
                font=("Segoe UI", 10),
                relief="flat",
                highlightthickness=1,
                highlightbackground="#334155",
                padx=8,
                pady=6,
                wrap="word"
            )
            txt_q.insert("1.0", it["text"])
            txt_q.config(state="disabled")
            txt_q.pack(fill="both", expand=True)

    # Bottom Control Bar
    bot_bar = tk.Frame(win, bg="#1e293b", padx=16, pady=10)
    bot_bar.pack(fill="x", side="bottom")

    btn_compile = tk.Button(
        bot_bar,
        text="🚀 COMPILE TO DOCX",
        command=lambda: [win.destroy(), parent_app.compile_docx()],
        bg="#16a34a",
        fg="#ffffff",
        activebackground="#15803d",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=16,
        pady=5,
        cursor="hand2"
    )
    btn_compile.pack(side="left")

    btn_close = tk.Button(
        bot_bar,
        text="Đóng (Esc)",
        command=win.destroy,
        bg="#475569",
        fg="#ffffff",
        activebackground="#64748b",
        font=("Segoe UI", 9),
        relief="flat",
        padx=14,
        pady=5,
        cursor="hand2"
    )
    btn_close.pack(side="right")

    win.bind("<Escape>", lambda e: win.destroy())

    refresh_list()

