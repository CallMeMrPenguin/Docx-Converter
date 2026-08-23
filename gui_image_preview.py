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
    card_registry = []

    # Drag & Drop State
    drag_data = {
        "active": False,
        "source_idx": None,
        "ghost_win": None,
        "target_idx": None,
        "drag_start_y": 0,
    }

    def on_drag_start(event, from_idx):
        if from_idx >= len(parent_app.user_image_paths):
            return
        drag_data["active"] = True
        drag_data["source_idx"] = from_idx
        drag_data["target_idx"] = from_idx
        drag_data["drag_start_y"] = event.y_root

        # Create floating semi-transparent drag preview window (Ghost)
        try:
            if drag_data["ghost_win"] and drag_data["ghost_win"].winfo_exists():
                drag_data["ghost_win"].destroy()
        except Exception:
            pass

        ghost = tk.Toplevel(win)
        ghost.overrideredirect(True)
        try:
            ghost.attributes("-alpha", 0.90)
            ghost.attributes("-topmost", True)
        except Exception:
            pass
        ghost.configure(bg="#0f172a", highlightthickness=2, highlightbackground="#38bdf8")

        img_p = parent_app.user_image_paths[from_idx]
        f_name = os.path.basename(img_p)
        g_thumb = load_and_scale_image(img_p, max_w=90, max_h=60)
        if g_thumb:
            g_photo, _, _, _, _ = g_thumb
            ghost.photo = g_photo
            lbl_g_img = tk.Label(ghost, image=g_photo, bg="#0f172a")
            lbl_g_img.pack(padx=4, pady=(4, 0))

        lbl_g_txt = tk.Label(
            ghost,
            text=f"Đang kéo [PIC #{from_idx + 1}]\n{f_name[:20]}...",
            font=("Segoe UI", 8, "bold"),
            bg="#1e293b",
            fg="#38bdf8",
            padx=6,
            pady=3
        )
        lbl_g_txt.pack(fill="x", padx=2, pady=2)

        ghost.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")
        drag_data["ghost_win"] = ghost

    def on_drag_motion(event):
        if not drag_data["active"]:
            return
        if drag_data["ghost_win"] and drag_data["ghost_win"].winfo_exists():
            drag_data["ghost_win"].geometry(f"+{event.x_root + 15}+{event.y_root + 15}")

        # Auto-scroll canvas when dragging near top or bottom edges
        win_y = canvas.winfo_rooty()
        win_h = canvas.winfo_height()
        if event.y_root < win_y + 40:
            canvas.yview_scroll(-1, "units")
        elif event.y_root > win_y + win_h - 40:
            canvas.yview_scroll(1, "units")

        # Detect which question card the cursor is hovering over
        hovered_idx = None
        for c_info in card_registry:
            c_widget = c_info["widget"]
            try:
                rx = c_widget.winfo_rootx()
                ry = c_widget.winfo_rooty()
                rw = c_widget.winfo_width()
                rh = c_widget.winfo_height()
                if rx <= event.x_root <= rx + rw and ry <= event.y_root <= ry + rh:
                    hovered_idx = c_info["idx"]
                    break
            except Exception:
                continue

        drag_data["target_idx"] = hovered_idx

        # Highlight target card with luminous border
        for c_info in card_registry:
            c_widget = c_info["widget"]
            try:
                if hovered_idx is not None and c_info["idx"] == hovered_idx:
                    c_widget.configure(highlightbackground="#38bdf8", highlightthickness=2)
                else:
                    c_widget.configure(highlightbackground="#334155", highlightthickness=1)
            except Exception:
                pass

    def on_drag_release(event):
        if not drag_data["active"]:
            return
        drag_data["active"] = False

        if drag_data["ghost_win"] and drag_data["ghost_win"].winfo_exists():
            drag_data["ghost_win"].destroy()
            drag_data["ghost_win"] = None

        src_idx = drag_data["source_idx"]
        tgt_idx = drag_data["target_idx"]

        # Reset card highlights
        for c_info in card_registry:
            try:
                c_info["widget"].configure(highlightbackground="#334155", highlightthickness=1)
            except Exception:
                pass

        if src_idx is not None and tgt_idx is not None and src_idx != tgt_idx:
            # Perform Image Swap / Move between src_idx and tgt_idx
            num_imgs = len(parent_app.user_image_paths)
            if 0 <= src_idx < num_imgs:
                if 0 <= tgt_idx < num_imgs:
                    parent_app.user_image_paths[src_idx], parent_app.user_image_paths[tgt_idx] = parent_app.user_image_paths[tgt_idx], parent_app.user_image_paths[src_idx]
                else:
                    # Target is unassigned card beyond user_image_paths length
                    val = parent_app.user_image_paths[src_idx]
                    parent_app.user_image_paths.pop(src_idx)
                    while len(parent_app.user_image_paths) < tgt_idx:
                        parent_app.user_image_paths.append("")
                    parent_app.user_image_paths.insert(tgt_idx, val)
                    parent_app.user_image_paths = [p for p in parent_app.user_image_paths if p]

                if hasattr(parent_app, 'update_image_listbox'):
                    parent_app.update_image_listbox(selected_index=min(tgt_idx, len(parent_app.user_image_paths) - 1))
                refresh_list()
                lbl_status.config(text=f"🔄 Đã chuyển ảnh từ vị trí #{src_idx + 1} sang Câu #{tgt_idx + 1} thành công!", fg="#38bdf8")

    def choose_image_for_slot(slot_idx):
        """Allows direct image selection from file picker specifically for this question slot."""
        filetypes = [
            ("All Supported Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff;*.webp;*.svg;*.wmf;*.emf"),
            ("PNG Images", "*.png"),
            ("JPEG Images", "*.jpg;*.jpeg"),
            ("SVG Vector Images", "*.svg"),
            ("All Files", "*.*")
        ]
        chosen = filedialog.askopenfilename(
            title=f"Chọn ảnh cho Câu #{slot_idx + 1}",
            filetypes=filetypes,
            parent=win
        )
        if chosen and os.path.exists(chosen):
            if slot_idx < len(parent_app.user_image_paths):
                parent_app.user_image_paths[slot_idx] = chosen
            else:
                while len(parent_app.user_image_paths) < slot_idx:
                    parent_app.user_image_paths.append("")
                parent_app.user_image_paths.append(chosen)
                parent_app.user_image_paths = [p for p in parent_app.user_image_paths if p]

            if hasattr(parent_app, 'update_image_listbox'):
                parent_app.update_image_listbox(selected_index=min(slot_idx, len(parent_app.user_image_paths) - 1))
            refresh_list()
            lbl_status.config(text=f"✅ Đã gán ảnh mới cho Câu #{slot_idx + 1}: {os.path.basename(chosen)}", fg="#4ade80")

    def remove_image_at_slot(slot_idx):
        """Removes image assigned to this question slot."""
        if 0 <= slot_idx < len(parent_app.user_image_paths):
            f_name = os.path.basename(parent_app.user_image_paths[slot_idx])
            parent_app.user_image_paths.pop(slot_idx)
            if hasattr(parent_app, 'update_image_listbox'):
                parent_app.update_image_listbox(selected_index=min(slot_idx, len(parent_app.user_image_paths) - 1) if parent_app.user_image_paths else None)
            refresh_list()
            lbl_status.config(text=f"🗑️ Đã xóa ảnh '{f_name}' khỏi Câu #{slot_idx + 1}", fg="#f43f5e")

    def refresh_list():
        photo_cache.clear()
        card_registry.clear()
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
            lbl_status.config(text=f"✅ Hoàn hảo! Khớp đủ {num_pics} / {num_pics} ảnh theo thứ tự câu hỏi. (Kéo thả ảnh giữa các câu để đổi vị trí)", fg="#4ade80")
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

            # Register card for drag & drop target detection
            card_registry.append({"widget": card, "idx": idx})

            # Left Box: Image & Drag Controls
            left_box = tk.Frame(card, bg="#1e293b", width=230)
            left_box.pack(side="left", fill="y", padx=(0, 12))

            # Badge & Drag Handle Bar
            badge_bar = tk.Frame(left_box, bg="#1e293b")
            badge_bar.pack(fill="x", anchor="w")

            lbl_badge = tk.Label(
                badge_bar,
                text=f"🖼️ [PIC #{it['pic_index']}]",
                font=("Segoe UI", 9, "bold"),
                bg="#1e293b",
                fg="#38bdf8"
            )
            lbl_badge.pack(side="left")

            lbl_drag_hint = tk.Label(
                badge_bar,
                text="⠿ Kéo thả",
                font=("Segoe UI", 7, "bold"),
                bg="#334155",
                fg="#94a3b8",
                padx=4,
                pady=1,
                cursor="fleur"
            )
            lbl_drag_hint.pack(side="right", padx=(4, 0))

            has_img = idx < len(images)
            if has_img:
                img_path = images[idx]
                fname = os.path.basename(img_path)

                # Thumbnail Canvas with Drag & Drop Event Bindings
                thumb_res = load_and_scale_image(img_path, max_w=130, max_h=85)
                if thumb_res:
                    photo, ow, oh, fmt, sz = thumb_res
                    photo_cache.append(photo)
                    c_thumb = tk.Canvas(left_box, width=130, height=85, bg="#0f172a", highlightthickness=1, highlightbackground="#475569", cursor="fleur")
                    c_thumb.pack(pady=4)
                    c_thumb.create_image(65, 42, image=photo, anchor="center")

                    # Bind Drag & Drop Events on Thumbnail and Drag Hint
                    for w in (c_thumb, lbl_drag_hint, lbl_badge):
                        w.bind("<ButtonPress-1>", lambda e, p_idx=idx: on_drag_start(e, p_idx))
                        w.bind("<B1-Motion>", on_drag_motion)
                        w.bind("<ButtonRelease-1>", on_drag_release)

                    # Double click to view large
                    c_thumb.bind("<Double-Button-1>", lambda e, p_idx=idx: open_image_preview_dialog(parent_app, initial_index=p_idx))

                    lbl_fname = tk.Label(left_box, text=f"{fname}\n({fmt}, {ow}×{oh})", font=("Segoe UI", 8), bg="#1e293b", fg="#94a3b8", justify="center", wraplength=140)
                    lbl_fname.pack()
                else:
                    lbl_err = tk.Label(left_box, text=f"⚠️ {fname}\n(Không đọc được ảnh)", font=("Segoe UI", 8), bg="#1e293b", fg="#ef4444")
                    lbl_err.pack(pady=4)

                # Action Button Row 1: Direct File Change & Delete
                action_row1 = tk.Frame(left_box, bg="#1e293b")
                action_row1.pack(fill="x", pady=(4, 2))

                btn_change = tk.Button(
                    action_row1,
                    text="📁 Đổi ảnh",
                    command=lambda p_idx=idx: choose_image_for_slot(p_idx),
                    bg="#0369a1",
                    fg="#ffffff",
                    activebackground="#0284c7",
                    font=("Segoe UI", 7, "bold"),
                    relief="flat",
                    padx=4,
                    pady=1,
                    cursor="hand2"
                )
                btn_change.pack(side="left", fill="x", expand=True, padx=(0, 1))

                btn_del = tk.Button(
                    action_row1,
                    text="❌ Xóa",
                    command=lambda p_idx=idx: remove_image_at_slot(p_idx),
                    bg="#475569",
                    fg="#fca5a5",
                    activebackground="#dc2626",
                    activeforeground="#ffffff",
                    font=("Segoe UI", 7, "bold"),
                    relief="flat",
                    padx=4,
                    pady=1,
                    cursor="hand2"
                )
                btn_del.pack(side="right", padx=(1, 0))

                # Action Button Row 2: Up / Down Reordering
                btn_row = tk.Frame(left_box, bg="#1e293b")
                btn_row.pack(fill="x", pady=(0, 0))

                def make_move_cmd(from_i, to_i):
                    def cmd():
                        if 0 <= to_i < len(parent_app.user_image_paths):
                            parent_app.user_image_paths[from_i], parent_app.user_image_paths[to_i] = parent_app.user_image_paths[to_i], parent_app.user_image_paths[from_i]
                            if hasattr(parent_app, 'update_image_listbox'):
                                parent_app.update_image_listbox(selected_index=to_i)
                            refresh_list()
                    return cmd

                if idx > 0:
                    btn_up = tk.Button(btn_row, text="▲ Lên", command=make_move_cmd(idx, idx - 1), bg="#334155", fg="#f8fafc", font=("Segoe UI", 7, "bold"), relief="flat", padx=4, pady=1, cursor="hand2")
                    btn_up.pack(side="left", fill="x", expand=True, padx=(0, 1))
                if idx < len(images) - 1:
                    btn_dn = tk.Button(btn_row, text="▼ Xuống", command=make_move_cmd(idx, idx + 1), bg="#334155", fg="#f8fafc", font=("Segoe UI", 7, "bold"), relief="flat", padx=4, pady=1, cursor="hand2")
                    btn_dn.pack(side="right", fill="x", expand=True, padx=(1, 0))

            else:
                lbl_missing = tk.Label(
                    left_box,
                    text="⚠️ CHƯA CÓ ẢNH\n(Kéo thả hoặc bấm nút bên dưới)",
                    font=("Segoe UI", 8, "bold"),
                    bg="#0f172a",
                    fg="#f59e0b",
                    padx=8,
                    pady=12,
                    highlightthickness=1,
                    highlightbackground="#f59e0b"
                )
                lbl_missing.pack(fill="x", pady=4)

                btn_add_slot = tk.Button(
                    left_box,
                    text="➕ Gán ảnh cho câu này",
                    command=lambda p_idx=idx: choose_image_for_slot(p_idx),
                    bg="#2563eb",
                    fg="#ffffff",
                    activebackground="#1d4ed8",
                    font=("Segoe UI", 8, "bold"),
                    relief="flat",
                    pady=3,
                    cursor="hand2"
                )
                btn_add_slot.pack(fill="x", pady=(2, 0))

            # Right Box: Question Context & Drop Target
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

            # Bind drag & drop onto right box and text as well for seamless drop targeting
            for w in (right_box, txt_q):
                w.bind("<B1-Motion>", on_drag_motion)
                w.bind("<ButtonRelease-1>", on_drag_release)

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

    lbl_drag_instruction = tk.Label(
        bot_bar,
        text="💡 Mẹo: Nhấn giữ và kéo ảnh từ câu này thả vào câu khác để hoán đổi vị trí lập tức.",
        font=("Segoe UI", 8, "italic"),
        bg="#1e293b",
        fg="#94a3b8"
    )
    lbl_drag_instruction.pack(side="left", padx=12)

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


