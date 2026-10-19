"""
gui/dialogs/docx_picker.py - Flat Minimal DOCX Explorer & Embedded ULN Importer.
Scans Word documents asynchronously in a background thread, displays embedded raw ULN metadata,
and allows instant 1-click re-import into the editor with zero component borders.
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from core.compiler import scan_folder_for_uln_docx
from gui.theme import Theme


def open_docx_picker_dialog(parent_app):
    """Opens a modern Flat Minimal Light UI dialog to scan and re-import embedded ULN from DOCX files."""
    root = parent_app.root
    win = tk.Toplevel(root)
    win.title("Nhập file DOCX có nhúng mã ULN")
    win.geometry("1060x680")
    win.configure(bg=Theme.BG_CANVAS)
    win.transient(root)
    win.grab_set()

    try:
        x = root.winfo_x() + (root.winfo_width() // 2) - 530
        y = root.winfo_y() + (root.winfo_height() // 2) - 340
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except Exception:
        pass

    current_folder = tk.StringVar(value=parent_app.last_output_dir if os.path.exists(parent_app.last_output_dir) else os.getcwd())
    only_uln_var = tk.BooleanVar(value=True)
    search_var = tk.StringVar(value="")
    all_items = []
    filtered_items = []

    # ── Top Header Bar ───────────────────────────────────────────
    top_bar = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    top_bar.pack(fill="x")

    tk.Label(
        top_bar,
        text="Trình duyệt & Nạp file DOCX nhúng mã nguồn ULN",
        font=Theme.FONT_TITLE,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_PRIMARY
    ).pack(anchor="w")

    tk.Label(
        top_bar,
        text="Tự động quét thư mục và phát hiện các file Word (.docx) chứa bản mã nguồn ULN gốc để nạp lại vào trình soạn thảo.",
        font=Theme.FONT_SUBTITLE,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_SECONDARY
    ).pack(anchor="w", pady=(2, 0))

    # ── Filter & Search Toolbar ──────────────────────────────────
    filter_bar = tk.Frame(win, bg=Theme.BG_SURFACE_ALT, padx=16, pady=8)
    filter_bar.pack(fill="x")

    f_row1 = tk.Frame(filter_bar, bg=Theme.BG_SURFACE_ALT)
    f_row1.pack(fill="x", pady=(0, 6))

    tk.Label(f_row1, text="Thư mục:", font=Theme.FONT_SMALL_BOLD, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).pack(side="left", padx=(0, 6))
    lbl_dir = tk.Label(f_row1, textvariable=current_folder, font=Theme.FONT_BODY, bg=Theme.BG_SURFACE, fg=Theme.PRIMARY, padx=8, pady=3, anchor="w", relief="flat")
    lbl_dir.pack(side="left", fill="x", expand=True, padx=(0, 8))

    def browse_folder():
        chosen = filedialog.askdirectory(initialdir=current_folder.get(), title="Chọn thư mục chứa file DOCX")
        if chosen and os.path.exists(chosen):
            current_folder.set(os.path.abspath(chosen))
            parent_app.last_output_dir = os.path.abspath(chosen)
            start_async_scan()

    btn_browse = tk.Button(
        f_row1,
        text="Đổi thư mục...",
        command=browse_folder,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_PRIMARY,
        activebackground=Theme.BG_HOVER,
        font=Theme.FONT_SMALL_BOLD,
        relief="flat",
        bd=0,
        padx=10,
        pady=3,
        cursor="hand2"
    )
    btn_browse.pack(side="right")

    f_row2 = tk.Frame(filter_bar, bg=Theme.BG_SURFACE_ALT)
    f_row2.pack(fill="x")

    chk_only = tk.Checkbutton(
        f_row2,
        text="Chỉ hiện file có nhúng Raw ULN",
        variable=only_uln_var,
        command=lambda: filter_and_render_tree(),
        bg=Theme.BG_SURFACE_ALT,
        fg=Theme.TEXT_PRIMARY,
        selectcolor=Theme.BG_SURFACE,
        activebackground=Theme.BG_SURFACE_ALT,
        font=Theme.FONT_BODY,
        bd=0,
        highlightthickness=0
    )
    chk_only.pack(side="left")

    search_frame = tk.Frame(f_row2, bg=Theme.BG_SURFACE_ALT)
    search_frame.pack(side="right")

    tk.Label(search_frame, text="Tìm kiếm:", font=Theme.FONT_SMALL_BOLD, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).pack(side="left", padx=(0, 4))
    search_entry = tk.Entry(
        search_frame,
        textvariable=search_var,
        font=Theme.FONT_BODY,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_PRIMARY,
        insertbackground=Theme.PRIMARY,
        width=22,
        relief="flat",
        bd=0,
        highlightthickness=0
    )
    search_entry.pack(side="left", ipady=2)
    search_var.trace_add("write", lambda *args: filter_and_render_tree())

    # ── Center Split: Treeview (Left) & Preview (Right) ──────────
    main_split = tk.Frame(win, bg=Theme.BG_CANVAS, padx=16, pady=10)
    main_split.pack(fill="both", expand=True)

    # Left: File List Treeview Card
    tree_card = tk.Frame(main_split, bg=Theme.BG_SURFACE, padx=2, pady=2)
    tree_card.pack(side="left", fill="both", expand=True, padx=(0, 8))

    columns = ("has_uln", "filename", "size_kb", "mtime")
    tree = ttk.Treeview(tree_card, columns=columns, show="headings", selectmode="browse")

    tree.heading("has_uln", text="Nhúng ULN")
    tree.heading("filename", text="Tên file DOCX")
    tree.heading("size_kb", text="Kích thước")
    tree.heading("mtime", text="Thời gian sửa đổi")

    tree.column("has_uln", width=95, anchor="center")
    tree.column("filename", width=260, anchor="w")
    tree.column("size_kb", width=85, anchor="e")
    tree.column("mtime", width=125, anchor="center")

    tree_scroll = ttk.Scrollbar(tree_card, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll.set)
    tree.pack(side="left", fill="both", expand=True)
    tree_scroll.pack(side="right", fill="y")

    # Right: Preview Card
    preview_card = tk.Frame(main_split, bg=Theme.BG_SURFACE, padx=12, pady=10, width=420)
    preview_card.pack(side="right", fill="both")
    preview_card.pack_propagate(False)

    tk.Label(
        preview_card,
        text="Xem trước nội dung ULN nhúng:",
        font=Theme.FONT_SECTION,
        bg=Theme.BG_SURFACE,
        fg=Theme.TEXT_PRIMARY
    ).pack(anchor="w", pady=(0, 6))

    txt_scroll = ttk.Scrollbar(preview_card)
    txt_scroll.pack(side="right", fill="y")

    preview_text = tk.Text(
        preview_card,
        wrap="word",
        bg=Theme.BG_SURFACE_ALT,
        fg=Theme.TEXT_PRIMARY,
        font=Theme.FONT_CODE,
        yscrollcommand=txt_scroll.set,
        padx=10,
        pady=8,
        bd=0,
        highlightthickness=0,
        relief="flat"
    )
    preview_text.pack(fill="both", expand=True)
    txt_scroll.config(command=preview_text.yview)

    # ── Bottom Action Bar ────────────────────────────────────────
    bot_bar = tk.Frame(win, bg=Theme.BG_SURFACE, padx=20, pady=12)
    bot_bar.pack(fill="x", side="bottom")

    status_lbl = tk.Label(bot_bar, text="", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY)
    status_lbl.pack(side="left")

    def show_msg(msg: str):
        status_lbl.config(text=msg)

    def on_tree_select(event=None):
        sel = tree.selection()
        if not sel:
            preview_text.delete("1.0", tk.END)
            return
        item_id = sel[0]
        try:
            idx = int(item_id)
            if 0 <= idx < len(filtered_items):
                item = filtered_items[idx]
                preview_text.delete("1.0", tk.END)
                if item["has_uln"]:
                    preview_text.insert("1.0", item["raw_uln"])
                    show_msg(f"✓ File: {item['filename']} (Có mã nhúng)")
                else:
                    preview_text.insert("1.0", "[File này không chứa mã nguồn Raw ULN nhúng]")
                    show_msg(f"File: {item['filename']} (Không có mã nhúng)")
        except Exception:
            pass

    tree.bind("<<TreeviewSelect>>", on_tree_select)
    tree.bind("<Double-Button-1>", lambda e: import_selected())

    def filter_and_render_tree():
        filtered_items.clear()
        query = search_var.get().strip().lower()
        only_uln = only_uln_var.get()

        for item in all_items:
            if only_uln and not item["has_uln"]:
                continue
            if query and (query not in item["filename"].lower()):
                continue
            filtered_items.append(item)

        tree.delete(*tree.get_children())
        for idx, item in enumerate(filtered_items):
            uln_badge = "✓ Có mã" if item["has_uln"] else "-"
            tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    uln_badge,
                    item["filename"],
                    f"{item['size_kb']:.1f} KB",
                    item["mtime"]
                )
            )

        if filtered_items:
            tree.selection_set("0")
            on_tree_select()
        else:
            preview_text.delete("1.0", tk.END)
            show_msg("Không tìm thấy file phù hợp.")

    def start_async_scan():
        target_dir = current_folder.get()
        show_msg("Đang quét thư mục...")
        tree.delete(*tree.get_children())

        def _worker():
            results = scan_folder_for_uln_docx(target_dir)
            def _apply():
                if not win.winfo_exists():
                    return
                all_items.clear()
                all_items.extend(results)
                filter_and_render_tree()
                show_msg(f"Quét xong: tìm thấy {len(all_items)} file DOCX.")
            win.after(0, _apply)

        threading.Thread(target=_worker, daemon=True).start()

    def import_selected():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Thông báo", "Vui lòng chọn một file DOCX từ danh sách.")
            return
        idx = int(sel[0])
        item = filtered_items[idx]
        if not item["has_uln"] or not item["raw_uln"]:
            messagebox.showwarning("Không có mã nhúng", f"File '{item['filename']}' không chứa mã nguồn Raw ULN nhúng bên trong.")
            return

        parent_app.text_editor.delete("1.0", tk.END)
        parent_app.text_editor.insert("1.0", item["raw_uln"])
        win.destroy()
        messagebox.showinfo("Thành công", f"Đã nạp thành công mã ULN từ file:\n{item['filename']}")

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

    btn_import = tk.Button(
        bot_bar,
        text="📥 Nạp Vào Editor",
        command=import_selected,
        bg=Theme.PRIMARY,
        fg=Theme.TEXT_ON_PRIMARY,
        activebackground=Theme.PRIMARY_HOVER,
        activeforeground=Theme.TEXT_ON_PRIMARY,
        font=Theme.FONT_BODY_BOLD,
        relief="flat",
        bd=0,
        padx=16,
        pady=5,
        cursor="hand2"
    )
    btn_import.pack(side="right")

    # Start initial async scan
    start_async_scan()
