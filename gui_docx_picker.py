import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from uln_compiler import scan_folder_for_uln_docx

def open_docx_picker_dialog(parent_app):
    """
    Opens a modern dark modal dialog to browse, scan, filter, and preview .docx files
    that contain embedded raw ULN data, allowing one-click re-import into the editor.
    """
    root = parent_app.root
    win = tk.Toplevel(root)
    win.title("📂 Trình Duyệt & Nhập File DOCX Có Nhúng ULN")
    win.geometry("1060x680")
    win.configure(bg="#0f172a")
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

    # ── TOP HEADER BAR ─────────────────────────────────────────────
    top_bar = tk.Frame(win, bg="#1e293b", padx=20, pady=12)
    top_bar.pack(fill="x")

    tk.Label(
        top_bar,
        text="📂 Trình Duyệt & Nạp File DOCX Đã Nhúng Mã Nguồn ULN",
        font=("Segoe UI", 13, "bold"),
        bg="#1e293b",
        fg="#38bdf8"
    ).pack(anchor="w")

    tk.Label(
        top_bar,
        text="Hệ thống tự động quét thư mục và lọc ra các file Word (.docx) chứa bản mã nguồn ULN gốc để bạn nạp lại vào trình soạn thảo.",
        font=("Segoe UI", 9),
        bg="#1e293b",
        fg="#94a3b8"
    ).pack(anchor="w", pady=(2, 0))

    # ── FOLDER & FILTER CONTROLS ──────────────────────────────────
    filter_bar = tk.Frame(win, bg="#111827", padx=16, pady=10)
    filter_bar.pack(fill="x")

    f_row1 = tk.Frame(filter_bar, bg="#111827")
    f_row1.pack(fill="x", pady=(0, 6))

    tk.Label(f_row1, text="📁 Thư mục quét:", font=("Segoe UI", 9, "bold"), bg="#111827", fg="#f8fafc").pack(side="left", padx=(0, 6))
    lbl_dir = tk.Label(f_row1, textvariable=current_folder, font=("Segoe UI", 9), bg="#090d16", fg="#38bdf8", padx=8, pady=3, anchor="w", relief="flat")
    lbl_dir.pack(side="left", fill="x", expand=True, padx=(0, 8))

    def browse_folder():
        chosen = filedialog.askdirectory(initialdir=current_folder.get(), title="Chọn thư mục chứa file DOCX")
        if chosen and os.path.exists(chosen):
            current_folder.set(os.path.abspath(chosen))
            parent_app.last_output_dir = os.path.abspath(chosen)
            refresh_list()

    btn_browse = tk.Button(
        f_row1,
        text="📁 Đổi Thư Mục...",
        command=browse_folder,
        bg="#334155",
        fg="#ffffff",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=10,
        pady=3,
        cursor="hand2"
    )
    btn_browse.pack(side="right")

    f_row2 = tk.Frame(filter_bar, bg="#111827")
    f_row2.pack(fill="x")

    chk_only = tk.Checkbutton(
        f_row2,
        text="Chỉ hiển thị các file DOCX có nhúng bản Raw ULN",
        variable=only_uln_var,
        command=lambda: filter_and_render_tree(),
        bg="#111827",
        fg="#38bdf8",
        selectcolor="#090d16",
        activebackground="#111827",
        activeforeground="#38bdf8",
        font=("Segoe UI", 9, "bold")
    )
    chk_only.pack(side="left")

    search_frame = tk.Frame(f_row2, bg="#111827")
    search_frame.pack(side="right")

    tk.Label(search_frame, text="🔍 Tìm tên file:", font=("Segoe UI", 9), bg="#111827", fg="#94a3b8").pack(side="left", padx=(0, 4))
    ent_search = tk.Entry(search_frame, textvariable=search_var, bg="#090d16", fg="#ffffff", insertbackground="#ffffff", font=("Segoe UI", 9), width=22)
    ent_search.pack(side="left")
    ent_search.bind("<KeyRelease>", lambda e: filter_and_render_tree())

    def clear_search_filter():
        search_var.set("")
        filter_and_render_tree()

    btn_clear_s = tk.Button(search_frame, text="✖", command=clear_search_filter, bg="#334155", fg="#f43f5e", font=("Segoe UI", 8, "bold"), relief="flat", padx=5, pady=1, cursor="hand2")
    btn_clear_s.pack(side="left", padx=(3, 0))

    # ── MAIN SPLIT PANE ────────────────────────────────────────────
    main_split = tk.Frame(win, bg="#0f172a", padx=16, pady=10)
    main_split.pack(fill="both", expand=True)

    # Left Column: File Treeview
    left_col = tk.Frame(main_split, bg="#1e293b", padx=8, pady=8)
    left_col.pack(side="left", fill="both", expand=True, padx=(0, 8))

    tk.Label(left_col, text="📄 Danh Sách File DOCX Đã Quét", font=("Segoe UI", 10, "bold"), bg="#1e293b", fg="#38bdf8").pack(anchor="w", pady=(0, 6))

    tree_frame = tk.Frame(left_col, bg="#1e293b")
    tree_frame.pack(fill="both", expand=True)

    tree_scroll = ttk.Scrollbar(tree_frame)
    tree_scroll.pack(side="right", fill="y")

    columns = ("mtime", "size", "status")
    tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", selectmode="browse", yscrollcommand=tree_scroll.set)
    tree.heading("#0", text="Tên File DOCX", anchor="w")
    tree.heading("mtime", text="Thời Gian", anchor="center")
    tree.heading("size", text="Dung Lượng", anchor="center")
    tree.heading("status", text="Trạng Thái", anchor="center")

    tree.column("#0", width=280, minwidth=180, stretch=True)
    tree.column("mtime", width=120, minwidth=100, stretch=False, anchor="center")
    tree.column("size", width=80, minwidth=70, stretch=False, anchor="center")
    tree.column("status", width=100, minwidth=90, stretch=False, anchor="center")

    tree.tag_configure("has_uln", foreground="#4ade80")
    tree.tag_configure("no_uln", foreground="#94a3b8")

    tree.pack(fill="both", expand=True)
    tree_scroll.config(command=tree.yview)

    # Right Column: Preview Pane
    right_col = tk.Frame(main_split, bg="#1e293b", padx=8, pady=8, width=420)
    right_col.pack(side="right", fill="both", expand=False)
    right_col.pack_propagate(False)

    right_top = tk.Frame(right_col, bg="#1e293b")
    right_top.pack(fill="x", pady=(0, 6))

    tk.Label(right_top, text="👁️ Xem Trước Mã ULN (Raw)", font=("Segoe UI", 10, "bold"), bg="#1e293b", fg="#38bdf8").pack(side="left")
    preview_meta_lbl = tk.Label(right_top, text="", font=("Segoe UI", 8), bg="#1e293b", fg="#94a3b8")
    preview_meta_lbl.pack(side="right")

    prev_scroll = ttk.Scrollbar(right_col)
    prev_scroll.pack(side="right", fill="y")

    preview_text = tk.Text(
        right_col,
        wrap="none",
        bg="#090d16",
        fg="#e2e8f0",
        insertbackground="#ffffff",
        font=("Consolas", 9),
        yscrollcommand=prev_scroll.set,
        padx=8,
        pady=8
    )
    preview_text.pack(fill="both", expand=True)
    prev_scroll.config(command=preview_text.yview)

    # ── BOTTOM FOOTER BAR ──────────────────────────────────────────
    footer_bar = tk.Frame(win, bg="#1e293b", padx=16, pady=10)
    footer_bar.pack(fill="x", side="bottom")

    status_count_lbl = tk.Label(footer_bar, text="Đang quét thư mục...", font=("Segoe UI", 9), bg="#1e293b", fg="#94a3b8")
    status_count_lbl.pack(side="left")

    btn_close = tk.Button(
        footer_bar,
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
    btn_close.pack(side="right", padx=(8, 0))

    def load_selected_to_editor():
        sel_items = tree.selection()
        if not sel_items:
            messagebox.showwarning("Thông báo", "Vui lòng chọn một file DOCX có nhúng ULN từ danh sách.")
            return
        item_id = sel_items[0]
        f_data = file_data_map.get(item_id)
        if not f_data or not f_data.get("has_uln"):
            messagebox.showwarning("Thông báo", "File DOCX đã chọn không chứa mã nguồn ULN.")
            return

        raw_code = f_data.get("raw_uln", "")
        parent_app.text_editor.delete("1.0", tk.END)
        parent_app.text_editor.insert("1.0", raw_code)
        win.destroy()
        messagebox.showinfo("Thành công", f"Đã nạp thành công mã nguồn ULN từ file:\n{f_data.get('filename')}")

    def embed_current_uln_to_selected():
        sel_items = tree.selection()
        if not sel_items:
            messagebox.showwarning("Thông báo", "Vui lòng chọn một file DOCX từ danh sách.")
            return
        item_id = sel_items[0]
        f_data = file_data_map.get(item_id)
        if not f_data:
            return

        current_uln = parent_app.text_editor.get("1.0", tk.END).strip()
        if not current_uln:
            messagebox.showwarning("Thông báo", "Trình soạn thảo hiện tại đang trống. Vui lòng nhập mã ULN trước khi nhúng.")
            return

        fpath = f_data["filepath"]
        from uln_compiler import embed_raw_uln_docx
        ok = embed_raw_uln_docx(fpath, current_uln)
        if ok:
            messagebox.showinfo("Thành công", f"Đã nhúng thành công mã Raw ULN vào file:\n{f_data['filename']}")
            refresh_list()
        else:
            messagebox.showerror("Lỗi", "Không thể nhúng mã ULN vào file (vui lòng đảm bảo file đã được đóng trong Word).")

    btn_embed = tk.Button(
        footer_bar,
        text="💾 Nhúng mã ULN hiện tại vào file này",
        command=embed_current_uln_to_selected,
        bg="#0284c7",
        fg="#ffffff",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=14,
        pady=4,
        cursor="hand2"
    )
    btn_embed.pack(side="right", padx=(8, 0))

    btn_load = tk.Button(
        footer_bar,
        text="📥 Nạp bản Raw vào Trình Soạn Thảo (Load ULN)",
        command=load_selected_to_editor,
        bg="#16a34a",
        fg="#ffffff",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=16,
        pady=4,
        cursor="hand2",
        state="disabled"
    )
    btn_load.pack(side="right", padx=(8, 0))

    btn_refresh = tk.Button(
        footer_bar,
        text="🔄 Quét Lại",
        command=lambda: refresh_list(),
        bg="#334155",
        fg="#38bdf8",
        font=("Segoe UI", 9, "bold"),
        relief="flat",
        padx=12,
        pady=4,
        cursor="hand2"
    )
    btn_refresh.pack(side="right")

    # ── LOGIC FOR SCANNING & FILTERING ────────────────────────────
    scanned_files_cache = []
    file_data_map = {}

    def refresh_list():
        target_dir = current_folder.get()
        status_count_lbl.config(text="⏳ Đang quét thư mục...", fg="#38bdf8")
        win.update()
        
        scanned = scan_folder_for_uln_docx(target_dir)
        scanned_files_cache.clear()
        scanned_files_cache.extend(scanned)
        filter_and_render_tree()

    def filter_and_render_tree():
        tree.delete(*tree.get_children())
        file_data_map.clear()
        preview_text.delete("1.0", tk.END)
        preview_meta_lbl.config(text="")
        btn_load.config(state="disabled")

        only_uln = only_uln_var.get()
        query = search_var.get().strip().lower()

        total_docx = len(scanned_files_cache)
        uln_count = 0
        matching_count = 0

        for f in scanned_files_cache:
            if f["has_uln"]:
                uln_count += 1
            
            # Check filter
            if only_uln and not f["has_uln"]:
                continue
            if query and (query not in f["filename"].lower()):
                continue

            matching_count += 1
            status_text = "✓ Có ULN" if f["has_uln"] else "- Không có"
            tag_name = "has_uln" if f["has_uln"] else "no_uln"
            icon_prefix = "📄 "

            item_id = tree.insert(
                "",
                "end",
                text=f"{icon_prefix}{f['filename']}",
                values=(f["mtime"], f"{f['size_kb']:.1f} KB", status_text),
                tags=(tag_name,)
            )
            file_data_map[item_id] = f

        status_count_lbl.config(
            text=f"Tìm thấy {uln_count} file có ULN (Hiển thị {matching_count} / {total_docx} files)",
            fg="#4ade80" if uln_count > 0 else "#94a3b8"
        )

        children = tree.get_children()
        if children:
            tree.selection_set(children[0])
            on_tree_select(None)

    def on_tree_select(event):
        sel = tree.selection()
        if not sel:
            btn_load.config(state="disabled")
            preview_text.delete("1.0", tk.END)
            preview_meta_lbl.config(text="")
            return

        item_id = sel[0]
        f_data = file_data_map.get(item_id)
        if not f_data:
            return

        preview_text.delete("1.0", tk.END)
        if f_data.get("has_uln"):
            raw_code = f_data.get("raw_uln", "")
            preview_text.insert("1.0", raw_code)
            lines = raw_code.count('\n') + 1
            chars = len(raw_code)
            preview_meta_lbl.config(text=f"{lines} dòng | {chars} ký tự", fg="#4ade80")
            btn_load.config(state="normal")
        else:
            preview_text.insert("1.0", "// File DOCX này là tài liệu Word thông thường,\n// không chứa bản mã nguồn ULN gốc được nhúng.")
            preview_meta_lbl.config(text="Không có mã ULN", fg="#f43f5e")
            btn_load.config(state="disabled")

    tree.bind("<<TreeviewSelect>>", on_tree_select)
    tree.bind("<Double-1>", lambda e: load_selected_to_editor())

    refresh_list()
