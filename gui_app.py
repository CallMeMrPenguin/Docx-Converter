import os
import sys
import re
import threading
import subprocess
from typing import Optional, List, Dict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from uln_compiler import ULNCompiler, extract_raw_uln, has_embedded_uln, scan_folder_for_uln_docx
import updater
from gui_styles import setup_dark_theme
from gui_prompt_editor import (
    open_prompt_editor_dialog,
    load_prompt_text,
    save_prompt_text,
    reset_prompt_text,
    get_prompt_storage_path,
    get_default_prompt_path
)
from gui_update_modal import show_update_modal_dialog
from gui_docx_picker import open_docx_picker_dialog
from gui_image_preview import load_and_scale_image, open_image_preview_dialog
from renderer_utils import natural_sort_key

class ULNFormatterApp:
    def __init__(self, root):
        self.root = root
        self.version = updater.get_current_version()
        self.root.title(f"Universal Layout Notation (ULN) → DOCX Formatter v{self.version}")
        self.root.geometry("1220x840")
        self.root.minsize(1050, 700)
        self.root.configure(bg="#0f172a")

        self.last_output_dir = os.path.expanduser("~\\Documents")
        if not os.path.exists(self.last_output_dir):
            self.last_output_dir = os.getcwd()

        # Configure dark theme styling via gui_styles
        setup_dark_theme(self.root)

        # Top Title Bar
        header_frame = tk.Frame(self.root, bg="#1e293b", height=60, padx=20, pady=10)
        header_frame.pack(fill="x", side="top")
        
        title_left = tk.Frame(header_frame, bg="#1e293b")
        title_left.pack(side="left", fill="y")

        title_label = tk.Label(title_left, text="ULN to DOCX Custom Formatter", font=("Segoe UI", 16, "bold"), bg="#1e293b", fg="#38bdf8")
        title_label.pack(side="left")

        ver_badge = tk.Label(title_left, text=f"v{self.version}", font=("Segoe UI", 9, "bold"), bg="#0f172a", fg="#38bdf8", padx=6, pady=2, relief="flat")
        ver_badge.pack(side="left", padx=8)

        subtitle = tk.Label(title_left, text="Powered by pywin32 COM Automation Engine", font=("Segoe UI", 9, "italic"), bg="#1e293b", fg="#94a3b8")
        subtitle.pack(side="left", padx=10)

        # Top Right Actions (Prompt & Update buttons)
        header_right = tk.Frame(header_frame, bg="#1e293b")
        header_right.pack(side="right", fill="y")

        self.btn_prompt = tk.Button(
            header_right,
            text="📜 AI Prompt / Quy tắc ULN",
            command=self.open_prompt_editor,
            bg="#0369a1",
            fg="#ffffff",
            activebackground="#0284c7",
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2"
        )
        self.btn_prompt.pack(side="left", padx=(0, 8))

        self.btn_update = tk.Button(
            header_right,
            text="🔄 Kiểm tra Cập nhật",
            command=self.manual_check_updates,
            bg="#334155",
            fg="#38bdf8",
            activebackground="#0284c7",
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2"
        )
        self.btn_update.pack(side="right")

        # Main Layout: Left Sidebar (Customization Settings), Right Area (ULN Text Editor)
        main_container = tk.Frame(self.root, bg="#0f172a")
        main_container.pack(fill="both", expand=True, padx=15, pady=15)

        # Left Sidebar Frame (Settings)
        sidebar = ttk.LabelFrame(main_container, text=" Customization Settings ", padding=12)
        sidebar.pack(side="left", fill="y", padx=(0, 10))

        # Font Settings
        ttk.Label(sidebar, text="Font Name:").pack(anchor="w", pady=(5, 2))
        self.font_var = tk.StringVar(value="Times New Roman")
        font_cb = ttk.Combobox(sidebar, textvariable=self.font_var, values=["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Garamond"], width=20)
        font_cb.pack(fill="x", pady=(0, 10))

        ttk.Label(sidebar, text="Font Size (pt):").pack(anchor="w", pady=(5, 2))
        self.size_var = tk.DoubleVar(value=12.0)
        size_sp = ttk.Spinbox(sidebar, from_=8.0, to=24.0, increment=0.5, textvariable=self.size_var, width=20)
        size_sp.pack(fill="x", pady=(0, 10))

        # Margins Settings
        ttk.Label(sidebar, text="Margins (cm):", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 3))
        
        m_frame = tk.Frame(sidebar, bg="#1e293b")
        m_frame.pack(fill="x", pady=2)
        
        tk.Label(m_frame, text="Top:", bg="#1e293b", fg="#f8fafc").grid(row=0, column=0, sticky="w", pady=2)
        self.m_top_var = tk.DoubleVar(value=2.0)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_top_var, width=5, bg="#090d16", fg="#f8fafc", insertbackground="#ffffff", buttonbackground="#1e293b").grid(row=0, column=1, padx=4, pady=2)

        tk.Label(m_frame, text="Bottom:", bg="#1e293b", fg="#f8fafc").grid(row=0, column=2, sticky="w", pady=2, padx=(6, 0))
        self.m_bottom_var = tk.DoubleVar(value=2.0)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_bottom_var, width=5, bg="#090d16", fg="#f8fafc", insertbackground="#ffffff", buttonbackground="#1e293b").grid(row=0, column=3, padx=4, pady=2)

        tk.Label(m_frame, text="Left:", bg="#1e293b", fg="#f8fafc").grid(row=1, column=0, sticky="w", pady=2)
        self.m_left_var = tk.DoubleVar(value=3.0)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_left_var, width=5, bg="#090d16", fg="#f8fafc", insertbackground="#ffffff", buttonbackground="#1e293b").grid(row=1, column=1, padx=4, pady=2)

        tk.Label(m_frame, text="Right:", bg="#1e293b", fg="#f8fafc").grid(row=1, column=2, sticky="w", pady=2, padx=(6, 0))
        self.m_right_var = tk.DoubleVar(value=1.5)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_right_var, width=5, bg="#090d16", fg="#f8fafc", insertbackground="#ffffff", buttonbackground="#1e293b").grid(row=1, column=3, padx=4, pady=2)

        # Question & Option Numbering / Formatting Frame
        q_frame = ttk.LabelFrame(sidebar, text=" 🎯 Question, Option & Ins Styling ", padding=8)
        q_frame.pack(fill="x", pady=(8, 6))

        # Question Prefix (Default: "")
        ttk.Label(q_frame, text="Prefix:").grid(row=0, column=0, sticky="w", pady=2)
        self.q_prefix_var = tk.StringVar(value="")
        q_pref_cb = ttk.Combobox(q_frame, textvariable=self.q_prefix_var, values=["", "Question ", "Câu ", "Task ", "Exercise ", "Ex ", "Activity "], width=16)
        q_pref_cb.grid(row=0, column=1, padx=4, pady=2, sticky="ew")

        # Question Delimiter (Default: ".")
        ttk.Label(q_frame, text="Delimiter:").grid(row=1, column=0, sticky="w", pady=2)
        self.q_delim_var = tk.StringVar(value=".")
        q_delim_cb = ttk.Combobox(q_frame, textvariable=self.q_delim_var, values=[".", ":", ")", "-"], width=16)
        q_delim_cb.grid(row=1, column=1, padx=4, pady=2, sticky="ew")

        # Color Options Palette
        color_choices = [
            "Default (Black) #000000",
            "Blue #2563eb",
            "Navy #1e40af",
            "Red #dc2626",
            "Emerald #059669",
            "Purple #7c3aed",
            "Amber #d97706",
            "Dark Grey #475569"
        ]

        # Question Number Color
        ttk.Label(q_frame, text="Num Color:").grid(row=2, column=0, sticky="w", pady=2)
        self.q_color_var = tk.StringVar(value="Default (Black) #000000")
        q_col_cb = ttk.Combobox(q_frame, textvariable=self.q_color_var, values=color_choices, width=16)
        q_col_cb.grid(row=2, column=1, padx=4, pady=2, sticky="ew")

        # Option Letter (ABCD) Color
        ttk.Label(q_frame, text="Opt Color:").grid(row=3, column=0, sticky="w", pady=2)
        self.opt_color_var = tk.StringVar(value="Default (Black) #000000")
        opt_col_cb = ttk.Combobox(q_frame, textvariable=self.opt_color_var, values=color_choices, width=16)
        opt_col_cb.grid(row=3, column=1, padx=4, pady=2, sticky="ew")

        # Instruction Heading ([ins]) Color
        ttk.Label(q_frame, text="Ins Color:").grid(row=4, column=0, sticky="w", pady=2)
        self.ins_color_var = tk.StringVar(value="Default (Black) #000000")
        ins_col_cb = ttk.Combobox(q_frame, textvariable=self.ins_color_var, values=color_choices, width=16)
        ins_col_cb.grid(row=4, column=1, padx=4, pady=2, sticky="ew")

        # Page Numbering Checkbox
        self.pg_num_var = tk.BooleanVar(value=True)
        pg_check = tk.Checkbutton(sidebar, text="Include Page Numbers (Page X / Y)", variable=self.pg_num_var, bg="#1e293b", fg="#f8fafc", selectcolor="#0f172a", activebackground="#1e293b", activeforeground="#f8fafc")
        pg_check.pack(anchor="w", pady=(6, 4))

        # Open in Word after compile
        self.open_word_var = tk.BooleanVar(value=True)
        open_check = tk.Checkbutton(sidebar, text="Open DOCX in Word after saving", variable=self.open_word_var, bg="#1e293b", fg="#f8fafc", selectcolor="#0f172a", activebackground="#1e293b", activeforeground="#f8fafc")
        open_check.pack(anchor="w", pady=(0, 4))

        # Background Mode Checkbox (Silent, 0 Mouse Interruption)
        self.bg_mode_var = tk.BooleanVar(value=False)
        bg_check = tk.Checkbutton(sidebar, text="⚡ Background Mode (0 Mouse Freeze)", variable=self.bg_mode_var, bg="#1e293b", fg="#38bdf8", selectcolor="#0f172a", activebackground="#1e293b", activeforeground="#38bdf8")
        bg_check.pack(anchor="w", pady=(0, 6))

        # Image Queue (Order for [PIC] tags) with Live Preview
        self.user_image_paths = []
        self._current_preview_photo = None
        img_frame = ttk.LabelFrame(sidebar, text=" 🖼️ Image Queue ([PIC] Thứ tự ảnh) ", padding=6)
        img_frame.pack(fill="x", pady=(2, 4))

        # Listbox with Scrollbar
        list_container = tk.Frame(img_frame, bg="#090d16")
        list_container.pack(fill="x", pady=(0, 4))

        self.img_listbox = tk.Listbox(
            list_container,
            height=4,
            bg="#090d16",
            fg="#e2e8f0",
            font=("Segoe UI", 9),
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            activestyle="none",
            highlightthickness=0,
            relief="flat"
        )
        self.img_listbox.pack(side="left", fill="both", expand=True)

        list_scroll = ttk.Scrollbar(list_container, orient="vertical", command=self.img_listbox.yview)
        list_scroll.pack(side="right", fill="y")
        self.img_listbox.config(yscrollcommand=list_scroll.set)

        self.img_listbox.bind("<<ListboxSelect>>", self.on_image_selected)
        self.img_listbox.bind("<Double-Button-1>", lambda e: self.preview_selected_image())
        self.img_listbox.bind("<KeyRelease-Up>", self.on_image_selected)
        self.img_listbox.bind("<KeyRelease-Down>", self.on_image_selected)

        # Button Bar Row 1: Add, Sort, Clear
        img_btn_bar = tk.Frame(img_frame, bg="#1e293b")
        img_btn_bar.pack(fill="x", pady=(0, 2))

        btn_add_img = tk.Button(img_btn_bar, text="➕ Thêm...", command=self.add_images, bg="#2563eb", fg="#ffffff", activebackground="#1d4ed8", font=("Segoe UI", 8, "bold"), relief="flat", pady=2, cursor="hand2")
        btn_add_img.pack(side="left", fill="x", expand=True, padx=(0, 2))

        btn_sort_img = tk.Button(img_btn_bar, text="🔤 A-Z", command=self.sort_images, bg="#0284c7", fg="#ffffff", activebackground="#0369a1", font=("Segoe UI", 8, "bold"), relief="flat", pady=2, cursor="hand2")
        btn_sort_img.pack(side="left", padx=2)

        btn_clear_img = tk.Button(img_btn_bar, text="🗑️ Xóa hết", command=self.clear_images, bg="#475569", fg="#ffffff", activebackground="#64748b", font=("Segoe UI", 8), relief="flat", pady=2, cursor="hand2")
        btn_clear_img.pack(side="right", padx=(2, 0))

        # Button Bar Row 2: Up, Down, Remove, Full Preview
        img_btn_bar2 = tk.Frame(img_frame, bg="#1e293b")
        img_btn_bar2.pack(fill="x", pady=(2, 4))

        btn_up_img = tk.Button(img_btn_bar2, text="▲ Lên", command=self.move_image_up, bg="#334155", fg="#f8fafc", activebackground="#475569", font=("Segoe UI", 8), relief="flat", pady=1, cursor="hand2")
        btn_up_img.pack(side="left", fill="x", expand=True, padx=(0, 2))

        btn_down_img = tk.Button(img_btn_bar2, text="▼ Xuống", command=self.move_image_down, bg="#334155", fg="#f8fafc", activebackground="#475569", font=("Segoe UI", 8), relief="flat", pady=1, cursor="hand2")
        btn_down_img.pack(side="left", fill="x", expand=True, padx=2)

        btn_del_img = tk.Button(img_btn_bar2, text="❌ Xóa", command=self.remove_selected_image, bg="#334155", fg="#f43f5e", activebackground="#475569", font=("Segoe UI", 8), relief="flat", pady=1, cursor="hand2")
        btn_del_img.pack(side="left", fill="x", expand=True, padx=2)

        btn_view_img = tk.Button(img_btn_bar2, text="🔍 Xem Lớn", command=self.preview_selected_image, bg="#0d9488", fg="#ffffff", activebackground="#0f766e", font=("Segoe UI", 8, "bold"), relief="flat", pady=1, cursor="hand2")
        btn_view_img.pack(side="right", fill="x", expand=True, padx=(2, 0))

        # Live Image Thumbnail Preview Box
        self.preview_card = tk.Frame(img_frame, bg="#090d16", highlightthickness=1, highlightbackground="#334155", pady=4, padx=4)
        self.preview_card.pack(fill="x", pady=(2, 0))

        self.preview_canvas = tk.Canvas(self.preview_card, height=110, bg="#090d16", highlightthickness=0, cursor="hand2")
        self.preview_canvas.pack(fill="x")
        self.preview_canvas.bind("<Button-1>", lambda e: self.preview_selected_image())

        self.preview_info_lbl = tk.Label(self.preview_card, text="Chưa chọn ảnh nào", font=("Segoe UI", 8), bg="#090d16", fg="#94a3b8")
        self.preview_info_lbl.pack(fill="x", pady=(2, 0))

        # Right Text Area (Editor)
        editor_frame = ttk.LabelFrame(main_container, text=" Raw ULN Text Input (Paste or Edit) ", padding=10)
        editor_frame.pack(side="right", fill="both", expand=True)

        # ToolBar inside Editor
        btn_bar = tk.Frame(editor_frame, bg="#1e293b", pady=4)
        btn_bar.pack(fill="x", side="top")

        # Pack Compile button FIRST to right to ensure it is ALWAYS visible and never clipped
        btn_compile = tk.Button(
            btn_bar,
            text="🚀 COMPILE TO DOCX",
            command=self.compile_docx,
            bg="#16a34a",
            fg="#ffffff",
            activebackground="#15803d",
            activeforeground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=14,
            pady=4,
            cursor="hand2"
        )
        btn_compile.pack(side="right", padx=(6, 2))

        btn_clear = tk.Button(btn_bar, text="🗑️ Clear", command=self.clear_text, bg="#334155", fg="#f43f5e", activebackground="#475569", activeforeground="#f43f5e", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, cursor="hand2")
        btn_clear.pack(side="right", padx=2)

        btn_search = tk.Button(
            btn_bar,
            text="🔍 Tìm (Ctrl+F)",
            command=self.toggle_search_bar,
            bg="#334155",
            fg="#38bdf8",
            activebackground="#0284c7",
            activeforeground="#ffffff",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=8,
            pady=3,
            cursor="hand2"
        )
        btn_search.pack(side="right", padx=2)

        btn_import_docx = tk.Button(btn_bar, text="📥 Nhập DOCX", command=self.open_docx_picker, bg="#0284c7", fg="#ffffff", activebackground="#0369a1", activeforeground="#ffffff", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, cursor="hand2")
        btn_import_docx.pack(side="left", padx=2)

        btn_embed_file = tk.Button(btn_bar, text="💾 Nhúng Raw", command=self.embed_raw_to_docx_file, bg="#0369a1", fg="#ffffff", activebackground="#0284c7", activeforeground="#ffffff", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, cursor="hand2")
        btn_embed_file.pack(side="left", padx=2)

        btn_import = tk.Button(btn_bar, text="📁 File .txt", command=self.import_file, bg="#334155", fg="#38bdf8", activebackground="#475569", activeforeground="#38bdf8", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, cursor="hand2")
        btn_import.pack(side="left", padx=2)

        btn_sample = tk.Button(btn_bar, text="📄 Mẫu ULN", command=self.load_sample, bg="#475569", fg="#ffffff", activebackground="#334155", activeforeground="#ffffff", font=("Segoe UI", 8, "bold"), relief="flat", padx=8, pady=3, cursor="hand2")
        btn_sample.pack(side="left", padx=2)

        # Heading Quick Action Buttons (Alt+1..6)
        hdr_bar = tk.Frame(btn_bar, bg="#1e293b")
        hdr_bar.pack(side="left", padx=6)
        tk.Label(hdr_bar, text="H:", bg="#1e293b", fg="#94a3b8", font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 2))
        for lvl in range(1, 7):
            h_btn = tk.Button(
                hdr_bar,
                text=f"H{lvl}",
                command=lambda l=lvl: self.apply_heading_shortcut(l),
                bg="#090d16",
                fg="#38bdf8",
                activebackground="#3b82f6",
                activeforeground="#ffffff",
                font=("Segoe UI", 8, "bold"),
                relief="flat",
                padx=4,
                pady=1,
                cursor="hand2"
            )
            h_btn.pack(side="left", padx=1)

        # Search & Replace Panel (Hidden by default, toggled with Ctrl+F / Ctrl+H)
        self.search_matches = []
        self.current_match_idx = -1
        self.search_is_visible = False
        self.replace_is_visible = False

        self.search_frame = tk.Frame(editor_frame, bg="#1e293b", padx=10, pady=6, highlightbackground="#334155", highlightthickness=1)

        # Search Row (Find)
        search_row = tk.Frame(self.search_frame, bg="#1e293b")
        search_row.pack(fill="x", pady=(0, 2))

        tk.Label(search_row, text="🔍 Tìm:", bg="#1e293b", fg="#38bdf8", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.on_search_text_changed())

        self.search_entry = tk.Entry(
            search_row,
            textvariable=self.search_var,
            bg="#090d16",
            fg="#f8fafc",
            insertbackground="#ffffff",
            font=("Consolas", 10),
            relief="flat",
            width=26
        )
        self.search_entry.pack(side="left", padx=(0, 8), ipady=2)

        self.match_count_label = tk.Label(search_row, text="0 kết quả", bg="#1e293b", fg="#94a3b8", font=("Segoe UI", 9), width=12, anchor="w")
        self.match_count_label.pack(side="left", padx=(0, 6))

        btn_prev = tk.Button(
            search_row,
            text="▲ Trước",
            command=self.find_prev,
            bg="#334155",
            fg="#f8fafc",
            activebackground="#475569",
            activeforeground="#ffffff",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=6,
            pady=1,
            cursor="hand2"
        )
        btn_prev.pack(side="left", padx=2)

        btn_next = tk.Button(
            search_row,
            text="▼ Sau",
            command=self.find_next,
            bg="#334155",
            fg="#f8fafc",
            activebackground="#475569",
            activeforeground="#ffffff",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=6,
            pady=1,
            cursor="hand2"
        )
        btn_next.pack(side="left", padx=2)

        self.match_case_var = tk.BooleanVar(value=False)
        chk_case = tk.Checkbutton(
            search_row,
            text="Aa Phân biệt hoa/thường",
            variable=self.match_case_var,
            command=self.on_search_text_changed,
            bg="#1e293b",
            fg="#94a3b8",
            selectcolor="#090d16",
            activebackground="#1e293b",
            activeforeground="#f8fafc",
            font=("Segoe UI", 8)
        )
        chk_case.pack(side="left", padx=8)

        self.btn_toggle_replace = tk.Button(
            search_row,
            text="🔄 Thay thế",
            command=self.toggle_replace_mode,
            bg="#334155",
            fg="#f59e0b",
            activebackground="#d97706",
            activeforeground="#ffffff",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=8,
            pady=1,
            cursor="hand2"
        )
        self.btn_toggle_replace.pack(side="left", padx=4)

        btn_close_search = tk.Button(
            search_row,
            text="✕",
            command=self.hide_search_bar,
            bg="#1e293b",
            fg="#94a3b8",
            activebackground="#1e293b",
            activeforeground="#f43f5e",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            cursor="hand2",
            padx=4
        )
        btn_close_search.pack(side="right", padx=(4, 0))

        # Replace Row
        self.replace_row = tk.Frame(self.search_frame, bg="#1e293b")

        tk.Label(self.replace_row, text="🔄 Thay:", bg="#1e293b", fg="#f59e0b", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 6))

        self.replace_var = tk.StringVar()
        self.replace_entry = tk.Entry(
            self.replace_row,
            textvariable=self.replace_var,
            bg="#090d16",
            fg="#f8fafc",
            insertbackground="#ffffff",
            font=("Consolas", 10),
            relief="flat",
            width=26
        )
        self.replace_entry.pack(side="left", padx=(0, 8), ipady=2)

        btn_replace = tk.Button(
            self.replace_row,
            text="Thay thế (Enter)",
            command=self.replace_current_match,
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=8,
            pady=1,
            cursor="hand2"
        )
        btn_replace.pack(side="left", padx=2)

        btn_replace_all = tk.Button(
            self.replace_row,
            text="Thay tất cả",
            command=self.replace_all_matches,
            bg="#0369a1",
            fg="#ffffff",
            activebackground="#0284c7",
            activeforeground="#ffffff",
            font=("Segoe UI", 8, "bold"),
            relief="flat",
            padx=8,
            pady=1,
            cursor="hand2"
        )
        btn_replace_all.pack(side="left", padx=2)

        # Text Area with Scrollbar
        txt_scroll = ttk.Scrollbar(editor_frame)
        txt_scroll.pack(side="right", fill="y")

        self.text_editor = tk.Text(
            editor_frame,
            wrap="none",
            bg="#090d16",
            fg="#e2e8f0",
            insertbackground="#ffffff",
            font=("Consolas", 11),
            yscrollcommand=txt_scroll.set,
            padx=10,
            pady=10
        )
        self.text_editor.pack(fill="both", expand=True)
        txt_scroll.config(command=self.text_editor.yview)

        # Configure search highlighting tags
        self.text_editor.tag_config("search_match", background="#334155", foreground="#38bdf8")
        self.text_editor.tag_config("search_current", background="#f59e0b", foreground="#090d16")

        # Bind Keybindings
        self.root.bind("<Control-f>", lambda e: self.show_search_bar(with_replace=False))
        self.root.bind("<Control-F>", lambda e: self.show_search_bar(with_replace=False))
        self.root.bind("<Control-h>", lambda e: self.show_search_bar(with_replace=True))
        self.root.bind("<Control-H>", lambda e: self.show_search_bar(with_replace=True))
        self.root.bind("<F3>", lambda e: self.find_next())
        self.root.bind("<Shift-F3>", lambda e: self.find_prev())

        self.text_editor.bind("<Control-f>", lambda e: self.show_search_bar(with_replace=False))
        self.text_editor.bind("<Control-F>", lambda e: self.show_search_bar(with_replace=False))
        self.text_editor.bind("<Control-h>", lambda e: self.show_search_bar(with_replace=True))
        self.text_editor.bind("<Control-H>", lambda e: self.show_search_bar(with_replace=True))

        self.search_entry.bind("<Return>", self.find_next)
        self.search_entry.bind("<Shift-Return>", self.find_prev)
        self.search_entry.bind("<Escape>", lambda e: self.hide_search_bar())
        self.replace_entry.bind("<Return>", lambda e: self.replace_current_match())
        self.replace_entry.bind("<Escape>", lambda e: self.hide_search_bar())

        # Bind Alt+1 to Alt+6 heading shortcuts
        for lvl in range(1, 7):
            self.root.bind(f"<Alt-Key-{lvl}>", lambda e, l=lvl: self.apply_heading_shortcut(l, e))
            self.root.bind(f"<Alt-KP_{lvl}>", lambda e, l=lvl: self.apply_heading_shortcut(l, e))
            self.text_editor.bind(f"<Alt-Key-{lvl}>", lambda e, l=lvl: self.apply_heading_shortcut(l, e))
            self.text_editor.bind(f"<Alt-KP_{lvl}>", lambda e, l=lvl: self.apply_heading_shortcut(l, e))

        # Preload default sample text
        self.load_sample()

        # Check for updates automatically in background after startup
        self.root.after(1500, self.auto_check_updates_background)

    # ── SEARCH & REPLACE SYSTEM ────────────────────────────────────────
    def show_search_bar(self, with_replace: bool = False):
        """Displays search bar, populating with any currently selected text."""
        try:
            sel_text = self.text_editor.get(tk.SEL_FIRST, tk.SEL_LAST).strip()
            if sel_text and '\n' not in sel_text:
                self.search_var.set(sel_text)
        except Exception:
            pass

        self.search_frame.pack(fill="x", side="top", before=self.text_editor, pady=(0, 4))
        self.search_is_visible = True

        if with_replace:
            self.replace_row.pack(fill="x", pady=(4, 0))
            self.replace_is_visible = True
            self.replace_entry.focus_set()
            self.replace_entry.select_range(0, tk.END)
        else:
            if hasattr(self, 'replace_is_visible') and not self.replace_is_visible:
                self.replace_row.pack_forget()
            self.search_entry.focus_set()
            self.search_entry.select_range(0, tk.END)

        self.on_search_text_changed()

    def hide_search_bar(self):
        """Hides search bar and clears all highlight tags."""
        self.search_frame.pack_forget()
        self.search_is_visible = False
        self.clear_search_highlights()
        self.text_editor.focus_set()

    def toggle_search_bar(self):
        if getattr(self, 'search_is_visible', False):
            self.hide_search_bar()
        else:
            self.show_search_bar(with_replace=False)

    def toggle_replace_mode(self):
        if getattr(self, 'replace_is_visible', False):
            self.replace_row.pack_forget()
            self.replace_is_visible = False
            self.search_entry.focus_set()
        else:
            self.replace_row.pack(fill="x", pady=(4, 0))
            self.replace_is_visible = True
            self.replace_entry.focus_set()
            self.replace_entry.select_range(0, tk.END)

    def clear_search_highlights(self):
        self.text_editor.tag_remove("search_match", "1.0", tk.END)
        self.text_editor.tag_remove("search_current", "1.0", tk.END)
        self.search_matches = []
        self.current_match_idx = -1

    def on_search_text_changed(self):
        """Finds all occurrences in text_editor and highlights them."""
        self.clear_search_highlights()
        query = self.search_var.get()
        if not query:
            self.match_count_label.config(text="0 kết quả", fg="#94a3b8")
            return

        match_case = self.match_case_var.get()
        start = "1.0"
        while True:
            pos = self.text_editor.search(query, start, stopindex=tk.END, nocase=not match_case)
            if not pos:
                break
            end = f"{pos}+{len(query)}c"
            self.search_matches.append((pos, end))
            self.text_editor.tag_add("search_match", pos, end)
            start = end

        total = len(self.search_matches)
        if total > 0:
            self.current_match_idx = 0
            self._highlight_current_match()
        else:
            self.match_count_label.config(text="Không tìm thấy", fg="#f43f5e")

    def _highlight_current_match(self):
        self.text_editor.tag_remove("search_current", "1.0", tk.END)
        if 0 <= self.current_match_idx < len(self.search_matches):
            pos, end = self.search_matches[self.current_match_idx]
            self.text_editor.tag_add("search_current", pos, end)
            self.text_editor.see(pos)
            self.match_count_label.config(text=f"{self.current_match_idx + 1}/{len(self.search_matches)} kết quả", fg="#38bdf8")

    def find_next(self, event=None):
        if not self.search_matches:
            return "break"
        self.current_match_idx = (self.current_match_idx + 1) % len(self.search_matches)
        self._highlight_current_match()
        return "break"

    def find_prev(self, event=None):
        if not self.search_matches:
            return "break"
        self.current_match_idx = (self.current_match_idx - 1) % len(self.search_matches)
        self._highlight_current_match()
        return "break"

    def replace_current_match(self, event=None):
        if not self.search_matches or self.current_match_idx < 0 or self.current_match_idx >= len(self.search_matches):
            return "break"
        pos, end = self.search_matches[self.current_match_idx]
        replace_str = self.replace_var.get()
        self.text_editor.delete(pos, end)
        self.text_editor.insert(pos, replace_str)
        curr_idx = self.current_match_idx
        self.on_search_text_changed()
        if self.search_matches:
            self.current_match_idx = min(curr_idx, len(self.search_matches) - 1)
            self._highlight_current_match()
        return "break"

    def replace_all_matches(self):
        if not self.search_matches:
            return
        replace_str = self.replace_var.get()
        for pos, end in reversed(self.search_matches):
            self.text_editor.delete(pos, end)
            self.text_editor.insert(pos, replace_str)
        self.on_search_text_changed()

    # ── PROMPT MANAGEMENT ──────────────────────────────────────────────
    def get_prompt_storage_path(self) -> str:
        return get_prompt_storage_path()

    def get_default_prompt_path(self) -> str:
        return get_default_prompt_path()

    def load_prompt_text(self) -> str:
        return load_prompt_text()

    def save_prompt_text(self, text: str) -> bool:
        return save_prompt_text(text)

    def reset_prompt_text(self) -> str:
        return reset_prompt_text()

    def open_prompt_editor(self):
        open_prompt_editor_dialog(self.root)

    # ── UPDATE SYSTEM ──────────────────────────────────────────────────
    def auto_check_updates_background(self):
        """Silently checks for updates in background on startup."""
        thread = threading.Thread(target=self._async_check_update, args=(True,), daemon=True)
        thread.start()

    def manual_check_updates(self):
        """User-triggered manual update check."""
        self.btn_update.config(text="⏳ Đang kiểm tra...", state="disabled")
        thread = threading.Thread(target=self._async_check_update, args=(False,), daemon=True)
        thread.start()

    def _async_check_update(self, silent: bool):
        try:
            info = updater.check_for_updates()
            def _safe_res():
                try:
                    if self.root.winfo_exists():
                        self._handle_update_result(info, silent)
                except Exception:
                    pass
            self.root.after(0, _safe_res)
        except Exception as e:
            if not silent:
                def _safe_err():
                    try:
                        if self.root.winfo_exists():
                            messagebox.showerror("Lỗi Cập Nhật", f"Không thể kết nối đến máy chủ cập nhật:\n{e}")
                    except Exception:
                        pass
                self.root.after(0, _safe_err)
        finally:
            def _safe_btn():
                try:
                    if self.root.winfo_exists():
                        self.btn_update.config(text="🔄 Kiểm tra Cập nhật", state="normal")
                except Exception:
                    pass
            try:
                self.root.after(0, _safe_btn)
            except Exception:
                pass

    def _handle_update_result(self, info: dict, silent: bool):
        if info.get("has_update"):
            self.show_update_modal(info)
        elif not silent:
            if info.get("success"):
                messagebox.showinfo("Cập nhật Phần mềm", f"Bạn đang sử dụng phiên bản mới nhất (v{self.version}).")
            else:
                err = info.get("error", "Lỗi không xác định.")
                messagebox.showwarning("Kiểm tra Cập nhật", f"Không thể lấy thông tin cập nhật:\n{err}")

    def show_update_modal(self, info: dict):
        show_update_modal_dialog(self.root, info)

    # ── HEADING SHORTCUTS ──────────────────────────────────────────────
    def apply_heading_shortcut(self, level: int, event=None):
        tag_str = f"[H{level}]"
        try:
            sel_ranges = self.text_editor.tag_ranges("sel")
            if sel_ranges:
                start_line = int(self.text_editor.index("sel.first").split('.')[0])
                end_line = int(self.text_editor.index("sel.last").split('.')[0])
                if self.text_editor.index("sel.last").endswith(".0") and end_line > start_line:
                    end_line -= 1
            else:
                insert_pos = self.text_editor.index(tk.INSERT)
                start_line = int(insert_pos.split('.')[0])
                end_line = start_line

            for line_num in range(start_line, end_line + 1):
                line_start = f"{line_num}.0"
                line_end = f"{line_num}.end"
                line_text = self.text_editor.get(line_start, line_end)
                if not line_text.strip():
                    continue

                # Strip existing block tag if present
                clean_text = re.sub(r'^\s*\[(?:H[1-6]|P[0-2]|INS|QUOTE|BOX|TABLE(?::\s*borderless)?|TAB2|PIC)\]\s*', '', line_text, flags=re.IGNORECASE)
                new_line = f"{tag_str} {clean_text}"
                
                self.text_editor.delete(line_start, line_end)
                self.text_editor.insert(line_start, new_line)

            return "break"
        except Exception as e:
            print(f"Error applying heading shortcut: {e}")
            return "break"

    def add_images(self):
        files = filedialog.askopenfilenames(
            title="Select Images for [PIC] Tags (Auto-sorted naturally by name)",
            filetypes=[
                ("All Supported Images", "*.png;*.jpg;*.jpeg;*.avif;*.avifs;*.webp;*.heic;*.heif;*.jfif;*.bmp;*.dib;*.gif;*.tiff;*.tif;*.ico;*.svg;*.wmf;*.emf"),
                ("AVIF Images (*.avif)", "*.avif;*.avifs"),
                ("WebP Images (*.webp)", "*.webp"),
                ("HEIC/HEIF Images (*.heic, *.heif)", "*.heic;*.heif"),
                ("Standard Images (*.png, *.jpg, *.jpeg, *.jfif)", "*.png;*.jpg;*.jpeg;*.jfif"),
                ("All Files (*.*)", "*.*")
            ]
        )
        if files:
            for f in files:
                abs_f = os.path.abspath(f)
                if abs_f not in self.user_image_paths:
                    self.user_image_paths.append(abs_f)
            # Automatically sort all image paths naturally by filename (e.g. 1, 2, 3... 10)
            self.user_image_paths.sort(key=lambda p: natural_sort_key(os.path.basename(p)))
            self.update_image_listbox(selected_index=0)

    def sort_images(self):
        """Manually sorts the image queue naturally by filename."""
        self.user_image_paths.sort(key=lambda p: natural_sort_key(os.path.basename(p)))
        self.update_image_listbox(selected_index=0)

    def clear_images(self):
        self.user_image_paths.clear()
        self.update_image_listbox()

    def on_image_selected(self, event=None):
        """Updates live thumbnail preview and metadata when user clicks or navigates listbox."""
        sel = self.img_listbox.curselection()
        if not sel or not self.user_image_paths:
            self.preview_canvas.delete("all")
            self.preview_info_lbl.config(text="Chưa chọn ảnh nào")
            self._current_preview_photo = None
            return

        idx = sel[0]
        if idx < 0 or idx >= len(self.user_image_paths):
            return

        img_path = self.user_image_paths[idx]
        fname = os.path.basename(img_path)

        self.preview_canvas.update_idletasks()
        cw = max(120, self.preview_canvas.winfo_width() - 8)
        ch = max(80, self.preview_canvas.winfo_height() - 8)

        res = load_and_scale_image(img_path, max_w=cw, max_h=ch)
        self.preview_canvas.delete("all")
        if res:
            photo, orig_w, orig_h, fmt, size_str = res
            self._current_preview_photo = photo
            self.preview_canvas.create_image(self.preview_canvas.winfo_width() // 2, self.preview_canvas.winfo_height() // 2, anchor="center", image=photo)
            self.preview_info_lbl.config(text=f"[PIC #{idx + 1}] {fname}\n{orig_w}×{orig_h} px | {size_str} ({fmt})")
        else:
            self._current_preview_photo = None
            self.preview_canvas.create_text(self.preview_canvas.winfo_width() // 2, self.preview_canvas.winfo_height() // 2, text=f"⚠️ {fname}", fill="#f43f5e", font=("Segoe UI", 8))
            self.preview_info_lbl.config(text=f"[PIC #{idx + 1}] Không thể mở ảnh")

    def preview_selected_image(self):
        """Opens full-size interactive preview modal."""
        sel = self.img_listbox.curselection()
        idx = sel[0] if sel else 0
        open_image_preview_dialog(self, initial_index=idx)

    def move_image_up(self):
        sel = self.img_listbox.curselection()
        if not sel or sel[0] <= 0:
            return
        idx = sel[0]
        self.user_image_paths[idx - 1], self.user_image_paths[idx] = self.user_image_paths[idx], self.user_image_paths[idx - 1]
        self.update_image_listbox(selected_index=idx - 1)

    def move_image_down(self):
        sel = self.img_listbox.curselection()
        if not sel or sel[0] >= len(self.user_image_paths) - 1:
            return
        idx = sel[0]
        self.user_image_paths[idx + 1], self.user_image_paths[idx] = self.user_image_paths[idx], self.user_image_paths[idx + 1]
        self.update_image_listbox(selected_index=idx + 1)

    def remove_selected_image(self):
        sel = self.img_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.user_image_paths.pop(idx)
        new_idx = min(idx, len(self.user_image_paths) - 1) if self.user_image_paths else None
        self.update_image_listbox(selected_index=new_idx)

    def update_image_listbox(self, selected_index: Optional[int] = None):
        self.img_listbox.delete(0, tk.END)
        for idx, p in enumerate(self.user_image_paths, 1):
            self.img_listbox.insert(tk.END, f"{idx}. {os.path.basename(p)}")

        if self.user_image_paths:
            if selected_index is not None and 0 <= selected_index < len(self.user_image_paths):
                self.img_listbox.selection_set(selected_index)
                self.img_listbox.see(selected_index)
            elif not self.img_listbox.curselection():
                self.img_listbox.selection_set(0)
            self.on_image_selected()
        else:
            self.on_image_selected()

    def load_sample(self):
        # Look in bundle dir first (PyInstaller), then current dir
        sample_path = os.path.join(updater.get_bundle_dir(), "uln_test.txt")
        if not os.path.exists(sample_path):
            sample_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uln_test.txt")
        
        if os.path.exists(sample_path):
            try:
                with open(sample_path, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception:
                content = '[H1] SECTION A: PHONETICS\n[P0] 1. Choose the word whose underlined part is pronounced differently:\n[P1] A. [c]{u}at \t B. [c]{u}ity \t C. [c]{u}ar \t D. [c]{u}up'
        else:
            content = '[H1] SECTION A: PHONETICS\n[P0] 1. Choose the word whose underlined part is pronounced differently:\n[P1] A. [c]{u}at \t B. [c]{u}ity \t C. [c]{u}ar \t D. [c]{u}up'
        
        self.text_editor.delete("1.0", tk.END)
        self.text_editor.insert("1.0", content)

    def open_docx_picker(self):
        """
        Opens a modern dark modal dialog to browse, scan, filter, and preview .docx files
        that contain embedded raw ULN data, allowing one-click re-import into the editor.
        """
        open_docx_picker_dialog(self)

    def embed_raw_to_docx_file(self):
        """Allows user to select any existing DOCX file and inject the current editor's ULN code into it."""
        uln_text = self.text_editor.get("1.0", tk.END).strip()
        if not uln_text:
            messagebox.showwarning("Thông báo", "Trình soạn thảo hiện tại đang trống. Vui lòng nhập mã ULN trước khi nhúng.")
            return

        fpath = filedialog.askopenfilename(
            title="Chọn file DOCX để nhúng mã nguồn ULN",
            initialdir=self.last_output_dir,
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if fpath and os.path.exists(fpath):
            from uln_compiler import embed_raw_uln_docx
            ok = embed_raw_uln_docx(fpath, uln_text)
            if ok:
                messagebox.showinfo("Thành công", f"Đã chèn/nhúng thành công mã Raw ULN vào file:\n{os.path.basename(fpath)}")
            else:
                messagebox.showerror("Lỗi", "Không thể chèn mã ULN vào file Word (vui lòng đảm bảo file đã được đóng trong Word trước khi chèn).")



    def import_file(self):
        file_path = filedialog.askopenfilename(
            title="Select ULN Text File",
            filetypes=[("Text Files", "*.txt"), ("ULN Files", "*.uln"), ("All Files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.text_editor.delete("1.0", tk.END)
                self.text_editor.insert("1.0", content)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file: {e}")

    def clear_text(self):
        self.text_editor.delete("1.0", tk.END)

    def compile_docx(self):
        uln_text = self.text_editor.get("1.0", tk.END).strip()
        if not uln_text:
            messagebox.showwarning("Warning", "Please enter or import ULN text before compiling.")
            return

        ts_filename = f"uln_document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

        out_path = filedialog.asksaveasfilename(
            title="Save Formatted DOCX File",
            initialdir=self.last_output_dir,
            initialfile=ts_filename,
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )

        # If user cancelled file save dialog, stop gracefully
        if not out_path:
            return

        # Save last chosen folder for next time
        self.last_output_dir = os.path.dirname(os.path.abspath(out_path))

        # Auto-minimize GUI window so it does not block the live Microsoft Word window
        try:
            self.root.iconify()
            self.root.update()
        except Exception:
            pass

        settings = {
            "font_name": self.font_var.get(),
            "font_size": self.size_var.get(),
            "margin_top": self.m_top_var.get(),
            "margin_bottom": self.m_bottom_var.get(),
            "margin_left": self.m_left_var.get(),
            "margin_right": self.m_right_var.get(),
            "enable_page_numbers": self.pg_num_var.get(),
            "question_prefix": self.q_prefix_var.get(),
            "question_delimiter": self.q_delim_var.get(),
            "question_color": self.q_color_var.get(),
            "opt_color": self.opt_color_var.get(),
            "instruction_color": self.ins_color_var.get(),
            "ins_color": self.ins_color_var.get(),
            "user_images": list(self.user_image_paths),
        }

        try:
            # Dynamic Hot-Reload of format engine modules when running in dev mode
            if not getattr(sys, 'frozen', False):
                try:
                    import importlib
                    import renderer_utils
                    import renderer_blocks
                    import uln_parser
                    import uln_renderer
                    import uln_compiler
                    importlib.reload(renderer_utils)
                    importlib.reload(renderer_blocks)
                    importlib.reload(uln_parser)
                    importlib.reload(uln_renderer)
                    importlib.reload(uln_compiler)
                except Exception:
                    pass

            from uln_compiler import ULNCompiler
            compiler = ULNCompiler(settings)
            keep_open_val = self.open_word_var.get()
            bg_mode_val = self.bg_mode_var.get()
            compiled_file = compiler.compile(uln_text, out_path, keep_open=keep_open_val, background_mode=bg_mode_val)

            try:
                self.root.deiconify()
            except Exception:
                pass

            if not keep_open_val:
                msg = f"Successfully generated DOCX file:\n{compiled_file}"
                messagebox.showinfo("Success", msg)

        except KeyboardInterrupt:
            try:
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
            except Exception:
                pass
            messagebox.showinfo("Đã tạm dừng", "Đã tạm dừng quá trình biên dịch (Phím ESC).\nFile Word hiện tại vẫn được giữ nguyên để bạn xem và chỉnh sửa tiếp.")

        except Exception as e:
            try:
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
            except Exception:
                pass
            messagebox.showerror("Compilation Error", f"Failed to generate Word document:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ULNFormatterApp(root)
    root.mainloop()
