import os
import sys
import re
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from uln_compiler import ULNCompiler, extract_raw_uln, has_embedded_uln, scan_folder_for_uln_docx
import updater

class ULNFormatterApp:
    def __init__(self, root):
        self.root = root
        self.version = updater.get_current_version()
        self.root.title(f"Universal Layout Notation (ULN) → DOCX Formatter v{self.version}")
        self.root.geometry("1220x840")
        self.root.configure(bg="#0f172a")

        self.last_output_dir = os.path.expanduser("~\\Documents")
        if not os.path.exists(self.last_output_dir):
            self.last_output_dir = os.getcwd()

        # Configure dark theme styling
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background="#0f172a", foreground="#f8fafc", font=("Segoe UI", 10))
        style.configure("TLabel", background="#0f172a", foreground="#f8fafc")
        style.configure("TFrame", background="#0f172a")
        style.configure("TLabelframe", background="#1e293b", foreground="#38bdf8", borderwidth=1)
        style.configure("TLabelframe.Label", background="#1e293b", foreground="#38bdf8", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), background="#3b82f6", foreground="#ffffff", borderwidth=0, padding=8)
        style.map("TButton", background=[("active", "#2563eb")])

        # Configure Combobox and Spinbox dark inputs with crisp white readable text
        style.configure("TCombobox", fieldbackground="#090d16", background="#1e293b", foreground="#f8fafc", selectbackground="#3b82f6", selectforeground="#ffffff", arrowcolor="#38bdf8", insertcolor="#ffffff")
        style.map("TCombobox", fieldbackground=[("readonly", "#090d16"), ("active", "#1e293b"), ("focus", "#090d16"), ("!disabled", "#090d16")], foreground=[("readonly", "#f8fafc"), ("active", "#ffffff"), ("!disabled", "#f8fafc")])
        style.configure("TSpinbox", fieldbackground="#090d16", background="#1e293b", foreground="#f8fafc", selectbackground="#3b82f6", selectforeground="#ffffff", arrowcolor="#38bdf8", insertcolor="#ffffff")
        style.map("TSpinbox", fieldbackground=[("readonly", "#090d16"), ("active", "#1e293b"), ("focus", "#090d16"), ("!disabled", "#090d16")], foreground=[("readonly", "#f8fafc"), ("active", "#ffffff"), ("!disabled", "#f8fafc")])

        self.root.option_add("*TCombobox*Listbox.background", "#090d16")
        self.root.option_add("*TCombobox*Listbox.foreground", "#f8fafc")
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#3b82f6")
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        self.root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 9))

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

        # Image Queue (Order for [PIC] tags)
        self.user_image_paths = []
        img_frame = ttk.LabelFrame(sidebar, text=" 🖼️ Image Queue ([PIC] Order) ", padding=6)
        img_frame.pack(fill="x", pady=(2, 6))

        self.img_listbox = tk.Listbox(img_frame, height=3, bg="#090d16", fg="#e2e8f0", font=("Segoe UI", 9), selectbackground="#3b82f6")
        self.img_listbox.pack(fill="x", pady=(0, 4))

        img_btn_bar = tk.Frame(img_frame, bg="#1e293b")
        img_btn_bar.pack(fill="x")

        btn_add_img = tk.Button(img_btn_bar, text="➕ Add Images...", command=self.add_images, bg="#3b82f6", fg="#ffffff", font=("Segoe UI", 8, "bold"), relief="flat", pady=2)
        btn_add_img.pack(side="left", fill="x", expand=True, padx=(0, 2))

        btn_clear_img = tk.Button(img_btn_bar, text="🗑️ Clear", command=self.clear_images, bg="#64748b", fg="#ffffff", font=("Segoe UI", 8), relief="flat", pady=2)
        btn_clear_img.pack(side="right", padx=(2, 0))

        # Action Buttons on Sidebar
        btn_docx_side = tk.Button(sidebar, text="📥 Nhập từ file DOCX...", command=self.open_docx_picker, bg="#0284c7", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", pady=4)
        btn_docx_side.pack(fill="x", pady=(6, 2))

        btn_prompt_side = tk.Button(sidebar, text="📜 Xem / Sửa AI Prompt...", command=self.open_prompt_editor, bg="#0369a1", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", pady=4)
        btn_prompt_side.pack(fill="x", pady=2)

        btn_sample = tk.Button(sidebar, text="Load Sample ULN", command=self.load_sample, bg="#475569", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", pady=4)
        btn_sample.pack(fill="x", pady=2)

        # Right Text Area (Editor)
        editor_frame = ttk.LabelFrame(main_container, text=" Raw ULN Text Input (Paste or Edit) ", padding=10)
        editor_frame.pack(side="right", fill="both", expand=True)

        # ToolBar inside Editor
        btn_bar = tk.Frame(editor_frame, bg="#1e293b", pady=4)
        btn_bar.pack(fill="x", side="top")

        btn_import_docx = tk.Button(btn_bar, text="📥 Nhập từ file DOCX", command=self.open_docx_picker, bg="#0284c7", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4)
        btn_import_docx.pack(side="left", padx=5)

        btn_import = tk.Button(btn_bar, text="📁 Import .txt", command=self.import_file, bg="#334155", fg="#38bdf8", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4)
        btn_import.pack(side="left", padx=5)

        btn_clear = tk.Button(btn_bar, text="🗑️ Clear Text", command=self.clear_text, bg="#334155", fg="#f43f5e", font=("Segoe UI", 9), relief="flat", padx=10, pady=4)
        btn_clear.pack(side="left", padx=5)

        # Heading Quick Action Buttons (Alt+1..6)
        hdr_bar = tk.Frame(btn_bar, bg="#1e293b")
        hdr_bar.pack(side="left", padx=10)
        tk.Label(hdr_bar, text="Headings:", bg="#1e293b", fg="#94a3b8", font=("Segoe UI", 9)).pack(side="left", padx=(0, 2))
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
                pady=2
            )
            h_btn.pack(side="left", padx=1)

        btn_compile = tk.Button(btn_bar, text="🚀 COMPILE TO DOCX", command=self.compile_docx, bg="#16a34a", fg="#ffffff", font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=4)
        btn_compile.pack(side="right", padx=5)

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

    # ── PROMPT MANAGEMENT ──────────────────────────────────────────────
    def get_prompt_storage_path(self) -> str:
        """Returns path where user-edited prompt is saved and loaded."""
        if getattr(sys, 'frozen', False):
            app_data_dir = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "DocxConverter")
            os.makedirs(app_data_dir, exist_ok=True)
            return os.path.join(app_data_dir, "prompt.txt")
        else:
            return os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompt.txt")

    def get_default_prompt_path(self) -> str:
        """Returns bundled default prompt path."""
        bundle_p = os.path.join(updater.get_bundle_dir(), "prompt.txt")
        if os.path.exists(bundle_p):
            return bundle_p
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompt.txt")

    def load_prompt_text(self) -> str:
        """Loads prompt from storage path or fallback to default bundle."""
        storage_p = self.get_prompt_storage_path()
        if os.path.exists(storage_p):
            try:
                with open(storage_p, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception:
                pass
        
        default_p = self.get_default_prompt_path()
        if os.path.exists(default_p):
            try:
                with open(default_p, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception:
                pass
        return "You are an expert OCR and Universal Layout Notation (ULN) extraction engine."

    def save_prompt_text(self, text: str) -> bool:
        """Saves prompt text to storage path (and project directory if dev mode)."""
        try:
            storage_p = self.get_prompt_storage_path()
            parent = os.path.dirname(storage_p)
            if parent:
                os.makedirs(parent, exist_ok=True)
            with open(storage_p, "w", encoding="utf-8") as f:
                f.write(text)
            
            # If in dev mode, also ensure local prompt.txt is synced
            if not getattr(sys, 'frozen', False):
                local_p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompt.txt")
                if local_p != storage_p:
                    with open(local_p, "w", encoding="utf-8") as f:
                        f.write(text)
            return True
        except Exception as e:
            print(f"Error saving prompt: {e}")
            return False

    def reset_prompt_text(self) -> str:
        """Resets prompt text from default bundled prompt and saves to storage path."""
        default_p = self.get_default_prompt_path()
        content = ""
        if os.path.exists(default_p):
            try:
                with open(default_p, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception:
                pass
        if content:
            self.save_prompt_text(content)
        return content

    def open_prompt_editor(self):
        """Opens a modern dark modal to view, edit, copy, and manage AI prompt."""
        win = tk.Toplevel(self.root)
        win.title("📜 Quản Lý & Chỉnh Sửa AI Prompt (ULN Extraction)")
        win.geometry("960x720")
        win.configure(bg="#0f172a")
        win.transient(self.root)
        win.grab_set()

        try:
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 480
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 360
            win.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        # Top Header Bar
        top_bar = tk.Frame(win, bg="#1e293b", padx=20, pady=12)
        top_bar.pack(fill="x")

        tk.Label(
            top_bar,
            text="📜 AI System Prompt & Quy Tắc Trích Xuất ULN",
            font=("Segoe UI", 14, "bold"),
            bg="#1e293b",
            fg="#38bdf8"
        ).pack(anchor="w")

        tk.Label(
            top_bar,
            text="Prompt này được dùng để nạp vào ChatGPT, Claude, Gemini hoặc DeepSeek để nhận diện & chuyển đổi đề thi thành ULN.",
            font=("Segoe UI", 9),
            bg="#1e293b",
            fg="#94a3b8"
        ).pack(anchor="w", pady=(2, 0))

        # Toolbar
        tool_frame = tk.Frame(win, bg="#111827", padx=16, pady=8)
        tool_frame.pack(fill="x")

        def copy_all():
            text = p_text.get("1.0", tk.END).strip()
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            status_lbl.config(text="✓ Đã sao chép toàn bộ Prompt vào Clipboard!", fg="#4ade80")
            self.root.after(3000, lambda: status_lbl.config(text=f"File: {self.get_prompt_storage_path()}", fg="#94a3b8"))

        def save_changes():
            text = p_text.get("1.0", tk.END).rstrip() + "\n"
            if self.save_prompt_text(text):
                status_lbl.config(text="✓ Đã lưu thay đổi Prompt thành công!", fg="#4ade80")
                messagebox.showinfo("Thành công", "Đã lưu nội dung Prompt thành công!")
            else:
                status_lbl.config(text="⚠ Lỗi khi lưu file prompt!", fg="#f43f5e")
                messagebox.showerror("Lỗi", "Không thể lưu file prompt.")
            self.root.after(3000, lambda: status_lbl.config(text=f"File: {self.get_prompt_storage_path()}", fg="#94a3b8"))

        def reset_default():
            if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn khôi phục lại Prompt mặc định ban đầu không?"):
                default_content = self.reset_prompt_text()
                p_text.delete("1.0", tk.END)
                p_text.insert("1.0", default_content)
                status_lbl.config(text="✓ Đã khôi phục Prompt về mặc định!", fg="#38bdf8")
                self.root.after(3000, lambda: status_lbl.config(text=f"File: {self.get_prompt_storage_path()}", fg="#94a3b8"))

        def open_folder():
            storage_path = self.get_prompt_storage_path()
            folder = os.path.dirname(storage_path)
            if not os.path.exists(folder):
                os.makedirs(folder, exist_ok=True)
            if not os.path.exists(storage_path):
                self.save_prompt_text(p_text.get("1.0", tk.END))
            try:
                subprocess.Popen(f'explorer /select,"{storage_path}"')
            except Exception:
                try:
                    os.startfile(folder)
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể mở thư mục: {e}")

        btn_copy = tk.Button(
            tool_frame,
            text="📋 Sao Chép Prompt (Copy)",
            command=copy_all,
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=12,
            pady=5,
            cursor="hand2"
        )
        btn_copy.pack(side="left", padx=(0, 6))

        btn_save = tk.Button(
            tool_frame,
            text="💾 Lưu Thay Đổi (Save)",
            command=save_changes,
            bg="#16a34a",
            fg="#ffffff",
            activebackground="#15803d",
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=12,
            pady=5,
            cursor="hand2"
        )
        btn_save.pack(side="left", padx=6)

        btn_reset = tk.Button(
            tool_frame,
            text="🔄 Khôi Phục Mặc Định",
            command=reset_default,
            bg="#475569",
            fg="#ffffff",
            activebackground="#334155",
            activeforeground="#ffffff",
            font=("Segoe UI", 9),
            relief="flat",
            padx=10,
            pady=5,
            cursor="hand2"
        )
        btn_reset.pack(side="left", padx=6)

        btn_open_file = tk.Button(
            tool_frame,
            text="📁 Mở Thư Mục Chứa File",
            command=open_folder,
            bg="#334155",
            fg="#e2e8f0",
            activebackground="#1e293b",
            activeforeground="#ffffff",
            font=("Segoe UI", 9),
            relief="flat",
            padx=10,
            pady=5,
            cursor="hand2"
        )
        btn_open_file.pack(side="left", padx=6)

        # Editor Area
        editor_box = tk.Frame(win, bg="#0f172a", padx=16, pady=10)
        editor_box.pack(fill="both", expand=True)

        scroll_y = ttk.Scrollbar(editor_box, orient="vertical")
        scroll_y.pack(side="right", fill="y")

        p_text = tk.Text(
            editor_box,
            wrap="word",
            bg="#090d16",
            fg="#f8fafc",
            insertbackground="#ffffff",
            font=("Consolas", 10),
            yscrollcommand=scroll_y.set,
            padx=12,
            pady=12,
            relief="flat"
        )
        p_text.pack(fill="both", expand=True)
        scroll_y.config(command=p_text.yview)

        # Load current content
        current_content = self.load_prompt_text()
        p_text.insert("1.0", current_content)

        # Status Footer
        footer = tk.Frame(win, bg="#1e293b", padx=16, pady=8)
        footer.pack(fill="x", side="bottom")

        status_lbl = tk.Label(
            footer,
            text=f"File: {self.get_prompt_storage_path()}",
            font=("Segoe UI", 9),
            bg="#1e293b",
            fg="#94a3b8"
        )
        status_lbl.pack(side="left")

        btn_close = tk.Button(
            footer,
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
        btn_close.pack(side="right")

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
            self.root.after(0, lambda: self._handle_update_result(info, silent))
        except Exception as e:
            if not silent:
                self.root.after(0, lambda: messagebox.showerror("Lỗi Cập Nhật", f"Không thể kết nối đến máy chủ cập nhật:\n{e}"))
        finally:
            self.root.after(0, lambda: self.btn_update.config(text="🔄 Kiểm tra Cập nhật", state="normal"))

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
        """Modern dark-themed update popup dialog."""
        win = tk.Toplevel(self.root)
        win.title("Có Bản Cập Nhật Mới")
        win.geometry("560x420")
        win.configure(bg="#0f172a")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        # Center modal over root window
        try:
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 280
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 210
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
                    self.root.after(0, lambda: [
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
                    self.root.after(0, lambda: [
                        progress_bar.pack_forget(),
                        status_lbl.config(text=msg, fg="#f43f5e"),
                        btn_update_now.config(state="normal", text="Mở trang tải trên Web"),
                        btn_update_now.config(command=lambda: updater.webbrowser.open(info.get("release_url", ""))),
                        btn_cancel.config(state="normal")
                    ])
                else:
                    self.root.after(0, lambda: [
                        status_lbl.config(text=msg, fg="#4ade80"),
                        btn_cancel.config(state="normal", text="Đóng")
                    ])

            threading.Thread(target=run_dl, daemon=True).start()

        btn_cancel = tk.Button(btn_frame, text="Để Sau", command=win.destroy, bg="#475569", fg="#ffffff", font=("Segoe UI", 9), relief="flat", padx=14, pady=6)
        btn_cancel.pack(side="right", padx=(6, 0))

        btn_update_now = tk.Button(btn_frame, text="🚀 Cập Nhật Ngay", command=start_update, bg="#16a34a", fg="#ffffff", font=("Segoe UI", 9, "bold"), relief="flat", padx=16, pady=6)
        btn_update_now.pack(side="right")

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
            title="Select Images for [PIC] Tags (In Order)",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"), ("All Files", "*.*")]
        )
        if files:
            for f in files:
                abs_f = os.path.abspath(f)
                if abs_f not in self.user_image_paths:
                    self.user_image_paths.append(abs_f)
            self.update_image_listbox()

    def clear_images(self):
        self.user_image_paths.clear()
        self.update_image_listbox()

    def update_image_listbox(self):
        self.img_listbox.delete(0, tk.END)
        for idx, p in enumerate(self.user_image_paths, 1):
            self.img_listbox.insert(tk.END, f"{idx}. {os.path.basename(p)}")

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
        win = tk.Toplevel(self.root)
        win.title("📂 Trình Duyệt & Nhập File DOCX Có Nhúng ULN")
        win.geometry("1060x680")
        win.configure(bg="#0f172a")
        win.transient(self.root)
        win.grab_set()

        try:
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 530
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 340
            win.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        current_folder = tk.StringVar(value=self.last_output_dir if os.path.exists(self.last_output_dir) else os.getcwd())
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
                self.last_output_dir = os.path.abspath(chosen)
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
        ent_search = tk.Entry(search_frame, textvariable=search_var, bg="#090d16", fg="#ffffff", insertbackground="#ffffff", font=("Segoe UI", 9), width=24)
        ent_search.pack(side="left")
        ent_search.bind("<KeyRelease>", lambda e: filter_and_render_tree())

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

        tree.column("#0", width=260, minwidth=180)
        tree.column("mtime", width=120, minwidth=100, anchor="center")
        tree.column("size", width=80, minwidth=70, anchor="center")
        tree.column("status", width=100, minwidth=90, anchor="center")

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
            self.text_editor.delete("1.0", tk.END)
            self.text_editor.insert("1.0", raw_code)
            win.destroy()
            messagebox.showinfo("Thành công", f"Đã nạp thành công mã nguồn ULN từ file:\n{f_data.get('filename')}")

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
                icon_prefix = "📄 "
                
                item_id = tree.insert(
                    "",
                    "end",
                    text=f"{icon_prefix}{f['filename']}",
                    values=(f["mtime"], f"{f['size_kb']:.1f} KB", status_text)
                )
                file_data_map[item_id] = f

            status_count_lbl.config(
                text=f"Tìm thấy {uln_count} file có ULN (Hiển thị {matching_count} / {total_docx} files)",
                fg="#4ade80" if uln_count > 0 else "#94a3b8"
            )

            # Auto-select first item if available
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

        # Initial scan
        refresh_list()

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
                    import uln_parser
                    import uln_renderer
                    import uln_compiler
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

            msg = f"Successfully generated DOCX file:\n{compiled_file}"
            messagebox.showinfo("Success", msg)

        except Exception as e:
            try:
                self.root.deiconify()
            except Exception:
                pass
            messagebox.showerror("Compilation Error", f"Failed to generate Word document:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ULNFormatterApp(root)
    root.mainloop()
