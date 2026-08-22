from tkinter import ttk

def setup_dark_theme(root):
    """Configures the dark space palette and TTK styling for the Tkinter desktop application."""
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

    root.option_add("*TCombobox*Listbox.background", "#090d16")
    root.option_add("*TCombobox*Listbox.foreground", "#f8fafc")
    root.option_add("*TCombobox*Listbox.selectBackground", "#3b82f6")
    root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
    root.option_add("*TCombobox*Listbox.font", ("Segoe UI", 9))

    # Configure Treeview dark styling
    style.configure("Treeview", background="#090d16", foreground="#f8fafc", fieldbackground="#090d16", rowheight=28, font=("Segoe UI", 9), borderwidth=0)
    style.map("Treeview", background=[("selected", "#2563eb")], foreground=[("selected", "#ffffff")])
    style.configure("Treeview.Heading", background="#1e293b", foreground="#38bdf8", font=("Segoe UI", 9, "bold"), borderwidth=1, relief="flat")
    style.map("Treeview.Heading", background=[("active", "#334155")])

    return style
