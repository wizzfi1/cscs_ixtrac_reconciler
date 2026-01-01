import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

from wizard.preview import generate_preview
from core.mapping import save_mapping_safely


# =========================
# STEP 1 — WELCOME
# =========================
class WelcomeStep(ttk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent, padding=20)
        self.wizard = wizard

        wizard.render_header("Welcome")

        ttk.Label(
            self,
            text="Welcome to File Setup",
            style="Title.TLabel"
        ).pack(pady=(20, 10))

        ttk.Label(
            self,
            text=(
                "This wizard will help you tell the system how your Excel file is structured.\n\n"
                "Your file will NOT be modified during setup."
            ),
            wraplength=600,
            justify="center"
        ).pack(pady=10)

        ttk.Button(
            self,
            text="Start Setup",
            style="Primary.TButton",
            command=wizard.next
        ).pack(pady=30)


# =========================
# STEP 2 — FILE SELECT
# =========================
class FileSelectStep(ttk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent, padding=20)
        self.wizard = wizard
        self.file_var = tk.StringVar(master=wizard)

        wizard.render_header("Select Excel File")

        ttk.Label(
            self,
            text="Choose the Excel file you want to configure.",
            wraplength=600
        ).pack(pady=10)

        row = ttk.Frame(self)
        row.pack(pady=10)

        ttk.Entry(row, textvariable=self.file_var, width=50).pack(side="left", padx=5)
        ttk.Button(row, text="Browse", command=self.browse).pack(side="left")

        nav = ttk.Frame(self)
        nav.pack(fill="x", pady=30)

        ttk.Button(nav, text="Back", command=wizard.back).pack(side="left")
        ttk.Button(nav, text="Next", command=self.validate).pack(side="right")

    def browse(self):
        f = filedialog.askopenfilename(
            parent=self.wizard,
            filetypes=[("Excel files", "*.xlsx")]
        )
        if f:
            self.file_var.set(f)

    def validate(self):
        path = self.file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror(
                "Invalid file",
                "Please select a valid Excel file.",
                parent=self.wizard
            )
            return

        self.wizard.state.file_path = path
        self.wizard.next()


# =========================
# STEP 3 — SHEET SELECT
# =========================
class SheetSelectStep(ttk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent, padding=20)
        self.wizard = wizard

        wizard.render_header("Select Sheets")

        try:
            xl = pd.ExcelFile(wizard.state.file_path)
            sheets = xl.sheet_names
        except Exception:
            messagebox.showerror(
                "Error",
                "Unable to read Excel file.",
                parent=wizard
            )
            wizard.destroy()
            return

        ttk.Label(self, text="CSCS sheet (source of membercodes):").pack(anchor="w")
        self.cscs_var = tk.StringVar(
            master=wizard,
            value="CSCS" if "CSCS" in sheets else sheets[0]
        )
        ttk.Combobox(
            self,
            values=sheets,
            textvariable=self.cscs_var,
            state="readonly"
        ).pack(anchor="w", pady=(5, 15))

        ttk.Label(
            self,
            text="IX TRAC sheet (this sheet will be updated):",
            foreground="darkred"
        ).pack(anchor="w")
        self.ix_var = tk.StringVar(
            master=wizard,
            value="IX TRAC" if "IX TRAC" in sheets else sheets[0]
        )
        ttk.Combobox(
            self,
            values=sheets,
            textvariable=self.ix_var,
            state="readonly"
        ).pack(anchor="w", pady=(5, 15))

        nav = ttk.Frame(self)
        nav.pack(fill="x", pady=30)

        ttk.Button(nav, text="Back", command=wizard.back).pack(side="left")
        ttk.Button(nav, text="Next", command=self.validate).pack(side="right")

    def validate(self):
        if self.cscs_var.get() == self.ix_var.get():
            messagebox.showerror(
                "Invalid selection",
                "CSCS and IX TRAC sheets must be different.",
                parent=self.wizard
            )
            return

        self.wizard.state.cscs_sheet = self.cscs_var.get()
        self.wizard.state.ixtrac_sheet = self.ix_var.get()

        self.wizard.state.mapping["cscs_sheet"] = self.cscs_var.get()
        self.wizard.state.mapping["ixtrac_sheet"] = self.ix_var.get()

        df = pd.read_excel(
            self.wizard.state.file_path,
            sheet_name=self.ix_var.get(),
            nrows=1
        )

        self.wizard.state.headers = [
            h for h in df.columns
            if isinstance(h, str) and not h.strip().upper().startswith("UNNAMED")
        ]

        self.wizard.next()


# =========================
# STEP 4 — COLUMN MAPPING
# =========================
class ColumnMappingStep(ttk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent, padding=20)
        self.wizard = wizard
        headers = wizard.state.headers

        wizard.render_header("Map Columns")

        self.vars = {}

        fields = {
            "name": ("Full Name column", "select"),
            "chn": ("CHN / Account Number", "select"),
            "membercode_out": ("Where MEMBERCODE will be written", "entry"),
            "status_out": ("Where MATCH STATUS will be written", "entry"),
        }

        for key, (label, mode) in fields.items():
            ttk.Label(self, text=label).pack(anchor="w", pady=(10, 0))
            var = tk.StringVar(master=wizard)

            if mode == "select":
                box = ttk.Combobox(self, values=headers, textvariable=var, state="readonly")
            else:
                box = ttk.Combobox(self, values=headers, textvariable=var, state="normal")
                if key == "membercode_out":
                    var.set("MEMBERCODE")
                if key == "status_out":
                    var.set("MATCH_STATUS")

            box.pack(fill="x")
            self.vars[key] = var

        nav = ttk.Frame(self)
        nav.pack(fill="x", pady=30)

        ttk.Button(nav, text="Back", command=wizard.back).pack(side="left")
        ttk.Button(nav, text="Next", command=self.validate).pack(side="right")

    def validate(self):
        values = {k: v.get().strip() for k, v in self.vars.items()}

        if not values["name"] or not values["chn"]:
            messagebox.showerror(
                "Missing selection",
                "Name and CHN must be selected from existing columns.",
                parent=self.wizard
            )
            return

        if not values["membercode_out"] or not values["status_out"]:
            messagebox.showerror(
                "Missing output column",
                "Output column names cannot be empty.",
                parent=self.wizard
            )
            return

        self.wizard.state.mapping.update(values)
        self.wizard.next()


# =========================
# STEP 5 — PREVIEW (FIXED)
# =========================
class PreviewStep(ttk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent, padding=20)
        self.wizard = wizard

        wizard.render_header("Preview Mapping")
        ttk.Label(
            self,
            text="This is a sample preview. Final results will be calculated when you run reconciliation.",
            style="Subtitle.TLabel"
        ).pack(pady=(0, 10))

        preview = generate_preview(
            wizard.state.file_path,
            wizard.state.ixtrac_sheet,
            wizard.state.mapping["name"],
            wizard.state.mapping["chn"],
        )

        text = tk.Text(self, height=10, width=80)
        text.pack(pady=10)
        text.insert("end", preview.to_string(index=False))
        text.config(state="disabled")

        self.confirm = tk.BooleanVar(master=wizard)
        ttk.Checkbutton(
            self,
            text="This looks correct",
            variable=self.confirm
        ).pack(pady=10)

        nav = ttk.Frame(self)
        nav.pack(fill="x", pady=20)

        ttk.Button(nav, text="Back", command=wizard.back).pack(side="left")
        ttk.Button(nav, text="Next", command=self.validate).pack(side="right")

    def validate(self):
        if not self.confirm.get():
            messagebox.showerror(
                "Confirmation required",
                "Please confirm the preview looks correct.",
                parent=self.wizard
            )
            return

        self.wizard.next()


# =========================
# STEP 6 — SAVE
# =========================
class SaveStep(ttk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent, padding=20)
        self.wizard = wizard

        wizard.render_header("Save Mapping")

        self.name_var = tk.StringVar(master=wizard)
        ttk.Entry(self, textvariable=self.name_var, width=40).pack(pady=10)

        nav = ttk.Frame(self)
        nav.pack(fill="x", pady=30)

        ttk.Button(nav, text="Back", command=wizard.back).pack(side="left")
        ttk.Button(nav, text="Save", command=self.save).pack(side="right")

    def save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror(
                "Missing name",
                "Please provide a name for this mapping.",
                parent=self.wizard
            )
            return

        save_mapping_safely(name=name, mapping=self.wizard.state.mapping)

        messagebox.showinfo(
            "Saved",
            "The file format has been saved successfully.",
            parent=self.wizard
        )
        self.wizard.destroy()
