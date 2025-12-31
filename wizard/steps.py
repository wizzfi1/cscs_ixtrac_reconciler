import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

from wizard.preview import generate_preview
from core.mapping import save_mapping_safely


# =========================
# STEP 1 — WELCOME
# =========================
class WelcomeStep(tk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent)
        self.wizard = wizard

        tk.Label(
            self,
            text="Welcome to File Setup",
            font=("Segoe UI", 16, "bold")
        ).pack(pady=20)

        tk.Label(
            self,
            text=(
                "This guide will help you tell the system how your Excel file is structured.\n\n"
                "Nothing in your file will be changed."
            ),
            justify="center"
        ).pack(pady=10)

        tk.Button(self, text="Start Setup", command=wizard.next).pack(pady=30)


# =========================
# STEP 2 — FILE SELECT
# =========================
class FileSelectStep(tk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent)
        self.wizard = wizard
        self.file_var = tk.StringVar()

        tk.Label(
            self,
            text="Select Excel File",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=20)

        tk.Label(
            self,
            text="Choose the Excel file you want to set up.\nNothing will be modified.",
            justify="center"
        ).pack(pady=10)

        frame = tk.Frame(self)
        frame.pack(pady=10)

        tk.Entry(frame, textvariable=self.file_var, width=50).pack(side="left", padx=5)
        tk.Button(frame, text="Browse", command=self.browse).pack(side="left")

        nav = tk.Frame(self)
        nav.pack(fill="x", pady=30)

        tk.Button(nav, text="Back", command=wizard.back).pack(side="left", padx=40)
        tk.Button(nav, text="Next", command=self.validate).pack(side="right", padx=40)

    def browse(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if f:
            self.file_var.set(f)

    def validate(self):
        path = self.file_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("Invalid file", "Please select a valid Excel file.")
            return

        self.wizard.state.file_path = path
        self.wizard.next()


# =========================
# STEP 3 — SHEET SELECT
# =========================
class SheetSelectStep(tk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent)
        self.wizard = wizard

        tk.Label(
            self,
            text="Select Data Sheets",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=20)

        try:
            xl = pd.ExcelFile(wizard.state.file_path)
            sheets = xl.sheet_names
        except Exception:
            messagebox.showerror("Error", "Unable to read Excel file.")
            wizard.destroy()
            return

        # CSCS sheet
        tk.Label(self, text="CSCS sheet (source of membercodes):").pack(pady=(10, 0))
        self.cscs_var = tk.StringVar(value="CSCS" if "CSCS" in sheets else sheets[0])
        tk.OptionMenu(self, self.cscs_var, *sheets).pack()

        # IX TRAC sheet
        tk.Label(
            self,
            text="IX TRAC sheet (this sheet will be updated):",
            fg="darkred"
        ).pack(pady=(20, 0))
        self.ix_var = tk.StringVar(value="IX TRAC" if "IX TRAC" in sheets else sheets[0])
        tk.OptionMenu(self, self.ix_var, *sheets).pack()

        nav = tk.Frame(self)
        nav.pack(fill="x", pady=30)

        tk.Button(nav, text="Back", command=wizard.back).pack(side="left", padx=40)
        tk.Button(nav, text="Next", command=self.validate).pack(side="right", padx=40)

    def validate(self):
        if self.cscs_var.get() == self.ix_var.get():
            messagebox.showerror(
                "Invalid selection",
                "CSCS and IX TRAC sheets must be different."
            )
            return

        self.wizard.state.cscs_sheet = self.cscs_var.get()
        self.wizard.state.ixtrac_sheet = self.ix_var.get()

        self.wizard.state.mapping["cscs_sheet"] = self.cscs_var.get()
        self.wizard.state.mapping["ixtrac_sheet"] = self.ix_var.get()

        # Load headers from IX TRAC only (target sheet)
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
class ColumnMappingStep(tk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent)
        self.wizard = wizard
        headers = wizard.state.headers

        self.vars = {}
        fields = {
            "name": "Full Name column",
            "chn": "CHN / Account Number",
            "membercode_out": "Where MEMBERCODE will be written",
            "status_out": "Where MATCH STATUS will be written",
        }

        tk.Label(
            self,
            text="Match Columns",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=10)

        tk.Label(
            self,
            text="Select which columns in your file match each item below.",
            wraplength=600
        ).pack(pady=5)

        for key, label in fields.items():
            tk.Label(self, text=label).pack(anchor="w", padx=40, pady=(10, 0))
            var = tk.StringVar()
            tk.OptionMenu(self, var, *headers).pack(fill="x", padx=40)
            self.vars[key] = var

        nav = tk.Frame(self)
        nav.pack(fill="x", pady=30)

        tk.Button(nav, text="Back", command=wizard.back).pack(side="left", padx=40)
        tk.Button(nav, text="Next", command=self.validate).pack(side="right", padx=40)

    def validate(self):
        selected = [v.get() for v in self.vars.values()]

        if "" in selected:
            messagebox.showerror("Missing selection", "Please select all columns.")
            return

        if len(set(selected)) != len(selected):
            messagebox.showerror("Duplicate columns", "Each field must be unique.")
            return

        self.wizard.state.mapping.update(
            {k: v.get() for k, v in self.vars.items()}
        )
        self.wizard.next()


# =========================
# STEP 5 — PREVIEW
# =========================
class PreviewStep(tk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent)
        self.wizard = wizard

        tk.Label(
            self,
            text="Preview Mapping",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=20)

        preview = generate_preview(
            wizard.state.file_path,
            wizard.state.ixtrac_sheet,
            wizard.state.mapping["name"],
            wizard.state.mapping["chn"],
        )


        text = tk.Text(self, height=10, width=80)
        text.pack(padx=20, pady=10)
        text.insert("end", preview.to_string(index=False))
        text.config(state="disabled")

        self.confirm = tk.BooleanVar()
        tk.Checkbutton(
            self,
            text="This looks correct",
            variable=self.confirm
        ).pack(pady=10)

        nav = tk.Frame(self)
        nav.pack(fill="x", pady=20)

        tk.Button(nav, text="Back", command=wizard.back).pack(side="left", padx=40)
        tk.Button(nav, text="Next", command=self.validate).pack(side="right", padx=40)

    def validate(self):
        if not self.confirm.get():
            messagebox.showerror(
                "Confirmation required",
                "Please confirm the preview looks correct."
            )
            return
        self.wizard.next()


# =========================
# STEP 6 — SAVE
# =========================
class SaveStep(tk.Frame):
    def __init__(self, parent, wizard):
        super().__init__(parent)
        self.wizard = wizard

        tk.Label(
            self,
            text="Save This Setup",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=20)

        self.name_var = tk.StringVar()
        tk.Entry(self, textvariable=self.name_var, width=40).pack(pady=10)

        nav = tk.Frame(self)
        nav.pack(fill="x", pady=30)

        tk.Button(nav, text="Back", command=wizard.back).pack(side="left", padx=40)
        tk.Button(nav, text="Save", command=self.save).pack(side="right", padx=40)

    def save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Missing name", "Please provide a name.")
            return

        save_mapping_safely(name=name, mapping=self.wizard.state.mapping)

        messagebox.showinfo("Saved", "The file format has been saved successfully.")
        self.wizard.destroy()
