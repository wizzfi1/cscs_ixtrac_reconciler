import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

from reconcile import run_reconciliation
from core.mapping import load_mappings
from wizard.wizard import MappingWizard


def main():
    root = TkinterDnD.Tk()
    root.title("IX TRAC Reconciler")
    root.geometry("560x360")
    root.resizable(False, False)

    # =========================
    # STYLES (AFTER ROOT!)
    # =========================
    style = ttk.Style(root)
    style.theme_use("vista")

    style.configure("Title.TLabel", font=("Segoe UI", 15, "bold"))
    style.configure("Subtitle.TLabel", font=("Segoe UI", 10), foreground="#555555")
    style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=10)
    style.configure("Secondary.TButton", padding=8)
    style.configure("Card.TFrame", background="white")

    App(root)
    root.mainloop()


class App:
    def __init__(self, root):
        self.root = root

        self.file_path = tk.StringVar(master=root)
        self.mappings = load_mappings()
        self.mapping_choice = tk.StringVar(
            master=root,
            value=list(self.mappings.keys())[0] if self.mappings else ""
        )

        card = ttk.Frame(root, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True, padx=20, pady=20)

        ttk.Label(card, text="CSCS ↔ IX TRAC Reconciliation Tool", style="Title.TLabel").pack(anchor="w")
        ttk.Label(card, text="Select an Excel file and mapping template.", style="Subtitle.TLabel").pack(anchor="w", pady=(0, 20))

        row = ttk.Frame(card)
        row.pack(fill="x", pady=(0, 15))

        ttk.Entry(row, textvariable=self.file_path, width=45).pack(side="left", padx=(0, 8))
        ttk.Button(row, text="Browse", command=self.browse).pack(side="left")

        ttk.Label(card, text="Select Template", style="Subtitle.TLabel").pack(anchor="w")
        self.mapping_menu = ttk.Combobox(card, textvariable=self.mapping_choice, values=list(self.mappings.keys()), state="readonly", width=38)
        self.mapping_menu.pack(anchor="w", pady=(5, 10))

        ttk.Button(card, text="Add New File Format", command=self.open_mapping_wizard).pack(anchor="w", pady=(0, 15))

        ttk.Label(card, text="Drag & drop Excel file here", style="Subtitle.TLabel").pack(pady=(0, 15))

        ttk.Button(
            card,
            text="Run Reconciliation",
            command=self.run,
            width=24
        ).pack(pady=(10, 0))



        root.drop_target_register(DND_FILES)
        root.dnd_bind("<<Drop>>", self.drop)

    def browse(self):
        f = filedialog.askopenfilename(parent=self.root, filetypes=[("Excel files", "*.xlsx")])
        if f:
            self.file_path.set(f)

    def drop(self, event):
        f = event.data.strip().strip("{}").strip('"')
        if f.lower().endswith(".xlsx") and os.path.exists(f):
            self.file_path.set(f)

    def open_mapping_wizard(self):
        wizard = MappingWizard(self.root)
        self.root.wait_window(wizard)
        self.reload_mappings()

    def reload_mappings(self):
        self.mappings = load_mappings()
        values = list(self.mappings.keys())
        self.mapping_menu["values"] = values
        if values:
            self.mapping_choice.set(values[0])

    def run(self):
        if not os.path.exists(self.file_path.get()):
            messagebox.showerror("Error", "Invalid Excel file.", parent=self.root)
            return
        run_reconciliation(self.file_path.get(), self.mapping_choice.get())
        messagebox.showinfo("Success", "Reconciliation completed.", parent=self.root)


if __name__ == "__main__":
    main()
