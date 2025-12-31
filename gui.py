import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

from reconcile import run_reconciliation
from core.mapping import load_mappings
from wizard.wizard import MappingWizard

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("IX TRAC Reconciler")
        self.root.geometry("520x320")
        self.root.resizable(False, False)

        self.file_path = tk.StringVar()
        self.mappings = load_mappings()
        self.mapping_choice = tk.StringVar(value=list(self.mappings.keys())[0])

        tk.Label(root, text="CSCS ↔ IX TRAC Reconciliation Tool",
                 font=("Segoe UI", 12, "bold")).pack(pady=10)

        frame = tk.Frame(root)
        frame.pack()

        tk.Entry(frame, textvariable=self.file_path, width=45).pack(side=tk.LEFT, padx=5)
        tk.Button(frame, text="Browse", command=self.browse).pack(side=tk.LEFT)

        tk.Label(root, text="Select Template").pack(pady=5)

        self.mapping_menu = tk.OptionMenu(root, self.mapping_choice, *self.mappings.keys())
        self.mapping_menu.pack()

        tk.Button(
            root,
            text="Add New File Format",
            command=self.open_mapping_wizard,
            width=25
        ).pack(pady=5)

        tk.Label(root, text="Drag & drop Excel file here").pack(pady=5)

        tk.Button(root, text="Run Reconciliation",
                  command=self.run, bg="#0078D4", fg="white", width=25).pack(pady=15)

        root.drop_target_register(DND_FILES)
        root.dnd_bind("<<Drop>>", self.drop)

    def browse(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if f:
            self.file_path.set(f)

    def drop(self, event):
        f = event.data.strip("{}")
        if f.endswith(".xlsx"):
            self.file_path.set(f)

    def open_mapping_wizard(self):
        wizard = MappingWizard(self.root)
        self.root.wait_window(wizard)
        self.reload_mappings()

    def reload_mappings(self):
        self.mappings = load_mappings()

        menu = self.mapping_menu["menu"]
        menu.delete(0, "end")

        for name in self.mappings.keys():
            menu.add_command(
                label=name,
                command=lambda v=name: self.mapping_choice.set(v)
            )

        if self.mappings:
            self.mapping_choice.set(list(self.mappings.keys())[0])

    def run(self):
        file = self.file_path.get()
        if not os.path.exists(file):
            messagebox.showerror("Error", "Invalid Excel file.")
            return
        try:
            run_reconciliation(file, self.mapping_choice.get())
            messagebox.showinfo("Success", "Reconciliation completed.\nOutput saved to /output")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def main():
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
