import tkinter as tk
from tkinter import ttk

from wizard.state import WizardState
from wizard.steps import (
    WelcomeStep,
    FileSelectStep,
    SheetSelectStep,
    ColumnMappingStep,
    PreviewStep,
    SaveStep,
)


class MappingWizard(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("New File Mapping Setup")
        self.geometry("720x520")
        self.resizable(False, False)

        self.state = WizardState()
        self.steps = [
            WelcomeStep,
            FileSelectStep,
            SheetSelectStep,
            ColumnMappingStep,
            PreviewStep,
            SaveStep,
        ]
        self.step_index = 0

        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.show_step()

    def show_step(self):
        for w in self.container.winfo_children():
            w.destroy()

        step_class = self.steps[self.step_index]
        self.current = step_class(self.container, self)
        self.current.pack(fill="both", expand=True)

    def next(self):
        if self.step_index < len(self.steps) - 1:
            self.step_index += 1
            self.show_step()

    def back(self):
        if self.step_index > 0:
            self.step_index -= 1
            self.show_step()

    # =========================
    # STEP HEADER
    # =========================
    def render_header(self, title):
        header = ttk.Frame(self.container, padding=(20, 10))
        header.pack(fill="x")

        ttk.Label(
            header,
            text=f"Step {self.step_index + 1} of {len(self.steps)}",
            style="Subtitle.TLabel"
        ).pack(anchor="w")

        ttk.Label(
            header,
            text=title,
            style="Title.TLabel"
        ).pack(anchor="w")
