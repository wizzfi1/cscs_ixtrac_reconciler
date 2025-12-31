import tkinter as tk
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
        self.geometry("700x500")
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
        self.container = tk.Frame(self)
        self.container.pack(fill="both", expand=True)

        self.show_step()

    def show_step(self):
        for w in self.container.winfo_children():
            w.destroy()

        step_class = self.steps[self.step_index]
        self.current = step_class(self.container, self)
        self.current.pack(fill="both", expand=True)

    def next(self):
        self.step_index += 1
        self.show_step()

    def back(self):
        self.step_index -= 1
        self.show_step()
