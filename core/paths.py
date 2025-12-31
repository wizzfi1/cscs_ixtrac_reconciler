import sys
import os

def get_base_path():
    """
    Returns the base path for resources.
    Works for both normal Python runs and PyInstaller executables.
    """
    if hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS  # PyInstaller temp folder
    return os.path.abspath(".")
