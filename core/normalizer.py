import re

def normalize_name(name: str) -> str:
    if not name:
        return ""
    return re.sub(r"[^A-Z ]", " ", str(name).upper()).strip()

def first_two_names(name: str) -> str:
    parts = normalize_name(name).split()
    return " ".join(parts[:2])
