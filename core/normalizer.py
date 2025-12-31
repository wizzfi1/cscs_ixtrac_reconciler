import re

def normalize_name(name: str) -> str:
    if not isinstance(name, str):
        return ""
    name = name.upper()
    name = re.sub(r"[^A-Z ]", "", name)
    return " ".join(name.split())

def first_two_names(name: str) -> str:
    parts = normalize_name(name).split()
    return " ".join(parts[:2])
