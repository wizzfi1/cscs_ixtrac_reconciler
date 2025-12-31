import json
import os
from core.paths import get_base_path

class MappingError(Exception):
    pass

def load_mappings():
    base = get_base_path()
    path = os.path.join(base, "config", "mappings.json")

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def validate_mapping(sheet, mapping):
    headers = [cell.value for cell in sheet[1]]
    missing = [mapping["name"], mapping["chn"]]

    for col in missing:
        if col not in headers:
            raise MappingError(f"Missing column: {col}")

def resolve_columns(sheet, mapping):
    headers = [cell.value for cell in sheet[1]]

    def idx(col):
        return headers.index(col) + 1

    return {
        "name": idx(mapping["name"]),
        "chn": idx(mapping["chn"]),
        "membercode": idx(mapping["membercode_out"]),
        "status": idx(mapping["status_out"])
    }
