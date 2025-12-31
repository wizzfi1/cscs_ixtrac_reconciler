import json
import os
import tempfile
from datetime import datetime

from core.paths import get_base_path
from config.rules import RULES_VERSION


class MappingError(Exception):
    pass


def _mapping_file_path():
    """
    Resolve mappings.json path correctly in both dev and packaged exe.
    """
    base = get_base_path()
    return os.path.join(base, "config", "mappings.json")


def load_mappings():
    path = _mapping_file_path()
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def validate_mapping(sheet, mapping):
    headers = [cell.value for cell in sheet[1]]
    missing = [mapping["name"], mapping["chn"]]

    for col in missing:
        if col not in headers:
            raise MappingError(f"Missing column: {col}")


def resolve_columns(sheet, mapping):
    """
    Resolve column indexes.
    - Input columns MUST exist
    - Output columns are CREATED if missing
    """
    raw_headers = [cell.value for cell in sheet[1]]

    headers = []
    for i, h in enumerate(raw_headers):
        if h is None or not str(h).strip():
            headers.append(f"UNNAMED:{i}")
        else:
            headers.append(str(h).strip())

    def require(col):
        if col not in headers:
            raise MappingError(
                f"Required column '{col}' not found.\n\n"
                "Please ensure this column exists in row 1."
            )
        return headers.index(col) + 1

    def ensure(col):
        if col in headers:
            return headers.index(col) + 1

        # Create column at end
        new_col_idx = len(headers) + 1
        sheet.cell(row=1, column=new_col_idx).value = col
        headers.append(col)
        return new_col_idx

    return {
        "name": require(mapping["name"]),
        "chn": require(mapping["chn"]),
        "membercode": ensure(mapping["membercode_out"]),
        "status": ensure(mapping["status_out"]),
    }



def save_mapping_safely(name, mapping):
    """
    Safely add a new mapping to mappings.json using atomic write.
    """
    path = _mapping_file_path()

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if name in data:
        raise ValueError(f"Mapping '{name}' already exists.")

    mapping = mapping.copy()
    mapping["rules_version"] = RULES_VERSION
    mapping["created_at"] = datetime.now().isoformat()

    data[name] = mapping

    # Atomic write
    fd, tmp_path = tempfile.mkstemp()
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as tmp:
            json.dump(data, tmp, indent=2)
        os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
