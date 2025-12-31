from config.rules import (
    MAX_MEMBERCODE_LENGTH,
    INVALID_PREFIXES,
    STATUS_POSITION_CSCS,
    STATUS_MORE_THAN_5,
)

def validate_membercode(code):
    """
    CSCS membercode validation with strict precedence.
    Returns (is_valid, terminal_status)
    """
    if not isinstance(code, str) or not code.strip():
        return False, None

    code = code.strip()

    # HIGHEST PRIORITY: POSITION IN CSCS
    if code.startswith(INVALID_PREFIXES):
        return False, STATUS_POSITION_CSCS

    # SECOND PRIORITY: MORE THAN 5 CHARACTERS
    if len(code) > MAX_MEMBERCODE_LENGTH:
        return False, STATUS_MORE_THAN_5

    # VALID
    return True, None
