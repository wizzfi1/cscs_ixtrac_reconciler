from config.rules import MAX_MEMBERCODE_LENGTH, INVALID_PREFIXES

def validate_membercode(code):
    if code is None:
        return False, None

    code = str(code)

    if len(code) > MAX_MEMBERCODE_LENGTH:
        return False, "MORE THAN 5"

    if code.startswith(INVALID_PREFIXES):
        return False, "POSITION IN CSCS"

    return True, None
