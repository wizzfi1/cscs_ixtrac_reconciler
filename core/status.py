from config.rules import (
    STATUS_CONFIRMED,
    STATUS_CONFIRMED_2NAME,
    STATUS_AMBIGUOUS,
    STATUS_NOT_FOUND,
)

def assign_status(match_type, ambiguous=False):
    if ambiguous:
        return STATUS_AMBIGUOUS
    return STATUS_CONFIRMED if match_type == "EXACT" else STATUS_CONFIRMED_2NAME
