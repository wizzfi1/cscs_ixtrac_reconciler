# core/engine.py

from dataclasses import dataclass
from typing import Optional

from core.normalizer import normalize_name, first_two_names
from core.validator import validate_membercode
from config.rules import (
    STATUS_CONFIRMED,
    STATUS_CONFIRMED_2NAME,
    STATUS_AMBIGUOUS,
    STATUS_NOT_FOUND,
)


@dataclass
class MatchDecision:
    membercode: Optional[str]
    status: str
    reason: Optional[str] = None


def match_row(
    name: str | None,
    chn: str | None,
    exact_index: dict,
    two_name_index: dict,
) -> MatchDecision:
    """
    Core reconciliation decision engine.
    Deterministic, auditable, side-effect free.
    """

    if not name or not chn:
        return MatchDecision(None, STATUS_NOT_FOUND, "Missing name or CHN")

    norm = normalize_name(name)
    matches = exact_index.get((norm, chn), [])

    valid = []
    invalid_reason = None

    # ---------- EXACT MATCH ----------
    for m in matches:
        ok, reason = validate_membercode(m["MEMBERCODE"])
        if ok:
            valid.append(m)
        elif invalid_reason is None:
            invalid_reason = reason

    if len(valid) == 1:
        return MatchDecision(valid[0]["MEMBERCODE"], STATUS_CONFIRMED)

    if len(valid) > 1:
        return MatchDecision(None, STATUS_AMBIGUOUS, "Multiple valid exact matches")

    if invalid_reason:
        return MatchDecision(None, STATUS_NOT_FOUND, invalid_reason)

    # ---------- FALLBACK: FIRST TWO NAMES ----------
    first2 = first_two_names(name)
    matches = two_name_index.get((first2, chn), [])

    valid = []
    for m in matches:
        ok, _ = validate_membercode(m["MEMBERCODE"])
        if ok:
            valid.append(m)

    if len(valid) == 1:
        return MatchDecision(valid[0]["MEMBERCODE"], STATUS_CONFIRMED_2NAME)

    if len(valid) > 1:
        return MatchDecision(None, STATUS_AMBIGUOUS, "Multiple valid fallback matches")

    return MatchDecision(None, STATUS_NOT_FOUND, "No match found")
