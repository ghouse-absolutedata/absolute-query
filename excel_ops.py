"""Spreadsheet behaviours with no UI: autofill series and filter matching."""
from __future__ import annotations

import datetime as dt
import re

BLANK = "(Blanks)"
_NUM = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)$")
_TAIL_INT = re.compile(r"^(.*?)(\d+)$")


def _num(v: str):
    return float(v) if _NUM.match(v.strip()) else None


def _date(v: str):
    v = v.strip()
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f"):
        try:
            return dt.datetime.strptime(v, fmt), fmt
        except ValueError:
            pass
    return None


def _fmt_num(x: float, like: list[str]) -> str:
    decimals = max((len(v.split(".")[1]) if "." in v else 0) for v in like)
    return f"{x:.{decimals}f}" if decimals else str(int(round(x)))


def fill_series(src: list[str], n: int) -> list[str]:
    """Continue `src` for `n` more cells the way spreadsheet autofill does.

    one number            -> copied            (5 -> 5, 5, 5)
    two+ numbers          -> linear step       (1, 3 -> 5, 7)
    dates (YYYY-MM-DD)    -> +1 day, or the step between the last two
    text ending in digits -> counts on         (Bond1 -> Bond2, ID009 -> ID010)
    anything else         -> repeats the source block
    """
    if n <= 0 or not src:
        return []
    nums = [_num(v) for v in src]
    if len(src) >= 2 and all(x is not None for x in nums):
        step = (nums[-1] - nums[0]) / (len(nums) - 1)
        return [_fmt_num(nums[-1] + step * (i + 1), src) for i in range(n)]
    dates = [_date(v) for v in src]
    if all(dates):
        last, fmt = dates[-1]
        step = (last - dates[-2][0]) if len(dates) >= 2 else dt.timedelta(days=1)
        out = [(last + step * (i + 1)).strftime(fmt) for i in range(n)]
        return [o[:-3] if fmt.endswith("%f") else o for o in out]  # back to milliseconds
    tails = [_TAIL_INT.match(v) for v in src]
    if all(tails) and len({m.group(1) for m in tails}) == 1 and all(x is None for x in nums):
        prefix, digits = tails[-1].group(1), tails[-1].group(2)
        step = int(digits) - int(tails[-2].group(2)) if len(tails) >= 2 else 1
        return [f"{prefix}{int(digits) + step * (i + 1):0{len(digits)}d}" for i in range(n)]
    return [src[i % len(src)] for i in range(n)]


def matches(value: str, spec: dict) -> bool:
    """spec = {"values": set | None, "text": str}; both must pass when given."""
    shown = value if value != "" else BLANK
    if spec.get("values") is not None and shown not in spec["values"]:
        return False
    text = spec.get("text", "").strip().lower()
    return not text or text in value.lower()
