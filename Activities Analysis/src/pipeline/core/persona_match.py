"""Label normalization and matching between matrices and client workbook."""

from __future__ import annotations

import re
from typing import Any


_WS_RE = re.compile(r"\s+")
_PUNCT_RE = re.compile(r"[^a-z0-9\s]")


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip().lower().replace("\n", " ")
    s = _PUNCT_RE.sub(" ", s)
    s = _WS_RE.sub(" ", s).strip()
    return s


def infer_test_id(label: Any) -> str:
    s = normalize_label(label)
    if not s:
        return ""

    if "vo2" in s:
        return "vo2_max"
    if "fev1" in s:
        return "fev1"
    if "grip" in s:
        return "grip_strength"
    if "sts" in s and "power" in s:
        return "sts_power"
    if "vertical" in s and "jump" in s:
        return "vertical_jump"
    if "body" in s and "fat" in s:
        return "body_fat"
    if ("waist" in s and "height" in s) or ("w h" in s and "ratio" in s) or ("w h ratio" in s):
        return "waist_height_ratio"
    if "hba1c" in s:
        return "hba1c"
    if "homa" in s:
        return "homa_ir"
    if "apob" in s or s == "apob":
        return "apob"
    if "hscrp" in s or ("hs" in s and "crp" in s):
        return "hscrp"
    if "gait" in s and "speed" in s:
        return "gait_speed"
    if "tug" in s or ("timed" in s and "go" in s):
        return "tug"
    if "single" in s and "leg" in s:
        return "single_leg_stance"
    if "sit" in s and "reach" in s:
        return "sit_and_reach"
    if "processing" in s and "speed" in s:
        return "processing_speed"
    if "working" in s and "memory" in s:
        return "working_memory"

    # Fall back to normalized label to keep unknown tests stable.
    return s


def is_meta_column(col_name: Any) -> bool:
    s = normalize_label(col_name)
    return s in {"unnamed 0", "activity", "evidence quality", "key references"} or not s

