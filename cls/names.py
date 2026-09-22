"""Vessel-name normalisation.

The workbook spells the same vessel differently on every sheet ("HEUNG-A YOUNG" /
"Heung A Young", "TOYAMA TRADER" / "Toyoma Trader", "VOSTOCHNY" / "Volostchny",
"HAKATA EXPRESS" / "HAKATA EXPRESSS"). Joins between sheets therefore go through a
normalised key plus a small fuzzy fallback, never through the raw string.
"""
from __future__ import annotations

import difflib
import re
import unicodedata

# Spelling mistakes seen in the source workbook, keyed by normalised form.
_ALIASES: dict[str, str] = {
    "hakataexpresss": "hakataexpress",
    "toyomatrader": "toyamatrader",
    "volostchnyvoyager": "vostochnyvoyager",
    "heaungaakita": "heungaakita",
    "hanshungweihai": "hansungweihai",
    "altanticsouth": "atlanticsouth",
    "pacificnongbo": "pacificningbo",
}


def normalise(name: object) -> str:
    """Collapse a vessel name to a join key: lowercase, letters and digits only."""
    if name is None:
        return ""
    text = unicodedata.normalize("NFKC", str(name))
    text = text.replace("&", "and")
    text = re.sub(r"[^0-9A-Za-z]+", "", text).lower()
    return _ALIASES.get(text, text)


def display(name: object) -> str:
    """Tidy a raw cell value for display: collapse whitespace, strip, upper-case."""
    if name is None:
        return ""
    return re.sub(r"\s+", " ", str(name)).strip().upper()


def build_index(names) -> dict[str, str]:
    """Map every normalised key to the canonical display name."""
    return {normalise(n): display(n) for n in names if display(n)}


def match(name: object, index: dict[str, str], cutoff: float = 0.90) -> str | None:
    """Resolve ``name`` against a canonical index, falling back to close-match fuzzy lookup.

    The fuzzy step is deliberately timid. This fleet contains genuinely different vessels whose
    names are one or two letters apart - HAKATA VOYAGER and JAKARTA VOYAGER score 0.889, NINGBO
    TRADER and NINGBO VOYAGER share a word - and crediting one vessel's cleaning to another is
    far worse than leaving a row unmatched, which the data-quality tab reports. So a fuzzy hit
    is only accepted when it is clearly the single best answer: it must clear the cutoff, and no
    runner-up may be within ``margin`` of it.
    """
    key = normalise(name)
    if not key:
        return None
    if key in index:
        return index[key]
    close = difflib.get_close_matches(key, list(index), n=2, cutoff=cutoff)
    if not close:
        return None
    best = difflib.SequenceMatcher(None, key, close[0]).ratio()
    if len(close) > 1:
        runner_up = difflib.SequenceMatcher(None, key, close[1]).ratio()
        if best - runner_up < 0.05:
            return None
    return index[close[0]]
