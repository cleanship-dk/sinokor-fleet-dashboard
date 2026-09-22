"""Persistence: the library of uploaded workbooks, and the vessel watchlist.

Both live as plain files under ``data/`` so nothing is lost when Streamlit restarts, and so the
user can copy the folder to another machine.
"""
from __future__ import annotations

import datetime as _dt
import hashlib
import json
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
WATCHLIST_FILE = DATA_DIR / "watchlist.json"
STAMP = "%Y%m%d-%H%M%S"


# ---------------------------------------------------------------------------- workbook library


@dataclass
class DataFile:
    path: Path
    uploaded: _dt.datetime
    original_name: str

    @property
    def label(self) -> str:
        return "%s  -  %s" % (self.uploaded.strftime("%d %b %Y %H:%M"), self.original_name)

    @property
    def size_kb(self) -> int:
        return round(self.path.stat().st_size / 1024)


def _parse_name(path: Path) -> tuple[_dt.datetime, str]:
    m = re.match(r"^(\d{8}-\d{6})__(.+)$", path.name)
    if m:
        try:
            return _dt.datetime.strptime(m.group(1), STAMP), m.group(2)
        except ValueError:
            pass
    return _dt.datetime.fromtimestamp(path.stat().st_mtime), path.name


def list_data_files() -> list[DataFile]:
    """Every stored workbook, newest first."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    files = []
    for path in DATA_DIR.glob("*.xls*"):
        if path.name.startswith("~$"):
            continue
        uploaded, original = _parse_name(path)
        files.append(DataFile(path=path, uploaded=uploaded, original_name=original))
    return sorted(files, key=lambda f: f.uploaded, reverse=True)


def save_upload(raw: bytes, original_name: str) -> tuple[Path, str | None]:
    """Store an uploaded workbook. Returns (path, note) where note flags a duplicate."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256(raw).hexdigest()
    for existing in list_data_files():
        if hashlib.sha256(existing.path.read_bytes()).hexdigest() == digest:
            return existing.path, (
                "This file is byte-for-byte identical to " + existing.original_name
                + " uploaded on " + existing.uploaded.strftime("%d %b %Y %H:%M")
                + ", so the existing copy was reused."
            )

    safe = re.sub(r"[^A-Za-z0-9._ -]+", "_", Path(original_name).name) or "workbook.xlsx"
    if not safe.lower().endswith((".xlsx", ".xlsm", ".xls")):
        safe += ".xlsx"
    # The stamp only has one-second resolution, so two uploads of the same name in the same
    # second would otherwise land on one path and the second would destroy the first.
    stem = _dt.datetime.now().strftime(STAMP)
    path = DATA_DIR / (stem + "__" + safe)
    suffix = 1
    while path.exists():
        path = DATA_DIR / ("%s-%d__%s" % (stem, suffix, safe))
        suffix += 1
    path.write_bytes(raw)
    return path, None


def delete_data_file(path: Path) -> None:
    path = Path(path)
    if path.parent.resolve() != DATA_DIR.resolve():
        raise ValueError("Refusing to delete a file outside the data folder.")
    archive = DATA_DIR / "_deleted"
    archive.mkdir(exist_ok=True)
    shutil.move(str(path), str(archive / path.name))


# ---------------------------------------------------------------------------- watchlist


def load_watchlist() -> dict[str, dict[str, Any]]:
    """Vessel name -> {note, added}. Highlighted rows survive a restart."""
    if not WATCHLIST_FILE.exists():
        return {}
    try:
        raw = json.loads(WATCHLIST_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    if isinstance(raw, list):                      # tolerate an older plain-list format
        return {str(v): {"note": "", "added": ""} for v in raw}
    if not isinstance(raw, dict):                  # a string, a number, null - start clean
        return {}
    return {
        str(k): {"note": str(v.get("note", "")), "added": str(v.get("added", ""))}
        for k, v in raw.items()
        if isinstance(v, dict)
    }


def save_watchlist(watchlist: dict[str, dict[str, Any]]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    WATCHLIST_FILE.write_text(
        json.dumps(watchlist, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def toggle_watch(watchlist: dict, vessel: str, on: bool, note: str = "") -> dict:
    if on:
        existing = watchlist.get(vessel, {})
        watchlist[vessel] = {
            "note": note or existing.get("note", ""),
            "added": existing.get("added") or _dt.date.today().isoformat(),
        }
    else:
        watchlist.pop(vessel, None)
    save_watchlist(watchlist)
    return watchlist


# ---------------------------------------------------------------------------- dump comparison

_COMPARE_FIELDS = {
    "last_drydock": "Last drydocking",
    "next_drydock": "Next drydock",
    "clean_confirmed": "Confirmed cleaning",
    "clean_next_rec": "Next recommended cleaning",
}


def diff_workbooks(old, new) -> dict[str, pd.DataFrame]:
    """What changed between two dumps: vessels added/removed, and per-field date changes."""
    # A dump can repeat a vessel name; keeping the first occurrence keeps the index unique, so
    # .loc below returns a row rather than a frame.
    old_fleet = old.fleet.drop_duplicates(subset="key", keep="first").set_index("key")
    new_fleet = new.fleet.drop_duplicates(subset="key", keep="first").set_index("key")

    added = new_fleet.loc[new_fleet.index.difference(old_fleet.index), ["vessel"]].reset_index(drop=True)
    removed = old_fleet.loc[old_fleet.index.difference(new_fleet.index), ["vessel"]].reset_index(drop=True)

    changes = []
    for key in new_fleet.index.intersection(old_fleet.index):
        before_row, after_row = old_fleet.loc[key], new_fleet.loc[key]
        for field, label in _COMPARE_FIELDS.items():
            if field not in old_fleet.columns or field not in new_fleet.columns:
                continue
            before, after = before_row[field], after_row[field]
            if pd.isna(before) and pd.isna(after):
                continue
            if before != after:
                changes.append({
                    "vessel": after_row["vessel"], "field": label,
                    "before": before, "after": after,
                })

    old_calls = _call_index(old)
    new_calls = _call_index(new)
    schedule = []
    for key_seq in sorted(set(old_calls) | set(new_calls)):
        before, after = old_calls.get(key_seq), new_calls.get(key_seq)
        if before == after:
            continue
        vessel = (after or before)["vessel"]
        schedule.append({
            "vessel": vessel,
            "call": key_seq[1],
            "before": _fmt_call(before),
            "after": _fmt_call(after),
        })

    return {
        "added": added,
        "removed": removed,
        "changes": pd.DataFrame(changes, columns=["vessel", "field", "before", "after"]),
        "schedule": pd.DataFrame(schedule, columns=["vessel", "call", "before", "after"]),
    }


def _call_index(wb) -> dict[tuple[str, int], dict]:
    if wb.calls.empty:
        return {}
    return {
        (row.key, row.seq): {"vessel": row.vessel, "hub": row.hub, "eta": row.eta, "etd": row.etd}
        for row in wb.calls.itertuples()
    }


def _fmt_call(call: dict | None) -> str:
    if not call:
        return "-"
    eta = call["eta"].strftime("%d %b %Y") if call["eta"] else "no ETA"
    return ("%s %s" % (call["hub"] or "?", eta)).strip()
