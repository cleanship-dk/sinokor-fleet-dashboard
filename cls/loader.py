"""Read a Sinokor fleet workbook into tidy dataframes.

Everything here is defensive on purpose: the point of this dashboard is that the user drops in a
*new* Excel dump each month, and that dump will have moved columns, renamed headers, gained rows
and changed date formats. So sheets are found by fuzzy name, the header row is located by content,
and columns are mapped by normalised header text rather than by position.
"""
from __future__ import annotations

import datetime as _dt
import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd

from .dates import coerce_date, infer_dayfirst, is_datelike_text
from .names import display, match as fuzzy_match, normalise

# ---------------------------------------------------------------------------- sheet discovery

# Ordered most-specific first: the substring pass takes the first alias that hits, so a loose
# alias like "fleet" must never be tried before "fleet schedule" (otherwise a workbook whose
# fleet sheet has been renamed silently resolves to "Fleet Contact Info").
SHEET_ALIASES = {
    "fleet": ("fleet schedule", "vessel schedule", "schedule", "fleet"),
    "orders": ("order log", "job log", "orders", "order"),
    "contacts": ("fleet contact info", "contact info", "contacts", "contact"),
    "prices": ("price list", "pricing", "rates", "price"),
    "invoices": ("invoice", "invoices", "billing"),
    "asat": ("sheet1", "summary", "cover"),
}

# Resolved in this order; each role excludes the sheets already claimed before it, so two roles
# can never share one sheet.
SHEET_RESOLUTION_ORDER = ("contacts", "orders", "invoices", "prices", "asat", "fleet")

# Canonical field -> the header spellings that map onto it.
FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "vessel": ("vessel name", "vessel", "ship name", "ship"),
    "pic": ("vessel pic", "pic", "superintendent", "supt"),
    "pic_email": ("pic email", "email", "pic e mail"),
    "group_email": ("group email", "group e mail"),
    "imo": ("imo", "imo no", "imo number"),
    "vessel_type": ("vessel particular", "type", "vessel type"),
    "loa": ("loa m", "loa", "loa metres", "length"),
    "last_drydock": ("last dry docking", "last drydock", "last dd", "last docking"),
    "next_drydock": ("next dry dock", "next drydock", "next dd", "next docking"),
    "uwi_confirmed": ("confirmed inspection last next", "confirmed inspection"),
    "uwi_next_rec": ("next recommended inspection",),
    "ncps_last": ("last inspection cleaning",),
    "ncps_next_rec": ("next recommended inspection cleaning",),
    "clean_confirmed": ("confirmed cleaning last next", "confirmed cleaning", "last cleaning"),
    "clean_next_rec": (
        "next reccommended hull cleaning", "next recommended hull cleaning", "next cleaning",
    ),
    "recommend_action": ("recommend action", "recommended action", "action"),
    "remarks": ("remarks", "remark", "notes", "comment"),
}

# Headers that repeat once per scheduled port call.
CALL_FIELDS = {
    "hub": ("next cls hub", "cls hub", "hub", "port"),
    "terminal": ("terminal", "berth"),
    "eta": ("eta", "arrival"),
    "etd": ("etd", "departure"),
}

DATE_FIELDS = [
    "last_drydock", "next_drydock", "clean_confirmed", "clean_next_rec",
    "uwi_confirmed", "uwi_next_rec", "ncps_last", "ncps_next_rec",
]


def _norm_header(value: Any) -> str:
    """Normalise a header cell: strip newlines/punctuation, collapse space, lowercase.

    Non-ASCII letters are preserved because several columns on the contact sheet are only
    labelled in Korean (본선 메일, 대리점 정보, 북항터미널...).
    """
    if value is None:
        return ""
    text = re.sub(r"[\r\n]+", " ", str(value))
    text = re.sub(r"[^\w ]+", " ", text, flags=re.UNICODE)
    return re.sub(r"\s+", " ", text).strip().lower()


def _is(header: str, aliases: tuple[str, ...]) -> bool:
    """Compare ignoring spacing, so "Last DryDocking" still matches "last dry docking"."""
    squashed = header.replace(" ", "")
    return any(squashed == alias.replace(" ", "") for alias in aliases)


def find_sheet(book: openpyxl.Workbook, key: str,
               taken: set[str] | None = None) -> str | None:
    """Find the sheet playing ``key``, ignoring any sheet already claimed by another role."""
    wanted = SHEET_ALIASES[key]
    names = [n for n in book.sheetnames if n not in (taken or set())]
    for want in wanted:                                   # exact (normalised) match first
        for name in names:
            if _norm_header(name) == want:
                return name
    for want in wanted:                                   # then substring
        for name in names:
            if want in _norm_header(name):
                return name
    return None


def score_as_fleet_sheet(book: openpyxl.Workbook, name: str) -> int:
    """How much this sheet looks like the fleet schedule, judged on its headers alone."""
    rows = [list(r) for r in book[name].iter_rows(values_only=True, max_row=12)]
    if not rows:
        return -1
    headers = [_norm_header(c) for c in rows[_locate_header_row(rows)]]
    if not any(_is(h, FIELD_ALIASES["vessel"]) for h in headers):
        return -1                                   # no vessel column: cannot be the fleet sheet
    score = sum(1 for aliases in FIELD_ALIASES.values() for h in headers if _is(h, aliases))
    score += sum(1 for h in headers if any(_is(h, a) for a in CALL_FIELDS.values()))
    return score


def resolve_sheets(book: openpyxl.Workbook) -> dict[str, str | None]:
    """Assign one sheet to each role, never the same sheet twice.

    The supporting sheets are matched on name. The fleet sheet is then chosen from whatever is
    left by looking at its *headers*, because it is the one sheet the dashboard cannot do
    without - and a renamed fleet sheet must not quietly resolve to something else.
    """
    resolved: dict[str, str | None] = {}
    taken: set[str] = set()
    for key in SHEET_RESOLUTION_ORDER:
        if key == "fleet":
            continue
        name = find_sheet(book, key, taken)
        resolved[key] = name
        if name:
            taken.add(name)

    candidates = [(score_as_fleet_sheet(book, n), n) for n in book.sheetnames if n not in taken]
    candidates = [(score, n) for score, n in candidates if score > 0]
    by_name = find_sheet(book, "fleet", taken)
    if by_name is not None and score_as_fleet_sheet(book, by_name) > 0:
        resolved["fleet"] = by_name                 # named like the fleet sheet and looks like it
    elif candidates:
        resolved["fleet"] = max(candidates)[1]      # otherwise the best-looking sheet left
    else:
        # Nothing unclaimed looks like a fleet schedule; fall back to any sheet that does, even
        # one already claimed, so the caller can report the conflict rather than mis-parsing.
        everything = [(score_as_fleet_sheet(book, n), n) for n in book.sheetnames]
        everything = [(score, n) for score, n in everything if score > 0]
        resolved["fleet"] = max(everything)[1] if everything else None
    return {key: resolved.get(key) for key in SHEET_ALIASES}


# ---------------------------------------------------------------------------- result types


@dataclass
class LoadIssue:
    severity: str            # "error" | "warning" | "info"
    sheet: str
    vessel: str
    field: str
    detail: str
    raw: str = ""


@dataclass
class Workbook:
    """Everything the dashboard needs, already tidied."""

    path: Path
    fleet: pd.DataFrame                     # one row per vessel
    calls: pd.DataFrame                     # one row per scheduled port call
    orders: pd.DataFrame                    # one row per historical job
    invoices: pd.DataFrame                  # one row per billed job leg
    contacts: pd.DataFrame                  # one row per vessel (enrichment only)
    prices: dict[str, Any]
    as_at: _dt.date | None
    job_checksum: dict[str, int] = field(default_factory=dict)
    issues: list[LoadIssue] = field(default_factory=list)
    sheet_map: dict[str, str | None] = field(default_factory=dict)
    cleaning_interval_months: int = 12
    inspection_interval_months: int = 6
    digest: str = ""

    @property
    def vessel_names(self) -> list[str]:
        return sorted(self.fleet["vessel"].tolist())


# ---------------------------------------------------------------------------- fleet sheet


def _locate_header_row(rows: list[list[Any]], max_scan: int = 12) -> int:
    """The header row is the row that matches the most known field names."""
    best, best_score = 0, -1
    for idx, row in enumerate(rows[:max_scan]):
        headers = [_norm_header(c) for c in row]
        score = sum(1 for aliases in FIELD_ALIASES.values() for h in headers if _is(h, aliases))
        score += sum(1 for h in headers if any(_is(h, a) for a in CALL_FIELDS.values()))
        if any(_is(h, FIELD_ALIASES["vessel"]) for h in headers):
            score += 5
        if score > best_score:
            best, best_score = idx, score
    return best


def _map_columns(headers: list[str]) -> tuple[dict[str, int], list[dict[str, int]]]:
    """Return (scalar field -> column index, [per-port-call blocks])."""
    colmap: dict[str, int] = {}
    for field_name, aliases in FIELD_ALIASES.items():
        for idx, header in enumerate(headers):
            if _is(header, aliases) and field_name not in colmap:
                colmap[field_name] = idx

    # Every 'hub' header opens a new port-call block; terminal/eta/etd attach to the nearest
    # preceding hub. This survives a dump carrying 1, 2 or 5 scheduled calls per vessel.
    blocks: list[dict[str, int]] = []
    for idx, header in enumerate(headers):
        if _is(header, CALL_FIELDS["hub"]):
            # A hub header opens a block, unless the current one is still waiting for its own.
            if blocks and "hub" not in blocks[-1]:
                blocks[-1]["hub"] = idx
            else:
                blocks.append({"hub": idx})
            continue
        matched = next((k for k in ("terminal", "eta", "etd") if _is(header, CALL_FIELDS[k])), None)
        if matched is None:
            continue
        # Start a block on the first field of any kind, so a dump that writes eta/etd/hub in a
        # different order still yields complete calls rather than silently dropping them.
        if not blocks or matched in blocks[-1]:
            blocks.append({})
        blocks[-1][matched] = idx
    return colmap, [b for b in blocks if "eta" in b]


def _cell(row: list[Any], idx: int | None) -> Any:
    """Read a cell, treating whitespace-only strings as empty."""
    if idx is None or idx >= len(row):
        return None
    value = row[idx]
    if isinstance(value, str) and not value.strip():
        return None
    return value


def _text(row: list[Any], idx: int | None) -> str:
    value = _cell(row, idx)
    return "" if value is None else re.sub(r"\s+", " ", str(value)).strip()


def _read_fleet(book: openpyxl.Workbook, sheet_name: str, issues: list[LoadIssue]):
    ws = book[sheet_name]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    header_idx = _locate_header_row(rows)
    headers = [_norm_header(c) for c in rows[header_idx]]
    colmap, call_blocks = _map_columns(headers)

    if "vessel" not in colmap:
        raise ValueError(
            "Could not find a vessel-name column on sheet " + repr(sheet_name)
            + ". Headers found on row " + str(header_idx + 1) + ": "
            + ", ".join(h for h in headers if h)
        )
    for required in ("last_drydock", "clean_confirmed"):
        if required not in colmap:
            issues.append(LoadIssue(
                "warning", sheet_name, "", required,
                "Column not present in this dump, so the checks that depend on it are disabled.",
            ))
    if not call_blocks:
        issues.append(LoadIssue(
            "warning", sheet_name, "", "eta",
            "No ETA columns found, so port-call filtering is disabled.",
        ))

    body = rows[header_idx + 1:]

    # Per-column day-first inference, so a dump that flips to US date format still parses.
    dayfirst: dict[int, bool] = {}

    def resolve_convention(idx: int, label: str) -> None:
        values = [_cell(r, idx) for r in body]
        decided, guessed = infer_dayfirst(values, report=True)
        dayfirst[idx] = decided
        if guessed:
            issues.append(LoadIssue(
                "warning", sheet_name, "", label,
                "This column is written as text like 01/10/2026 with no value that settles "
                "whether the day or the month comes first, so day-first was assumed. If this "
                "dump uses US dates, every date in the column is wrong by up to eleven months.",
            ))

    for field_name in [f for f in DATE_FIELDS if f in colmap]:
        resolve_convention(colmap[field_name], field_name)
    for seq, block in enumerate(call_blocks, start=1):
        for key in ("eta", "etd"):
            if key in block:
                resolve_convention(block[key], "call" + str(seq) + "." + key)

    def as_date(row, idx, vessel, field_name):
        raw = _cell(row, idx)
        if raw is None:
            return None
        value = coerce_date(raw, dayfirst=dayfirst.get(idx, True))
        if value is None:
            issues.append(LoadIssue("warning", sheet_name, vessel, field_name,
                                    "Not a recognisable date, so it was ignored.", str(raw)))
        elif is_datelike_text(raw):
            issues.append(LoadIssue("info", sheet_name, vessel, field_name,
                                    "Stored as text, read as " + value.strftime("%d %b %Y") + ".",
                                    str(raw)))
        return value

    fleet_rows, call_rows = [], []
    for row in body:
        vessel = display(_cell(row, colmap["vessel"]))
        if not vessel:
            continue
        # Footer blocks such as "Cleaning interval | 12 | months" sit below the data with a label
        # in the vessel column, so a row only counts as a vessel if some other field is populated.
        if not any(_cell(row, i) is not None for f, i in colmap.items() if f != "vessel"):
            continue

        imo_raw = _cell(row, colmap.get("imo"))
        imo = None
        if imo_raw is not None:
            try:
                imo = int(float(imo_raw))
            except (TypeError, ValueError):
                issues.append(LoadIssue("warning", sheet_name, vessel, "imo",
                                        "IMO is not numeric.", str(imo_raw)))
        loa_raw = _cell(row, colmap.get("loa"))
        try:
            loa = float(loa_raw) if loa_raw is not None else None
        except (TypeError, ValueError):
            loa = None
            issues.append(LoadIssue(
                "warning", sheet_name, vessel, "loa",
                "LOA is not numeric, so no cleaning price can be estimated for this vessel.",
                str(loa_raw),
            ))

        record: dict[str, Any] = {
            "vessel": vessel,
            "key": normalise(vessel),
            "pic": _text(row, colmap.get("pic")),
            "pic_email": _text(row, colmap.get("pic_email")),
            "group_email": _text(row, colmap.get("group_email")),
            "imo": imo,
            "vessel_type": _text(row, colmap.get("vessel_type")),
            "loa": loa,
            "recommend_action": _text(row, colmap.get("recommend_action")),
            "remarks": _text(row, colmap.get("remarks")),
        }
        for field_name in DATE_FIELDS:
            record[field_name] = as_date(row, colmap.get(field_name), vessel, field_name)
        fleet_rows.append(record)

        for seq, block in enumerate(call_blocks, start=1):
            eta = as_date(row, block.get("eta"), vessel, "call" + str(seq) + ".eta")
            etd = as_date(row, block.get("etd"), vessel, "call" + str(seq) + ".etd")
            hub = _text(row, block.get("hub"))
            if eta is None and not hub:
                continue
            call_rows.append({
                "vessel": vessel,
                "key": normalise(vessel),
                "seq": seq,
                "hub": hub,
                "hub_key": normalise(hub),
                "terminal": _text(row, block.get("terminal")),
                "eta": eta,
                "etd": etd,
            })

    fleet = pd.DataFrame(fleet_rows)
    if not fleet.empty:
        repeated = fleet.loc[fleet["key"].duplicated(keep=False), "vessel"].unique()
        for vessel in repeated:
            issues.append(LoadIssue(
                "warning", sheet_name, vessel, "vessel",
                "This vessel appears on more than one row; only the first is used.",
            ))
        fleet = fleet.drop_duplicates(subset="key", keep="first").reset_index(drop=True)
    calls = pd.DataFrame(
        call_rows,
        columns=["vessel", "key", "seq", "hub", "hub_key", "terminal", "eta", "etd"],
    )

    # Interval constants live in a footer block below the vessel rows.
    intervals: dict[str, int] = {}
    for row in body:
        label = _norm_header(next((c for c in row if isinstance(c, str) and c.strip()), None))
        if "cleaning interval" in label or "inspection interval" in label:
            numbers = [c for c in row if isinstance(c, (int, float)) and not isinstance(c, bool)]
            if numbers:
                intervals["cleaning" if "cleaning" in label else "inspection"] = int(numbers[0])
    return fleet, calls, intervals


# ---------------------------------------------------------------------------- other sheets

ORDER_COLUMNS = ["vessel", "key", "matched", "scope", "eta", "hub", "status",
                 "remarks", "report", "invoiced", "fouling"]


def _read_orders(book, sheet_name, canonical: dict[str, str], issues) -> pd.DataFrame:
    if not sheet_name:
        issues.append(LoadIssue(
            "warning", "", "", "orders",
            "No order-log sheet was found in this workbook, so past cleanings are known only "
            "from the Fleet Schedule. Vessels may look more overdue than they are.",
        ))
        return pd.DataFrame(columns=ORDER_COLUMNS)
    rows = [list(r) for r in book[sheet_name].iter_rows(values_only=True)]
    if not rows:
        return pd.DataFrame(columns=ORDER_COLUMNS)
    headers = [_norm_header(c) for c in rows[0]]

    def idx_of(*names, contains=None):
        for name in names:
            if name in headers:
                return headers.index(name)
        if contains:
            for i, h in enumerate(headers):
                if contains in h:
                    return i
        return None

    ix = {
        "vessel": idx_of("vessel name", "vessel"),
        "scope": idx_of("scope", "service"),
        "eta": idx_of("eta", "date", "service date"),
        "hub": idx_of("hub", "port", "location"),
        "status": idx_of("status"),
        "remarks": idx_of("remarks", "remark", "notes"),
        "report": idx_of("report"),
        "invoiced": idx_of("invoiced"),
        "fouling": idx_of(contains="fouling"),
    }
    if ix["vessel"] is None:
        issues.append(LoadIssue("warning", sheet_name, "", "vessel",
                                "No vessel column on the order log, so cleaning history comes "
                                "from the Fleet Schedule alone."))
        return pd.DataFrame(columns=ORDER_COLUMNS)

    body = rows[1:]
    eta_dayfirst = infer_dayfirst(_cell(r, ix["eta"]) for r in body) if ix["eta"] is not None else True

    out = []
    for row in body:
        name = display(_cell(row, ix["vessel"]))
        if not name:
            continue
        matched = canonical.get(normalise(name)) or fuzzy_match(name, canonical)
        if matched is None:
            issues.append(LoadIssue(
                "warning", sheet_name, name, "vessel",
                "This order-log vessel does not match any vessel on the Fleet Schedule, so its "
                "history is not attached to any row.",
            ))
        fouling_raw = _cell(row, ix["fouling"])
        try:
            fouling = float(fouling_raw) if fouling_raw is not None else None
        except (TypeError, ValueError):
            fouling = None
        out.append({
            "vessel": matched or name,
            "key": normalise(matched or name),
            "matched": matched is not None,
            "scope": _text(row, ix["scope"]),
            "eta": coerce_date(_cell(row, ix["eta"]), dayfirst=eta_dayfirst),
            "hub": _text(row, ix["hub"]),
            "status": _text(row, ix["status"]),
            "remarks": _text(row, ix["remarks"]),
            "report": _text(row, ix["report"]),
            "invoiced": _text(row, ix["invoiced"]),
            "fouling": fouling,
        })
    return pd.DataFrame(out, columns=ORDER_COLUMNS)


CONTACT_COLUMNS = ["key", "vessel", "code", "call_cycle_days", "vessel_email", "agency",
                   "busan_terminal", "singapore_terminal", "other_ports",
                   "alt_drydock", "alt_next_drydock"]


def _read_contacts(book, sheet_name, canonical: dict[str, str]) -> pd.DataFrame:
    """Only the columns that are actually trustworthy.

    This sheet's LAST/NEXT CLEANING columns are mislabelled: they frequently hold DRYDOCK dates,
    not cleaning dates. The give-away is the Korean text in the NEXT column - 차기입거 means "next
    docking" - and the arithmetic: for the rows where the Fleet Schedule's own drydock dates are
    unusable, this sheet's LAST value plus five years reproduces them to within a few days.

    They are therefore read as ``alt_drydock`` / ``alt_next_drydock``, never as cleaning history,
    and the rule engine uses them only as a labelled fallback when the Fleet Schedule has no
    usable last-drydocking date.
    """
    if not sheet_name:
        return pd.DataFrame(columns=CONTACT_COLUMNS)
    rows = [list(r) for r in book[sheet_name].iter_rows(values_only=True)]
    header_idx = next(
        (i for i, r in enumerate(rows[:8])
         if any(_norm_header(c) in ("vessel name", "vessel") for c in r)),
        0,
    )
    headers = [_norm_header(c) for c in rows[header_idx]]

    def idx_of(*names, contains=None):
        for name in names:
            if name in headers:
                return headers.index(name)
        if contains:
            for i, h in enumerate(headers):
                if contains in h:
                    return i
        return None

    ix = {
        "vessel": idx_of("vessel name", "vessel"),
        "code": idx_of("code"),
        "cycle": idx_of(contains="day"),
        "email": idx_of(contains="본선 메일") or idx_of(contains="메일"),
        "agency": idx_of(contains="대리점"),
        "busan": idx_of(contains="북항터미널"),
        "singapore": idx_of(contains="싱가폴터미널"),
        "other": idx_of(contains="부산외"),
        "alt_dd": idx_of(contains="last cleaning") or idx_of(contains="last clean"),
        "alt_next": idx_of(contains="next cleaning") or idx_of(contains="next clean"),
    }
    if ix["vessel"] is None:
        return pd.DataFrame(columns=CONTACT_COLUMNS)

    body = rows[header_idx + 1:]
    # This sheet writes a/b/yyyy month-first while the Fleet Schedule writes it day-first,
    # so the convention is inferred here separately.
    contact_dayfirst = infer_dayfirst(
        [_cell(r, ix["alt_dd"]) for r in body] + [_cell(r, ix["alt_next"]) for r in body]
    )

    out = []
    for row in body:
        name = display(_cell(row, ix["vessel"]))
        if not name:
            continue
        matched = canonical.get(normalise(name)) or fuzzy_match(name, canonical)
        cycle_raw = _cell(row, ix["cycle"])
        try:
            cycle = int(float(cycle_raw))
        except (TypeError, ValueError):
            cycle = None
        out.append({
            "key": normalise(matched or name),
            "vessel": matched or name,
            "code": _text(row, ix["code"]),
            "call_cycle_days": cycle,
            "vessel_email": _text(row, ix["email"]),
            "agency": _text(row, ix["agency"]),
            "busan_terminal": _text(row, ix["busan"]),
            "singapore_terminal": _text(row, ix["singapore"]),
            "other_ports": _text(row, ix["other"]),
            "alt_drydock": coerce_date(_cell(row, ix["alt_dd"]), dayfirst=contact_dayfirst),
            "alt_next_drydock": coerce_date(_cell(row, ix["alt_next"]), dayfirst=contact_dayfirst),
        })
    return pd.DataFrame(out, columns=CONTACT_COLUMNS).drop_duplicates(subset="key", keep="first")


DEFAULT_PRICES: dict[str, Any] = {
    "inspection": 1500.0,
    "propeller": 4500.0,
    "second_mobilisation": 3000.0,
    "hull_bands": [(100.0, 150.0, 16000.0), (151.0, 200.0, 20000.0), (201.0, 250.0, 20000.0),
                   (251.0, 300.0, 22000.0), (301.0, 350.0, 24000.0)],
}


def _read_prices(book, sheet_name) -> dict[str, Any]:
    prices: dict[str, Any] = dict(DEFAULT_PRICES)
    prices["hull_bands"] = list(DEFAULT_PRICES["hull_bands"])
    if not sheet_name:
        return prices

    section = ""
    bands: list[tuple[float, float, float]] = []
    for row in book[sheet_name].iter_rows(values_only=True):
        text = " ".join(str(c) for c in row if c is not None)
        low = text.lower()
        if "inspection" in low and "service" in low:
            section = "inspection"
        elif "hull cleaning" in low and "service" in low:
            section = "hull"
        elif "propeller" in low and "service" in low:
            section = "propeller"
        elif "mobilization" in low or "mobilisation" in low:
            section = "mob"

        numbers = [float(c) for c in row if isinstance(c, (int, float)) and not isinstance(c, bool)]
        band = re.search(r"(\d{2,4})\s*-\s*(\d{2,4})\s*m", low)
        if section == "hull" and band and numbers:
            bands.append((float(band.group(1)), float(band.group(2)), numbers[-1]))
        elif section == "inspection" and numbers and "all" in low:
            prices["inspection"] = numbers[-1]
        elif section == "propeller" and numbers and "all" in low:
            prices["propeller"] = numbers[-1]
        elif section == "mob" and numbers:
            prices["second_mobilisation"] = numbers[-1]
    if bands:
        prices["hull_bands"] = sorted(bands)
    return prices



INVOICE_COLUMNS = ["vessel", "key", "matched", "etd", "total_usd", "invoiced"]

_MONEY = re.compile(r"[-+]?[0-9,]+(?:\.[0-9]+)?")


def _money(value: Any) -> float | None:
    """'USD 16,000.00' -> 16000.0"""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    m = _MONEY.search(str(value))
    if not m:
        return None
    try:
        return float(m.group(0).replace(",", ""))
    except ValueError:
        return None


def _read_invoices(book, sheet_name, canonical: dict[str, str],
                   issues: list[LoadIssue] | None = None) -> pd.DataFrame:
    """Billed job legs.

    This sheet matters because a Busan North Port cleaning is often split across two port calls:
    the Order Log records only the first leg and marks it 'Partial', while the invoice carries the
    date the work actually finished. Without it a vessel can look eight months overdue when it was
    in fact cleaned three months ago.
    """
    if not sheet_name:
        if issues is not None:
            issues.append(LoadIssue(
                "warning", "", "", "invoices",
                "No invoice sheet was found, so a cleaning split over two port calls will be "
                "dated from its first leg and the vessel will look more overdue than it is.",
            ))
        return pd.DataFrame(columns=INVOICE_COLUMNS)
    rows = [list(r) for r in book[sheet_name].iter_rows(values_only=True)]
    if not rows:
        return pd.DataFrame(columns=INVOICE_COLUMNS)
    headers = [_norm_header(c) for c in rows[0]]

    def idx_of(*names, contains=None):
        for name in names:
            if name in headers:
                return headers.index(name)
        if contains:
            for i, h in enumerate(headers):
                if contains in h:
                    return i
        return None

    ix = {
        "vessel": idx_of("vessel", "vessel name"),
        # The free-text 'ETA - ETD' column is unparseable; the '[for filtering]' one is a real date.
        "etd": idx_of(contains="for filtering") or idx_of("etd"),
        "total": idx_of(contains="total price"),
        "invoiced": idx_of(contains="invoiced"),
    }
    if ix["vessel"] is None or ix["etd"] is None:
        return pd.DataFrame(columns=INVOICE_COLUMNS)

    body = rows[1:]
    dayfirst = infer_dayfirst(_cell(r, ix["etd"]) for r in body)
    out = []
    for row in body:
        name = display(_cell(row, ix["vessel"]))
        if not name:
            continue
        matched = canonical.get(normalise(name)) or fuzzy_match(name, canonical)
        out.append({
            "vessel": matched or name,
            "key": normalise(matched or name),
            "matched": matched is not None,
            "etd": coerce_date(_cell(row, ix["etd"]), dayfirst=dayfirst),
            "total_usd": _money(_cell(row, ix["total"])),
            "invoiced": _text(row, ix["invoiced"]),
        })
    return pd.DataFrame(out, columns=INVOICE_COLUMNS)


def _read_job_checksum(book, sheet_name) -> dict[str, int]:
    """The summary sheet states how many jobs the dump should contain, per scope.

    Used to warn when an uploaded dump has a truncated Order Log, which would silently make
    vessels look overdue.
    """
    if not sheet_name:
        return {}
    out: dict[str, int] = {}
    for row in book[sheet_name].iter_rows(values_only=True):
        cells = [c for c in row if c is not None]
        if len(cells) < 2:
            continue
        label = str(cells[-2]).strip().upper()
        value = cells[-1]
        if label and isinstance(value, (int, float)) and not isinstance(value, bool)                 and re.fullmatch(r"[A-Z]{2,6}", label):
            out[label] = int(value)
    return out


def _read_as_at(book, sheet_name) -> _dt.date | None:
    """Pick up the 'Data as at <date>' stamp the workbook carries on its summary sheet."""
    if not sheet_name:
        return None
    for row in book[sheet_name].iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str) and "as at" in cell.lower():
                m = re.search(r"as at\s+(.+?)\s*$", cell, flags=re.I)
                if m:
                    try:
                        return pd.to_datetime(m.group(1).strip(), dayfirst=True).date()
                    except Exception:
                        pass
            if isinstance(cell, _dt.datetime):
                return cell.date()
    return None


# ---------------------------------------------------------------------------- entry point


def load_workbook(path: str | Path) -> Workbook:
    """Parse a fleet workbook. Raises ValueError if it is not a recognisable fleet file."""
    path = Path(path)
    book = openpyxl.load_workbook(path, data_only=True)
    issues: list[LoadIssue] = []
    sheet_map = resolve_sheets(book)

    if not sheet_map["fleet"]:
        raise ValueError(
            "No sheet in this file looks like a fleet schedule - none of them has a vessel-name "
            "column. Sheets present: " + ", ".join(book.sheetnames)
        )
    if sheet_map["fleet"] in (sheet_map["contacts"], sheet_map["orders"], sheet_map["invoices"]):
        raise ValueError(
            "The sheet " + repr(sheet_map["fleet"]) + " would have to serve as both the fleet "
            "schedule and another role, which means the fleet sheet is missing or named "
            "unexpectedly. Sheets present: " + ", ".join(book.sheetnames)
        )

    fleet, calls, intervals = _read_fleet(book, sheet_map["fleet"], issues)
    if fleet.empty:
        raise ValueError(
            "The sheet " + repr(sheet_map["fleet"]) + " was found but contained no vessel rows."
        )

    canonical = {row.key: row.vessel for row in fleet.itertuples()}
    orders = _read_orders(book, sheet_map["orders"], canonical, issues)
    invoices = _read_invoices(book, sheet_map["invoices"], canonical, issues)
    contacts = _read_contacts(book, sheet_map["contacts"], canonical)
    checksum = _read_job_checksum(book, sheet_map["asat"])
    prices = _read_prices(book, sheet_map["prices"])
    as_at = _read_as_at(book, sheet_map["asat"])
    book.close()

    return Workbook(
        path=path, fleet=fleet, calls=calls, orders=orders, invoices=invoices,
        contacts=contacts, prices=prices, as_at=as_at, issues=issues, sheet_map=sheet_map,
        job_checksum=checksum,
        cleaning_interval_months=intervals.get("cleaning", 12),
        inspection_interval_months=intervals.get("inspection", 6),
        digest=hashlib.sha256(path.read_bytes()).hexdigest()[:12],
    )
