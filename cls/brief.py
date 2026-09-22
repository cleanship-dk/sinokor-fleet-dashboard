"""Turn selected vessels into something that can be pasted straight into an email to the client."""
from __future__ import annotations

import datetime as _dt
import io
from typing import Any, Iterable

import pandas as pd

from .rules import Settings


def _date(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return value.strftime("%d %b %Y")


def _calls_line(row) -> str:
    calls = row.get("window_calls") or []
    if not calls:
        return "no CLS-hub call scheduled in the window"
    parts = []
    for call in calls:
        where = " ".join(p for p in (call.get("hub"), call.get("terminal")) if p)
        eta, etd = call.get("eta"), call.get("etd")
        if eta and etd and etd != eta:
            # Keep both months when the call straddles one, or a 30 Oct - 3 Nov call reads "30-03 Nov".
            left = eta.strftime("%d") if eta.month == etd.month else eta.strftime("%d %b")
            when = left + "-" + etd.strftime("%d %b")
        elif eta:
            when = eta.strftime("%d %b")
        else:
            when = "date TBC"
        parts.append((where + ", " + when).strip(", "))
    return "  |  ".join(parts)


def _cleaning_line(row) -> str:
    if row.get("never_cleaned"):
        return "none on record"
    months = row.get("months_since_clean")
    suffix = "" if months is None or pd.isna(months) else " (%.0f months ago)" % months
    return _date(row.get("last_clean")) + suffix


def _recommendation(row, status: str | None = None) -> str:
    calls = row.get("window_calls") or []
    if not calls:
        return "no call scheduled at a CLS hub in this period - keep under review."
    if status and status != "Prime candidate":
        reason = {
            "Too soon after drydock": "not yet far enough past its last drydocking",
            "Cleaned recently": "cleaned too recently to need it again",
            "Already booked": "a cleaning is already booked",
            "No drydock date": "we have no drydocking date on file for it",
            "Worth a look": "just short of the cleaning criteria",
        }.get(status, "outside the usual criteria")
        return "raised for discussion only - " + reason + "."
    if len(calls) == 1:
        return "schedule underwater hull cleaning on this call."
    return "schedule underwater hull cleaning on either call, whichever suits the berth plan."


# Alerts that belong in front of the client, as opposed to internal parsing notes.
_CLIENT_ALERT_PREFIXES = (
    "Next drydock in", "Next drydock date has passed", "Cleaning already booked",
    "Last job ran over two calls", "Heavy fouling",
)


def _client_alerts(row) -> list[str]:
    alerts = row.get("alerts")
    if not isinstance(alerts, (list, tuple)):
        return []
    return [a for a in alerts if a.startswith(_CLIENT_ALERT_PREFIXES)]


def _months(value) -> str:
    months = int(round(float(value or 12)))
    return "12 months" if months == 12 else "%d months" % months


def _years(value) -> str:
    years = float(value or 0)
    if years <= 0:
        return "any time"
    if abs(years - round(years)) < 0.05:
        return "%d year%s" % (round(years), "" if round(years) == 1 else "s")
    return "%.2f years" % years


def build_email(rows: Iterable[dict], settings: Settings,
                client: str = "Sinokor", sender: str = "") -> str:
    """A ready-to-send plain-text brief covering the given vessels."""
    rows = list(rows)
    window = settings.window_label or (
        settings.window_start.strftime("%d %b") + " - " + settings.window_end.strftime("%d %b %Y")
    )
    if not rows:
        return "No vessels selected for " + window + "."

    # The brief is built from whatever the user starred, which need not all be candidates.
    all_due = all(row.get("status") == "Prime candidate" for row in rows)
    lines = [
        "Subject: Hull cleaning candidates for " + window
        + " - " + str(len(rows)) + (" vessel" if len(rows) == 1 else " vessels"),
        "",
        "Dear " + client + " team,",
        "",
        "Ahead of the " + window + " port rotation we have reviewed the fleet against the",
        "cleaning criteria we have been working to: a hull cleaning roughly every "
        + _months(settings.cleaning_interval_months) + ", from",
        _years(settings.min_dd_age_years) + " after drydocking.",
        "",
    ]
    if all_due:
        lines.append("The vessels below are due for a hull cleaning and call a CLS service hub "
                     "in the period:")
    else:
        lines.append("The vessels below are the ones we would like to discuss for the period.")
        lines.append("Where a vessel is not yet due, or has no call scheduled, that is noted "
                     "against it:")
    lines.append("")

    for i, row in enumerate(rows, start=1):
        loa = row.get("loa")
        status = row.get("status")
        head = "%d. %s" % (i, row.get("vessel", ""))
        ident = []
        if row.get("imo") and not pd.isna(row.get("imo")):
            ident.append("IMO %d" % int(row["imo"]))
        if loa is not None and not pd.isna(loa):
            ident.append("%.0f m" % float(loa))
        if ident:
            head += "  (" + ", ".join(ident) + ")"
        lines.append(head)
        lines.append("     Last drydocking  : " + (_date(row.get("last_drydock")) or "not on file")
                     + _age_suffix(row.get("dd_age_years")))
        lines.append("     Last CLS cleaning: " + _cleaning_line(row))
        lines.append("     Port calls       : " + _calls_line(row))
        if row.get("pic"):
            contact = row["pic"]
            if row.get("pic_email"):
                contact += " (" + row["pic_email"] + ")"
            lines.append("     Vessel PIC       : " + contact)
        lines.append("     Recommendation   : " + _recommendation(row, status))
        # Anything the dashboard flagged is said out loud rather than left for the client to
        # find - a vessel due to drydock next month should not be pitched a cleaning in silence.
        for alert in _client_alerts(row):
            lines.append("     Please note      : " + alert)
        note = (row.get("note") or "").strip()
        if note:
            lines.append("     Note             : " + note)
        lines.append("")

    lines += [
        "Please confirm which vessels you would like us to book so we can hold the slots with",
        "the terminal and dive team.",
        "",
        "Kind regards,",
        sender or "CLS",
    ]
    return "\n".join(lines)


def _age_suffix(years) -> str:
    if years is None or pd.isna(years):
        return ""
    return "  (%.1f years ago)" % float(years)


BRIEF_COLUMNS = [
    ("vessel", "Vessel"),
    ("imo", "IMO"),
    ("loa", "LOA (m)"),
    ("status", "Status"),
    ("score", "Priority score"),
    ("last_drydock", "Last drydocking"),
    ("dd_age_years", "Years since drydock"),
    ("next_drydock", "Next drydock"),
    ("last_clean", "Last CLS cleaning"),
    ("months_since_clean", "Months since cleaning"),
    ("hubs", "Hub"),
    ("terminals", "Terminal"),
    ("first_eta", "First ETA in window"),
    ("first_etd", "First ETD in window"),
    ("calls_in_window", "Calls in window"),
    ("est_value_usd", "Est. value (USD)"),
    ("pic", "Vessel PIC"),
    ("pic_email", "PIC email"),
    ("group_email", "Group email"),
    ("why", "Why"),
    ("flag_text", "Flags"),
]


def build_table(df: pd.DataFrame) -> pd.DataFrame:
    """A flat, client-presentable table."""
    out = df.copy()
    if "alerts" in out.columns:
        out["flag_text"] = out["alerts"].apply(
            lambda f: "; ".join(f) if isinstance(f, (list, tuple)) else ""
        )
    else:
        out["flag_text"] = ""
    cols = [(src, label) for src, label in BRIEF_COLUMNS if src in out.columns]
    out = out[[src for src, _ in cols]].rename(columns=dict(cols))
    return out


def to_excel(tables: dict[str, pd.DataFrame]) -> bytes:
    """Multi-sheet xlsx export. Sheet name -> dataframe."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, table in tables.items():
            safe = str(name)[:31] or "Sheet1"
            frame = table.copy()
            for column in frame.columns:
                if frame[column].map(lambda v: isinstance(v, (list, tuple, dict))).any():
                    frame[column] = frame[column].astype(str)
            frame.to_excel(writer, sheet_name=safe, index=False)
            sheet = writer.sheets[safe]
            for idx, column in enumerate(frame.columns, start=1):
                width = max(len(str(column)), *(frame[column].astype(str).str.len().tolist() or [0]))
                sheet.column_dimensions[sheet.cell(row=1, column=idx).column_letter].width = \
                    min(max(width + 2, 10), 60)
    return buffer.getvalue()


def to_csv(table: pd.DataFrame) -> bytes:
    return table.to_csv(index=False).encode("utf-8-sig")


def filename(prefix: str, settings: Settings, extension: str) -> str:
    window = (settings.window_label or settings.window_start.strftime("%Y-%m")).replace(" ", "-")
    stamp = _dt.date.today().strftime("%Y%m%d")
    return "%s_%s_%s.%s" % (prefix, window, stamp, extension)
