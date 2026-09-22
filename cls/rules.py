"""The targeting rule engine.

The user's rule, in their own words:

    "I am looking at & suggesting vessels we have not cleaned in the last 8 months, are 2 years
     away from dry dock date, and have a next ETA in Busan/Singapore in October."

That becomes three gates, each independently adjustable from the sidebar:

    gate DD    - at least ``min_dd_age_years`` have passed since the vessel's last drydocking
    gate CLEAN - the last confirmed hull cleaning was more than ``min_months_since_clean`` ago
                 (never cleaned counts as maximally overdue)
    gate CALL  - the vessel has at least one scheduled ETA at a CLS hub inside the target window

Every vessel lands in exactly one status bucket, so the buckets always sum to the fleet size.
Ranking is a transparent 0-100 score whose three components are carried on the row, so the UI can
show exactly why a vessel sits where it does.

Two details of the source workbook make a naive reading of it wrong, and both are handled here:

* **Drydock dates.** Nine vessels have no usable "Last Dry Docking" value - it is either blank or
  dated in the future. For five of them the two drydock columns hold the *next special survey*
  (true drydock + 5 years) and the *intermediate survey* (+3 years), so swapping them does not
  recover the real date. The contact sheet's mislabelled "LAST CLEANING" column does: plus five
  years it reproduces the Fleet Schedule value to within a few days. It is used as a clearly
  labelled fallback, never silently.
* **Split cleanings.** A Busan North Port job is often spread over two port calls. The Order Log
  records only the first leg and marks it "Partial"; the Invoice sheet carries the date the work
  actually finished, billed as the LOA band plus a second mobilisation fee. Reading the Order Log
  alone makes a vessel look months overdue when it was cleaned recently.
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from typing import Any

import pandas as pd

from .dates import months_between, years_between
from .loader import Workbook

# ---------------------------------------------------------------------------- status taxonomy

PRIME = "Prime candidate"
WATCH = "Worth a look"
BOOKED = "Already booked"
RECENT = "Cleaned recently"
EARLY_DD = "Too soon after drydock"
NO_DD = "No drydock date"
NOT_CALLING = "Not calling in window"

STATUS_ORDER = [PRIME, WATCH, BOOKED, RECENT, EARLY_DD, NO_DD, NOT_CALLING]

# Palette slots from the data-viz reference palette. Every status is always rendered with its
# text label beside the swatch, so colour never carries the meaning on its own.
STATUS_COLOUR = {
    PRIME: "#0ca30c",        # status: good
    WATCH: "#fab219",        # status: warning
    BOOKED: "#2a78d6",       # categorical slot 1 (blue)
    RECENT: "#898781",       # muted ink
    EARLY_DD: "#ec835a",     # status: serious
    NO_DD: "#4a3aa7",        # categorical slot 7 (violet)
    NOT_CALLING: "#c3c2b7",  # baseline
}

STATUS_HELP = {
    PRIME: "Passes all three gates: overdue for cleaning, far enough past drydock, and calling a "
           "CLS hub inside the target window. These are the vessels to pitch.",
    WATCH: "Misses exactly one gate, and only just. Worth raising with the client if you want to "
           "widen the list - the 'Qualifies on' date says when it would pass.",
    BOOKED: "A cleaning is already confirmed on or before the end of the target window, so there "
            "is nothing to sell here.",
    RECENT: "Calling in the window and past drydock, but cleaned too recently to need it again.",
    EARLY_DD: "Calling in the window but not yet far enough past its last drydocking - the hull "
              "coating is still doing its job.",
    NO_DD: "Calling in the window but no last-drydocking date could be resolved from any sheet, "
           "so the rule cannot be applied. Fill the date in and this vessel gets classified.",
    NOT_CALLING: "No scheduled ETA at a CLS hub inside the target window.",
}

DD_FROM_FLEET = "Fleet Schedule"
DD_FROM_CONTACT = "Contact sheet (inferred)"
DD_UNKNOWN = "Not resolvable"

CLEANING_SCOPES = ("UWHC", "UWC", "HULL CLEAN", "CLEANING", "FOB")
INSPECTION_ONLY_SCOPES = ("UWI",)
CANCELLED = ("CANCEL", "ABORT")


# ---------------------------------------------------------------------------- settings


@dataclass
class Settings:
    today: _dt.date
    window_start: _dt.date
    window_end: _dt.date
    window_label: str = ""
    min_dd_age_years: float = 2.0
    min_months_since_clean: float = 8.0
    hubs: tuple[str, ...] = ("busan", "singapore")
    cleaning_interval_months: float = 12.0
    use_contact_drydock_fallback: bool = True
    measure_at_call: bool = False
    split_job_window_days: int = 90
    drydock_cleaning_clash_days: int = 30
    invoice_cleaning_min_usd: float = 8000.0
    near_miss_dd_slack_years: float = 0.5
    near_miss_clean_slack_months: float = 2.0
    drydock_soon_days: int = 120
    high_fouling_kg: float = 25.0

    def window_contains(self, day: _dt.date | None) -> bool:
        return day is not None and self.window_start <= day <= self.window_end


def month_window(year: int, month: int) -> tuple[_dt.date, _dt.date]:
    start = _dt.date(year, month, 1)
    end = _dt.date(year + (month == 12), (month % 12) + 1, 1) - _dt.timedelta(days=1)
    return start, end


# ---------------------------------------------------------------------------- helpers


def _clamp(value: float, low: float = 0.0, high: float = 1.0) -> float:
    return max(low, min(high, value))


def _is_cleaning(scope: str) -> bool:
    upper = (scope or "").upper()
    if any(tag in upper for tag in INSPECTION_ONLY_SCOPES) and not any(
        tag in upper for tag in ("UWHC", "UWC", "HULL")
    ):
        return False
    return any(tag in upper for tag in CLEANING_SCOPES) or not upper


def _is_cancelled(status: str) -> bool:
    return any(tag in (status or "").upper() for tag in CANCELLED)


def estimate_value(loa: float | None, prices: dict[str, Any]) -> float | None:
    """Hull-cleaning list price for a vessel of this length, from the workbook's price list.

    The bands in the workbook do not tile the number line - they run 100-150, 151-200, 201-250
    and so on, leaving a one-metre gap between each. A length that lands in a gap, or outside
    the whole table, takes the price of the band whose range it is closest to, never simply the
    last band in the list.
    """
    if loa is None or pd.isna(loa):
        return None
    bands = prices.get("hull_bands") or []
    if not bands:
        return None
    for low, high, price in bands:
        if low <= loa <= high:
            return float(price)

    def distance(band):
        low, high, _ = band
        return 0.0 if low <= loa <= high else min(abs(loa - low), abs(loa - high))

    return float(min(bands, key=distance)[2])


# ---------------------------------------------------------------------------- history


def _cleaning_history(wb: Workbook, settings: Settings) -> dict[str, dict[str, Any]]:
    """Per-vessel cleaning history from the Order Log, completed out by the Invoice sheet."""
    history: dict[str, dict[str, Any]] = {
        row.key: {"jobs": [], "fouling": [], "split": [], "invoice_dates": [],
                  "last_inspection": None}
        for row in wb.fleet.itertuples()
    }

    invoices_by_key: dict[str, list[_dt.date]] = {}
    billed_by_key: dict[str, list[tuple[_dt.date, float | None]]] = {}
    if wb.invoices is not None and not wb.invoices.empty:
        for row in wb.invoices.itertuples():
            if row.etd is not None and row.matched:
                invoices_by_key.setdefault(row.key, []).append(row.etd)
                billed_by_key.setdefault(row.key, []).append((row.etd, row.total_usd))
    for bucket_key, dates in invoices_by_key.items():
        if bucket_key in history:
            history[bucket_key]["invoice_dates"] = sorted(dates)

    if wb.orders is not None and not wb.orders.empty:
        for row in wb.orders.itertuples():
            bucket = history.get(row.key)
            if bucket is None or row.eta is None or _is_cancelled(row.status):
                continue
            if not _is_cleaning(row.scope):
                prev = bucket["last_inspection"]
                if prev is None or row.eta > prev:
                    bucket["last_inspection"] = row.eta
            else:
                # A job that ran over two port calls finishes on the date it was billed. Only a
                # job the order log actually marks unfinished is extended, and only up to the
                # vessel's next job - otherwise an ordinary invoice for a completed cleaning, or
                # the billing for a later separate campaign, gets folded in and the vessel looks
                # cleaner than it is.
                is_partial = "PARTIAL" in (row.status or "").upper()
                next_job = min(
                    (o.eta for o in wb.orders.itertuples()
                     if o.key == row.key and o.eta is not None and o.eta > row.eta),
                    default=None,
                )
                horizon = row.eta + _dt.timedelta(days=settings.split_job_window_days)
                if next_job is not None:
                    horizon = min(horizon, next_job - _dt.timedelta(days=1))
                later = [
                    etd for etd in invoices_by_key.get(row.key, [])
                    if is_partial and row.eta <= etd <= horizon
                ]
                finished = max(later) if later else row.eta
                bucket["jobs"].append({
                    "started": row.eta, "finished": finished, "status": row.status,
                    "hub": row.hub, "scope": row.scope,
                    "split": finished > row.eta + _dt.timedelta(days=7),
                })
                if finished > row.eta + _dt.timedelta(days=7):
                    bucket["split"].append((row.eta, finished))
            if row.fouling is not None and not pd.isna(row.fouling):
                bucket["fouling"].append(float(row.fouling))

    # A hull cleaning that was billed but never written to the order log is still a cleaning.
    # Only substantial invoices count, so an inspection or a propeller job is not mistaken for
    # one, and only those that do not already belong to a recorded job.
    for key, bucket in history.items():
        accounted = {job["started"] for job in bucket["jobs"]}
        accounted |= {job["finished"] for job in bucket["jobs"]}
        for when, amount in billed_by_key.get(key, []):
            if amount is None or amount < settings.invoice_cleaning_min_usd:
                continue
            if any(abs((when - seen).days) <= settings.split_job_window_days
                   for seen in accounted):
                continue
            bucket["jobs"].append({
                "started": when, "finished": when, "status": "Invoiced",
                "hub": "", "scope": "invoice only", "split": False,
            })
            accounted.add(when)
    return history


def _resolve_drydock(vessel, contact: pd.Series | None, settings: Settings,
                     known_cleanings: set[_dt.date] | None = None):
    """Pick the best available last-drydocking date, and say where it came from.

    The contact sheet's column is mislabelled "LAST CLEANING" and for most vessels really does
    hold the drydocking - but not for all of them, so it is only trusted when the workbook
    corroborates it. Two checks:

    * the same row's "next cleaning" value must sit roughly five years later, which is what a
      special-survey interval looks like and is what made this column identifiable; and
    * the date must not coincide with a cleaning the workbook records elsewhere.

    Without the second check SAWASDEE CAPELLA is given a "drydocking" of 2026-01-06 - the day
    after an Order Log cleaning on 2026-01-05, invoiced on 2026-01-09 - and is then excluded for
    being too soon after a drydocking that never happened.
    """
    fleet_last, fleet_next = vessel.last_drydock, vessel.next_drydock
    if fleet_last is not None and fleet_last <= settings.today:
        return fleet_last, fleet_next, DD_FROM_FLEET

    alt = alt_next = None
    if contact is not None:
        alt = contact.get("alt_drydock")
        alt_next = contact.get("alt_next_drydock")
        if alt is not None and pd.isna(alt):
            alt = None
        if alt_next is not None and pd.isna(alt_next):
            alt_next = None

    usable = (
        settings.use_contact_drydock_fallback
        and alt is not None
        and alt <= settings.today
        and _looks_like_a_survey_interval(alt, alt_next)
        and not _clashes_with_a_cleaning(alt, known_cleanings or set(),
                                         settings.drydock_cleaning_clash_days)
    )
    if usable:
        return alt, fleet_next or alt_next, DD_FROM_CONTACT
    return None, fleet_next or alt_next, DD_UNKNOWN


def _looks_like_a_survey_interval(alt: _dt.date, alt_next: _dt.date | None) -> bool:
    if alt_next is None:
        return False
    gap = years_between(alt, alt_next)
    return gap is not None and 4.0 <= gap <= 6.0


def _clashes_with_a_cleaning(alt: _dt.date, cleanings: set[_dt.date], window_days: int) -> bool:
    return any(abs((alt - when).days) <= window_days for when in cleanings)


# ---------------------------------------------------------------------------- main evaluation


def evaluate(wb: Workbook, settings: Settings) -> pd.DataFrame:
    """Score and classify every vessel. Returns one row per vessel."""
    history = _cleaning_history(wb, settings)
    contacts = (
        wb.contacts.set_index("key") if wb.contacts is not None and not wb.contacts.empty
        else pd.DataFrame()
    )
    calls = wb.calls
    today = settings.today
    interval = max(float(settings.cleaning_interval_months), 1.0)

    rows: list[dict[str, Any]] = []
    for v in wb.fleet.itertuples():
        contact = contacts.loc[v.key] if v.key in getattr(contacts, "index", []) else None
        if isinstance(contact, pd.DataFrame):          # duplicate keys, take the first
            contact = contact.iloc[0]

        # -- port calls in the target window ----------------------------------------
        mine = calls[calls["key"] == v.key] if not calls.empty else calls
        window_calls = []
        for c in mine.itertuples():
            if not settings.window_contains(c.eta):
                continue
            if settings.hubs and not any(h in (c.hub_key or "") for h in settings.hubs):
                continue
            stay_days = (c.etd - c.eta).days if (c.etd and c.eta and c.etd >= c.eta) else None
            window_calls.append({
                "eta": c.eta, "etd": c.etd, "hub": (c.hub or "").strip(),
                "terminal": c.terminal, "stay_days": stay_days,
            })
        window_calls.sort(key=lambda c: c["eta"])
        gate_call = bool(window_calls)

        # Ages can be measured today, or on the day the vessel actually berths - a vessel
        # calling on the 29th has four more weeks of coating age than one calling on the 1st.
        as_of = window_calls[0]["eta"] if (settings.measure_at_call and window_calls) else today

        # -- drydock ----------------------------------------------------------------
        # Every cleaning date the workbook knows for this vessel, so the contact-sheet fallback
        # can refuse a value that is really a cleaning.
        bucket_preview = history.get(v.key, {})
        known_cleanings = {job["started"] for job in bucket_preview.get("jobs", [])}
        known_cleanings |= set(bucket_preview.get("invoice_dates", []))
        if v.clean_confirmed:
            known_cleanings.add(v.clean_confirmed)

        last_dd, next_dd, dd_source = _resolve_drydock(v, contact, settings, known_cleanings)
        dd_age = years_between(last_dd, as_of)
        dd_known = dd_age is not None
        gate_dd = bool(dd_known and dd_age >= settings.min_dd_age_years)
        next_dd_days = (next_dd - today).days if next_dd else None

        # -- cleaning history -------------------------------------------------------
        bucket = history.get(v.key, {"jobs": [], "fouling": [], "split": [],
                                     "invoice_dates": [], "last_inspection": None})
        job_dates = [job["finished"] for job in bucket["jobs"]]
        confirmed = v.clean_confirmed

        past_dates = [d for d in job_dates if d <= today]
        if confirmed and confirmed <= today:
            past_dates.append(confirmed)
        last_clean = max(past_dates) if past_dates else None

        future_dates = [d for d in job_dates if d > today]
        if confirmed and confirmed > today:
            future_dates.append(confirmed)
        future_dates += [d for d in bucket["invoice_dates"] if d > today and d not in job_dates]
        next_booked = min(future_dates) if future_dates else None

        # Name the sheets that actually produced the date being shown, not every sheet that
        # mentions the vessel.
        sources: list[str] = []
        if last_clean is not None:
            if confirmed == last_clean:
                sources.append("Fleet Schedule")
            for job in bucket["jobs"]:
                if job["finished"] != last_clean:
                    continue
                if job["scope"] == "invoice only":
                    sources.append("Invoice")
                elif job["split"]:
                    sources.append("Order Log")
                    sources.append("Invoice")
                else:
                    sources.append("Order Log")
        clean_source = " + ".join(dict.fromkeys(sources))

        months_since = months_between(last_clean, as_of)
        never_cleaned = last_clean is None
        gate_clean = never_cleaned or months_since >= settings.min_months_since_clean

        # -- status -----------------------------------------------------------------
        booked_in_window = next_booked is not None and next_booked <= settings.window_end
        if booked_in_window:
            status = BOOKED
        elif not gate_call:
            status = NOT_CALLING
        elif not dd_known:
            status = NO_DD
        elif not gate_dd:
            status = EARLY_DD
        elif not gate_clean:
            status = RECENT
        else:
            status = PRIME

        # -- when would a failing vessel qualify? -----------------------------------
        qualifies_on = None
        if status in (EARLY_DD, RECENT):
            needed = []
            if not gate_dd and last_dd is not None:
                needed.append(last_dd + _dt.timedelta(days=settings.min_dd_age_years * 365.25))
            if not gate_clean and last_clean is not None:
                needed.append(
                    last_clean + _dt.timedelta(days=settings.min_months_since_clean * 30.4375)
                )
            qualifies_on = max(needed) if needed else None

        near_miss_reason = ""
        if status == EARLY_DD and dd_age is not None and gate_clean and \
                dd_age >= settings.min_dd_age_years - settings.near_miss_dd_slack_years:
            near_miss_reason = "drydock gate missed by %.1f months" % (
                (settings.min_dd_age_years - dd_age) * 12)
        elif status == RECENT and months_since is not None and \
                months_since >= settings.min_months_since_clean - settings.near_miss_clean_slack_months:
            near_miss_reason = "cleaning gate missed by %.1f months" % (
                settings.min_months_since_clean - months_since)
        if near_miss_reason:
            status = WATCH

        # -- score ------------------------------------------------------------------
        if never_cleaned:
            clean_pts = 50.0
        else:
            low = settings.min_months_since_clean
            high = max(interval * 1.5, low + 1.0)
            clean_pts = 50.0 * _clamp((months_since - low) / (high - low))

        if dd_age is None:
            dd_pts = 0.0
        else:
            low = settings.min_dd_age_years
            high = max(low + 3.0, 5.0)
            dd_pts = 30.0 * _clamp((dd_age - low) / (high - low))

        n_calls = len(window_calls)
        access_pts = 14.0 * _clamp(min(n_calls, 3) / 3.0)
        longest_stay = max((c["stay_days"] for c in window_calls if c["stay_days"] is not None),
                           default=None)
        if longest_stay is None:
            access_pts += 2.0 if n_calls else 0.0
        else:
            access_pts += 6.0 if longest_stay >= 1 else 4.0
        score = round(clean_pts + dd_pts + access_pts, 1)

        # -- flags ------------------------------------------------------------------
        alerts: list[str] = []
        if dd_source == DD_FROM_CONTACT:
            alerts.append(
                "Drydock date taken from the contact sheet - the Fleet Schedule value is "
                + ("blank" if v.last_drydock is None else "dated in the future")
            )
        elif dd_source == DD_UNKNOWN and v.last_drydock is not None:
            alerts.append("Fleet Schedule drydock date is in the future and no fallback was found")
        if next_booked:
            alerts.append("Cleaning already booked for " + next_booked.strftime("%d %b %Y"))
        if next_dd_days is not None and 0 <= next_dd_days <= settings.drydock_soon_days:
            alerts.append("Next drydock in " + str(next_dd_days) + " days")
        if next_dd_days is not None and next_dd_days < 0:
            alerts.append("Next drydock date has passed (" + next_dd.strftime("%d %b %Y") + ")")
        for started, finished in bucket["split"]:
            alerts.append(
                "Last job ran over two calls: started %s, completed %s"
                % (started.strftime("%d %b"), finished.strftime("%d %b %Y"))
            )
        peak_fouling = max(bucket["fouling"]) if bucket["fouling"] else None
        if peak_fouling is not None and peak_fouling >= settings.high_fouling_kg:
            alerts.append("Heavy fouling on a past job (%g)" % peak_fouling)
        for note in (v.recommend_action, v.remarks):
            if note:
                alerts.append("Workbook note: " + note)

        record = {
            "vessel": v.vessel,
            "key": v.key,
            "imo": v.imo,
            "loa": v.loa,
            "vessel_type": v.vessel_type,
            "pic": v.pic,
            "pic_email": v.pic_email,
            "group_email": v.group_email,
            "status": status,
            "score": score,
            "score_cleaning": round(clean_pts, 1),
            "score_drydock": round(dd_pts, 1),
            "score_access": round(access_pts, 1),
            "last_drydock": last_dd,
            "next_drydock": next_dd,
            "drydock_source": dd_source,
            "sheet_last_drydock": v.last_drydock,
            "dd_age_years": None if dd_age is None else round(dd_age, 2),
            "next_drydock_days": next_dd_days,
            "last_clean": last_clean,
            "last_clean_source": clean_source,
            "months_since_clean": None if months_since is None else round(months_since, 1),
            "never_cleaned": never_cleaned,
            "next_booked_clean": next_booked,
            "clean_next_rec": v.clean_next_rec,
            "last_inspection": bucket["last_inspection"] or v.uwi_confirmed,
            "job_count": len(bucket["jobs"]),
            "peak_fouling": peak_fouling,
            "measured_at": as_of,
            "gate_dd": gate_dd,
            "gate_clean": gate_clean,
            "gate_call": gate_call,
            "gates_passed": int(gate_dd) + int(gate_clean) + int(gate_call),
            "qualifies_on": qualifies_on,
            "near_miss_reason": near_miss_reason,
            "calls_in_window": n_calls,
            "window_calls": window_calls,
            "first_eta": window_calls[0]["eta"] if window_calls else None,
            "first_etd": window_calls[0]["etd"] if window_calls else None,
            "hubs": " / ".join(dict.fromkeys(c["hub"] for c in window_calls if c["hub"])),
            "terminals": " / ".join(dict.fromkeys(c["terminal"] for c in window_calls if c["terminal"])),
            "longest_stay_days": longest_stay,
            "est_value_usd": estimate_value(v.loa, wb.prices),
            "alerts": alerts,
            "remarks": v.remarks,
            "recommend_action": v.recommend_action,
        }
        # Built from the plain dict, before pandas coerces None into NaN/NaT.
        record["why"] = build_why(record, settings)
        rows.append(record)

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["status_rank"] = df["status"].map({s: i for i, s in enumerate(STATUS_ORDER)})
    return df.sort_values(["status_rank", "score"], ascending=[True, False]).reset_index(drop=True)


# ---------------------------------------------------------------------------- explanation


def build_why(row: pd.Series | dict, settings: Settings) -> str:
    """A one-line, human-readable justification for the vessel's position.

    Accepts either a raw record dict or a DataFrame row; missing values arrive as ``None`` from
    the former and as ``NaN``/``NaT`` from the latter, so every read goes through ``get``.
    """
    def get(name):
        value = row.get(name)
        try:
            if value is None or (not isinstance(value, (list, tuple, str)) and pd.isna(value)):
                return None
        except (TypeError, ValueError):
            pass
        return value

    parts: list[str] = []

    if get("never_cleaned"):
        parts.append("never cleaned by CLS")
    elif get("last_clean") is not None:
        parts.append("last cleaned %s (%.0f months ago)" % (
            get("last_clean").strftime("%d %b %Y"), get("months_since_clean") or 0))

    dd_age, last_dd = get("dd_age_years"), get("last_drydock")
    if dd_age is None or last_dd is None:
        parts.append("no drydock date on file")
    else:
        suffix = "*" if get("drydock_source") == DD_FROM_CONTACT else ""
        parts.append("%.1f yrs since drydock (%s%s)" % (dd_age, last_dd.strftime("%b %Y"), suffix))

    window_calls = get("window_calls") or []
    label = settings.window_label or "the window"
    if window_calls:
        when = ", ".join(
            "%s %s" % (c["hub"] or "?", c["eta"].strftime("%d %b").lstrip("0"))
            for c in window_calls[:3]
        )
        parts.append("%d call%s in %s (%s)" % (
            len(window_calls), "" if len(window_calls) == 1 else "s", label, when))
    else:
        parts.append("no CLS-hub call in " + label)

    if get("near_miss_reason"):
        parts.append(get("near_miss_reason"))
    if get("qualifies_on") is not None:
        parts.append("qualifies " + get("qualifies_on").strftime("%d %b %Y"))
    return " · ".join(parts)


# ---------------------------------------------------------------------------- roll-ups


def summarise(df: pd.DataFrame) -> dict[str, Any]:
    if df.empty:
        return {s: 0 for s in STATUS_ORDER} | {"fleet": 0, "pipeline_usd": 0.0,
                                               "never_cleaned": 0, "calling": 0}
    counts = df["status"].value_counts().to_dict()
    prime = df[df["status"] == PRIME]
    out: dict[str, Any] = {s: int(counts.get(s, 0)) for s in STATUS_ORDER}
    out["fleet"] = int(len(df))
    out["pipeline_usd"] = float(prime["est_value_usd"].fillna(0).sum())
    out["never_cleaned"] = int(df["never_cleaned"].sum())
    out["calling"] = int((df["calls_in_window"] > 0).sum())
    return out


def gate_funnel(df: pd.DataFrame, settings: Settings) -> pd.DataFrame:
    """How many vessels survive each gate, in the order the user describes them."""
    if df.empty:
        return pd.DataFrame(columns=["stage", "vessels"])
    calling = df[df["gate_call"]]
    stages = [
        ("Whole fleet", len(df)),
        ("Calling a CLS hub in " + (settings.window_label or "window"), len(calling)),
        ("... and %.1f+ yrs since drydock" % settings.min_dd_age_years,
         int(calling["gate_dd"].sum())),
        ("... and not cleaned in %.0f months" % settings.min_months_since_clean,
         int((calling["gate_dd"] & calling["gate_clean"]).sum())),
        ("... and not already booked", int((df["status"] == PRIME).sum())),
    ]
    return pd.DataFrame(stages, columns=["stage", "vessels"])
