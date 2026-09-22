"""CLS fleet cleaning targeting dashboard.

Which Sinokor vessels should we pitch a hull cleaning to, for a given month?
"""
from __future__ import annotations

import datetime as _dt
from pathlib import Path

import pandas as pd
import streamlit as st

from cls import brief as brief_mod
from cls import charts, store
from cls.loader import load_workbook
from cls.names import normalise
from cls.rules import (
    DD_FROM_CONTACT, DD_UNKNOWN, PRIME, STATUS_COLOUR, STATUS_HELP, STATUS_ORDER, WATCH,
    Settings, evaluate, gate_funnel, month_window, summarise,
)

st.set_page_config(page_title="CLS Fleet Targeting", page_icon="🚢", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1500px; }
      [data-testid="stMetricValue"] { font-size: 1.65rem; }
      [data-testid="stMetricLabel"] { font-size: 0.78rem; letter-spacing: .02em; }
      .pill { display:inline-block; padding:2px 9px; border-radius:11px; font-size:0.74rem;
              font-weight:600; margin-right:5px; border:1px solid rgba(0,0,0,.08); }
      .hl-card { border:1px solid #e6e5e0; border-left:4px solid #2a78d6; border-radius:8px;
                 padding:11px 14px; margin-bottom:9px; background:#fcfcfb; }
      .hl-card h4 { margin:0 0 4px 0; font-size:0.98rem; }
      .hl-card p  { margin:0; font-size:0.83rem; color:#52514e; line-height:1.5; }
      .muted { color:#898781; font-size:0.82rem; }
      div[data-testid="stDataFrame"] div[role="gridcell"] { font-size: 0.86rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================================ data plumbing


@st.cache_data(show_spinner="Reading workbook...")
def _load(path_str: str, mtime: float, size: int):
    """Cached on (path, mtime, size) so re-uploading the same file is instant."""
    return load_workbook(path_str)


def load_cached(path: Path):
    stat = Path(path).stat()
    return _load(str(path), stat.st_mtime, stat.st_size)


def status_pill(status: str) -> str:
    colour = STATUS_COLOUR.get(status, "#898781")
    return (
        '<span class="pill" style="background:%s22;color:#0b0b0b;border-color:%s55;">'
        '<span style="color:%s;">&#9679;</span> %s</span>' % (colour, colour, colour, status)
    )


def fmt_date(value) -> str:
    if value is None:
        return "—"
    try:
        if pd.isna(value):
            return "—"
    except (TypeError, ValueError):
        pass
    return value.strftime("%d %b %Y")


def fmt_money(value) -> str:
    if value is None or pd.isna(value):
        return "—"
    return "${:,.0f}".format(value)


# ============================================================================ sidebar

files = store.list_data_files()
if not files:
    st.title("🚢 CLS Fleet Targeting")
    st.warning("No workbook found yet. Upload a fleet Excel file to get started.")
    first = st.file_uploader("Fleet workbook (.xlsx)", type=["xlsx", "xlsm", "xls"])
    if first is not None:
        saved, note = store.save_upload(first.getvalue(), first.name)
        try:
            load_workbook(saved)
        except Exception as exc:                                    # noqa: BLE001
            store.delete_data_file(saved)
            st.error("That file could not be read: %s" % exc)
            st.stop()
        st.success("Loaded %s." % saved.name)
        st.rerun()
    st.stop()

if "watchlist" not in st.session_state:
    st.session_state.watchlist = store.load_watchlist()

with st.sidebar:
    st.markdown("### 📂 Data source")
    labels = [f.label for f in files]
    chosen_label = st.selectbox(
        "Workbook", labels, index=0,
        help="The newest upload is selected by default. Add a new dump on the "
             "'Data & updates' tab.",
    )
    active = files[labels.index(chosen_label)]

    try:
        wb = load_cached(active.path)
    except Exception as exc:                                        # noqa: BLE001
        st.error("Could not read this workbook: %s" % exc)
        st.stop()

    st.caption(
        "%d vessels · %d scheduled calls · %d past jobs"
        % (len(wb.fleet), len(wb.calls), len(wb.orders))
    )

    st.divider()
    st.markdown("### 🗓️ Target window")
    default_today = wb.as_at or _dt.date.today()
    today = st.date_input(
        "Evaluate as at", value=default_today,
        help="Drives 'months since cleaning' and 'years since drydock'. Defaults to the "
             "'data as at' date written in the workbook.",
    )
    if isinstance(today, (list, tuple)):
        today = today[0]

    next_month = _dt.date(today.year + (today.month == 12), (today.month % 12) + 1, 1)
    month_options = [
        _dt.date(next_month.year + ((next_month.month - 1 + i) // 12),
                 ((next_month.month - 1 + i) % 12) + 1, 1)
        for i in range(-1, 12)
    ]
    window_month = st.selectbox(
        "Vessels calling in", month_options,
        index=month_options.index(next_month),
        format_func=lambda d: d.strftime("%B %Y"),
    )
    window_start, window_end = month_window(window_month.year, window_month.month)

    hub_values = sorted({h.strip() for h in wb.calls["hub"].dropna() if h.strip()})
    hub_groups: dict[str, list[str]] = {}
    for hub in hub_values:
        hub_groups.setdefault(hub.strip().title(), []).append(hub)
    default_hubs = [g for g in hub_groups if g.lower() in ("busan", "singapore")] or list(hub_groups)
    picked_hubs = st.multiselect(
        "CLS service hubs", sorted(hub_groups), default=default_hubs,
        help="Only calls at these ports count towards the third gate.",
    )
    if not picked_hubs:
        st.caption("No hub selected, so every scheduled port call counts.")

    st.divider()
    st.markdown("### 🎚️ Rule thresholds")
    min_dd = st.slider(
        "Minimum years since drydocking", 0.0, 5.0, 2.0, 0.25,
        help="A vessel is only worth cleaning once its coating has aged. Your rule is 2 years.",
    )
    min_months = st.slider(
        "Minimum months since last cleaning", 0, 24, 8, 1,
        help="Never-cleaned vessels always pass this gate.",
    )
    with st.expander("Advanced"):
        interval = st.number_input(
            "Cleaning interval (months)", 1, 36, int(wb.cleaning_interval_months),
            help="Read from the workbook footer. Used for scoring, not for gating.",
        )
        use_fallback = st.checkbox(
            "Recover missing drydock dates from the contact sheet", value=True,
            help="Some vessels have no usable 'Last Dry Docking' value - it is blank, or it "
                 "holds the next survey date instead. The contact sheet often carries the real "
                 "one, and is used when the workbook corroborates it. Every vessel this affects "
                 "is listed on the Data quality tab and marked with an asterisk in the "
                 "explanation.",
        )
        measure_at = st.radio(
            "Measure vessel age", ["At today's date", "At the port call"], index=0,
            help="A vessel calling on the 29th has four more weeks of coating age than one "
                 "calling on the 1st. Measuring at the port call admits vessels that cross "
                 "a threshold part-way through the month.",
        )
        dd_slack = st.slider("'Worth a look' drydock slack (years)", 0.0, 2.0, 0.5, 0.25)
        clean_slack = st.slider("'Worth a look' cleaning slack (months)", 0.0, 6.0, 2.0, 0.5)

    settings = Settings(
        today=today, window_start=window_start, window_end=window_end,
        window_label=window_month.strftime("%B %Y"),
        min_dd_age_years=min_dd, min_months_since_clean=float(min_months),
        hubs=tuple(normalise(h) for h in picked_hubs),
        cleaning_interval_months=float(interval),
        use_contact_drydock_fallback=use_fallback,
        measure_at_call=(measure_at == "At the port call"),
        near_miss_dd_slack_years=dd_slack,
        near_miss_clean_slack_months=clean_slack,
    )

    st.divider()
    st.markdown("### ⭐ Highlighted vessels")
    highlighted = st.multiselect(
        "Single out for the client", wb.vessel_names,
        default=[v for v in st.session_state.watchlist if v in set(wb.vessel_names)],
        help="Highlighted vessels are starred in every table, pinned to the top of the "
             "Targeting tab, and are what the client brief is built from.",
        label_visibility="collapsed",
    )
    if set(highlighted) != set(st.session_state.watchlist):
        for vessel in highlighted:
            store.toggle_watch(st.session_state.watchlist, vessel, True)
        for vessel in list(st.session_state.watchlist):
            if vessel not in highlighted:
                store.toggle_watch(st.session_state.watchlist, vessel, False)
        st.rerun()

    if st.session_state.watchlist:
        if st.button("Clear highlights", use_container_width=True):
            st.session_state.watchlist = {}
            store.save_watchlist({})
            st.rerun()

df = evaluate(wb, settings)

# A call with an arrival date but no port is silently excluded by the hub filter. That happens
# when a dump is exported without cached formula results, since several hub cells are formulas.
hubless_calls = wb.calls[
    wb.calls["eta"].map(settings.window_contains) & (wb.calls["hub"].str.strip() == "")
] if not wb.calls.empty else wb.calls

watch = st.session_state.watchlist
df["starred"] = df["vessel"].isin(watch)
stats = summarise(df)


# ============================================================================ header

left, right = st.columns([3, 1])
with left:
    st.title("🚢 CLS Fleet Targeting")
    st.markdown(
        '<span class="muted">Vessels to pitch a hull cleaning to for '
        "<b>%s</b> · evaluated as at %s · source <b>%s</b></span>"
        % (settings.window_label, today.strftime("%d %b %Y"), active.original_name),
        unsafe_allow_html=True,
    )
with right:
    if wb.as_at and wb.as_at != today:
        st.info("Workbook is stamped %s." % wb.as_at.strftime("%d %b %Y"))

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Prime candidates", stats[PRIME], help=STATUS_HELP[PRIME])
k2.metric("Worth a look", stats[WATCH], help=STATUS_HELP[WATCH])
k3.metric("Calling in %s" % settings.window_label.split()[0], stats["calling"],
          help="Vessels with at least one scheduled ETA at a selected hub in the window.")
k4.metric("Never cleaned by CLS", stats["never_cleaned"],
          help="Across the whole fleet, regardless of whether they call in the window.")
k5.metric("Prime pipeline value", fmt_money(stats["pipeline_usd"]),
          help="List price for hull cleaning of the prime candidates, from the workbook's "
               "own price list by LOA band. Indicative only.")

tabs = st.tabs([
    "🎯 Targeting", "📅 Schedule", "⭐ Highlights & brief", "🚢 Fleet",
    "🔎 Vessel", "🧪 Data quality", "📂 Data & updates",
])


# ============================================================================ helpers

TABLE_COLUMNS = {
    "starred": st.column_config.CheckboxColumn("⭐", help="Highlight for the client brief", width="small"),
    "vessel": st.column_config.TextColumn("Vessel", width="medium"),
    "status": st.column_config.TextColumn("Status", width="medium"),
    "score": st.column_config.ProgressColumn("Priority", min_value=0, max_value=100,
                                             format="%.0f", width="small"),
    "last_drydock": st.column_config.DateColumn("Last drydock", format="DD MMM YYYY", width="small"),
    "dd_age_years": st.column_config.NumberColumn("Yrs since DD", format="%.1f", width="small"),
    "last_clean": st.column_config.DateColumn("Last cleaned", format="DD MMM YYYY", width="small"),
    "months_since_clean": st.column_config.NumberColumn("Months since", format="%.0f", width="small"),
    "cleaning_age": st.column_config.TextColumn("Since clean", width="small"),
    "hubs": st.column_config.TextColumn("Hub", width="small"),
    "terminals": st.column_config.TextColumn("Terminal", width="small"),
    "first_eta": st.column_config.DateColumn("First ETA", format="DD MMM", width="small"),
    "first_etd": st.column_config.DateColumn("ETD", format="DD MMM", width="small"),
    "calls_in_window": st.column_config.NumberColumn("Calls", width="small"),
    "est_value_usd": st.column_config.NumberColumn("Est. value", format="$%d", width="small"),
    "pic": st.column_config.TextColumn("Vessel PIC", width="small"),
    "pic_email": st.column_config.TextColumn("PIC email", width="medium"),
    "why": st.column_config.TextColumn("Why", width="large"),
    "next_drydock": st.column_config.DateColumn("Next drydock", format="DD MMM YYYY", width="small"),
    "job_count": st.column_config.NumberColumn("Past jobs", width="small"),
    "imo": st.column_config.NumberColumn("IMO", format="%d", width="small"),
    "qualifies_on": st.column_config.DateColumn("Qualifies on", format="DD MMM YYYY",
                                                width="small"),
    "drydock_source": st.column_config.TextColumn("Drydock date from", width="medium"),
    "near_miss_reason": st.column_config.TextColumn("Misses by", width="medium"),
    "loa": st.column_config.NumberColumn("LOA", format="%.0f m", width="small"),
}


DATE_DISPLAY_COLUMNS = {
    "last_drydock", "next_drydock", "last_clean", "first_eta", "first_etd", "eta", "etd",
    "clean_next_rec", "next_booked_clean", "last_inspection", "qualifies_on",
    "sheet_last_drydock",
}

SHORT_DATE_COLUMNS = {"first_eta", "first_etd", "eta", "etd"}

DATE_LABELS = {
    "last_drydock": "Last drydock", "next_drydock": "Next drydock",
    "last_clean": "Last cleaned", "first_eta": "First ETA", "first_etd": "ETD",
    "eta": "ETA", "etd": "ETD", "clean_next_rec": "Next rec. cleaning",
    "next_booked_clean": "Booked", "last_inspection": "Last inspection",
    "qualifies_on": "Qualifies on", "sheet_last_drydock": "On the Fleet Schedule",
}


def to_display(frame: pd.DataFrame, text_dates: bool = False) -> pd.DataFrame:
    """Prepare a frame for display.

    ``st.dataframe`` renders an empty date as a blank cell, but the editable grid behind
    ``st.data_editor`` renders it as the literal word "None". So tables that carry a pin
    checkbox ask for ``text_dates`` and get formatted strings instead.
    """
    out = frame.copy()
    for column in out.columns:
        if column in DATE_DISPLAY_COLUMNS:
            converted = pd.to_datetime(out[column], errors="coerce")
            if text_dates:
                # The window's year is already in the page heading, so in-window dates stay short.
                fmt = "%d %b" if column in SHORT_DATE_COLUMNS else "%d %b %Y"
                out[column] = converted.dt.strftime(fmt).fillna("—")
            else:
                out[column] = converted
    if text_dates and "last_clean" in out.columns and "never_cleaned" in out.columns:
        out["last_clean"] = [
            "never" if never else value
            for never, value in zip(out["never_cleaned"], out["last_clean"])
        ]
    if "months_since_clean" in out.columns:
        # "Last cleaned" already reads "never" for these, so the age column stays blank rather
        # than saying it twice.
        out["cleaning_age"] = [
            "—" if never or pd.isna(m) else "%.0f mo" % m
            for never, m in zip(out.get("never_cleaned", [False] * len(out)),
                                out["months_since_clean"])
        ]
    return out


def pinnable_table(frame: pd.DataFrame, columns: list[str], key: str, height: int | None = None):
    """A table whose ⭐ column writes straight back to the persisted highlight list."""
    frame = to_display(frame, text_dates=True)
    view = frame[["vessel"] + [c for c in columns if c != "vessel"]].copy()
    view.insert(0, "starred", frame["starred"].values)
    config = {}
    for name in view.columns:
        spec = TABLE_COLUMNS.get(name)
        if spec is None:
            continue
        # Dates arrive as text here, so a DateColumn would refuse them.
        config[name] = (
            st.column_config.TextColumn(DATE_LABELS.get(name, name), width="small")
            if name in DATE_DISPLAY_COLUMNS else spec
        )
    extra = {"height": height} if height else {}
    edited = st.data_editor(
        view,
        column_config=config,
        disabled=[c for c in view.columns if c != "starred"],
        hide_index=True, use_container_width=True, key=key, **extra,
    )
    changed = False
    for vessel, want in zip(edited["vessel"], edited["starred"]):
        if bool(want) != (vessel in st.session_state.watchlist):
            store.toggle_watch(st.session_state.watchlist, vessel, bool(want))
            changed = True
    if changed:
        st.rerun()


CANDIDATE_COLUMNS = [
    "vessel", "score", "last_drydock", "dd_age_years", "last_clean", "cleaning_age",
    "hubs", "terminals", "first_eta", "first_etd", "calls_in_window", "est_value_usd", "why",
]


# ============================================================================ 1. targeting

with tabs[0]:
    if watch:
        starred = df[df["starred"]].sort_values("score", ascending=False)
        st.markdown("#### ⭐ Highlighted for the client")
        cols = st.columns(min(3, max(1, len(starred))))
        for i, (_, row) in enumerate(starred.iterrows()):
            with cols[i % len(cols)]:
                st.markdown(
                    '<div class="hl-card"><h4>%s</h4>%s<p>%s</p></div>'
                    % (row["vessel"], status_pill(row["status"]), row["why"]),
                    unsafe_allow_html=True,
                )
        st.divider()

    parse_warnings = [i for i in wb.issues if i.severity in ("warning", "error")]
    if parse_warnings:
        kinds = sorted({i.field for i in parse_warnings})
        st.warning(
            "**%d problem%s reading this workbook** (%s). The candidate list below may be "
            "incomplete or wrong — see the Data quality tab."
            % (len(parse_warnings), "" if len(parse_warnings) == 1 else "s",
               ", ".join(kinds[:6]) + (", ..." if len(kinds) > 6 else ""))
        )

    if not hubless_calls.empty:
        st.warning(
            "%d scheduled call%s in %s ha%s an arrival date but no port, so %s excluded from "
            "the rule: **%s**. In the source workbook some hub cells are formulas; an export "
            "that drops their cached results loses the port name."
            % (len(hubless_calls), "" if len(hubless_calls) == 1 else "s", settings.window_label,
               "s" if len(hubless_calls) == 1 else "ve",
               "it is" if len(hubless_calls) == 1 else "they are",
               ", ".join(sorted(set(hubless_calls["vessel"]))))
        )

    prime = df[df["status"] == PRIME]
    st.markdown("#### 🎯 Prime candidates for %s" % settings.window_label)
    st.caption(
        "Not cleaned in %.0f+ months · %.1f+ years since drydocking · calling %s in %s. "
        "Tick ⭐ to add a vessel to the client brief."
        % (settings.min_months_since_clean, settings.min_dd_age_years,
           " or ".join(picked_hubs) or "any hub", settings.window_label)
    )
    if prime.empty:
        st.info(
            "No vessel passes all three gates for %s. Loosen a threshold in the sidebar, or "
            "check the 'Worth a look' list below." % settings.window_label
        )
    else:
        pinnable_table(prime, CANDIDATE_COLUMNS, key="prime_table")
        flagged = prime[prime["alerts"].map(bool)]
        if not flagged.empty:
            with st.expander("⚠️ Things to know before you call (%d)" % len(flagged)):
                for _, row in flagged.iterrows():
                    st.markdown("**%s** — %s" % (row["vessel"], " · ".join(row["alerts"])))

    near = df[df["status"] == WATCH]
    if not near.empty:
        st.markdown("#### 🟠 Worth a look — misses one gate, narrowly")
        st.caption(STATUS_HELP[WATCH])
        near_columns = [c for c in CANDIDATE_COLUMNS if c != "why"] + ["qualifies_on", "why"]
        pinnable_table(near, near_columns, key="near_table")

    st.divider()
    map_col, funnel_col = st.columns([5, 3])
    with map_col:
        st.markdown("##### Where every calling vessel sits against your rule")
        st.caption(
            "Each dot is a vessel calling a CLS hub in %s, positioned by how long since its "
            "drydocking and how long its hull has been fouling. For a hull that has never been "
            "cleaned, fouling is counted from the drydocking — those carry a diamond marker. "
            "The shaded corner is where both thresholds are met." % settings.window_label
        )
        st.plotly_chart(charts.targeting_map(df, settings), use_container_width=True)
    with funnel_col:
        st.markdown("##### How the fleet narrows down")
        st.caption("Each gate applied in turn, starting from the whole fleet.")
        st.plotly_chart(charts.funnel(gate_funnel(df, settings)), use_container_width=True)


# ============================================================================ 2. schedule

with tabs[1]:
    st.markdown("#### 📅 Port calls at CLS hubs in %s" % settings.window_label)
    st.caption(
        "Berth windows from ETA to ETD. Prime candidates and near misses are coloured; "
        "the rest of the calling fleet is shown in grey for context."
    )
    st.plotly_chart(charts.call_timeline(df, settings), use_container_width=True)

    st.divider()
    st.markdown("#### 📈 Which month is worth a campaign?")
    st.caption(
        "Scheduled calls per month, split by whether the vessel is already due a cleaning "
        "under the current thresholds. Useful for picking the next target month."
    )
    st.plotly_chart(charts.call_load_by_month(wb, df, today=today), use_container_width=True)

    st.divider()
    st.markdown("#### All scheduled calls in the window")
    rows = []
    for row in df.itertuples():
        for call in (row.window_calls or []):
            rows.append({
                "starred": row.starred, "vessel": row.vessel, "status": row.status,
                "hub": call["hub"], "terminal": call["terminal"],
                "eta": call["eta"], "etd": call["etd"],
                "stay_days": call["stay_days"], "score": row.score,
            })
    calls_frame = pd.DataFrame(rows)
    if calls_frame.empty:
        st.info("No scheduled calls at the selected hubs in %s." % settings.window_label)
    else:
        calls_frame = calls_frame.sort_values("eta")
        calls_frame.insert(0, "star", calls_frame["starred"].map({True: "⭐", False: ""}))
        st.dataframe(
            to_display(calls_frame).drop(columns=["starred"]),
            column_config={
                "star": st.column_config.TextColumn("", width="small"),
                "vessel": TABLE_COLUMNS["vessel"], "status": TABLE_COLUMNS["status"],
                "hub": st.column_config.TextColumn("Hub", width="small"),
                "terminal": st.column_config.TextColumn("Terminal", width="small"),
                "eta": st.column_config.DateColumn("ETA", format="DD MMM YYYY"),
                "etd": st.column_config.DateColumn("ETD", format="DD MMM YYYY"),
                "stay_days": st.column_config.NumberColumn("Stay (days)", width="small"),
                "score": TABLE_COLUMNS["score"],
            },
            hide_index=True, use_container_width=True,
        )


# ============================================================================ 3. highlights

with tabs[2]:
    st.markdown("#### ⭐ Highlighted vessels")
    if not watch:
        st.info(
            "Nothing highlighted yet. Tick the ⭐ box beside a vessel on the Targeting or Fleet "
            "tab, or pick vessels in the sidebar, and they will appear here ready to send."
        )
    else:
        selected = df[df["starred"]].sort_values("score", ascending=False)
        missing = [v for v in watch if v not in set(df["vessel"])]
        if missing:
            st.warning("Not in this workbook: %s" % ", ".join(missing))

        notes: dict[str, str] = {}
        with st.expander("Add a per-vessel note for the brief"):
            for vessel in selected["vessel"]:
                notes[vessel] = st.text_input(
                    vessel, value=watch.get(vessel, {}).get("note", ""),
                    key="note_%s" % vessel,
                )
            if st.button("Save notes"):
                for vessel, note in notes.items():
                    store.toggle_watch(st.session_state.watchlist, vessel, True, note)
                st.success("Notes saved.")

        st.dataframe(
            to_display(selected)[["vessel", "status", "score", "last_drydock", "dd_age_years",
                                  "last_clean", "cleaning_age", "hubs", "first_eta", "first_etd",
                                  "pic", "pic_email", "est_value_usd", "why"]],
            column_config=TABLE_COLUMNS, hide_index=True, use_container_width=True,
        )

        st.divider()
        st.markdown("#### ✉️ Client brief")
        bcol1, bcol2 = st.columns(2)
        client = bcol1.text_input("Client name", value="Sinokor")
        sender = bcol2.text_input("Sign-off", value="")

        records = []
        for _, row in selected.iterrows():
            record = row.to_dict()
            record["note"] = st.session_state.watchlist.get(row["vessel"], {}).get("note", "")
            records.append(record)
        email = brief_mod.build_email(records, settings, client=client, sender=sender)
        st.text_area("Copy and paste into your email", email, height=440)

        table = brief_mod.build_table(selected)
        d1, d2 = st.columns(2)
        d1.download_button(
            "⬇️ Download brief as Excel",
            brief_mod.to_excel({"Highlighted": table,
                                "Prime candidates": brief_mod.build_table(df[df["status"] == PRIME])}),
            file_name=brief_mod.filename("CLS_brief", settings, "xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        d2.download_button(
            "⬇️ Download brief as CSV", brief_mod.to_csv(table),
            file_name=brief_mod.filename("CLS_brief", settings, "csv"),
            mime="text/csv", use_container_width=True,
        )


# ============================================================================ 4. fleet

with tabs[3]:
    st.markdown("#### 🚢 Whole fleet")
    f1, f2, f3 = st.columns([2, 2, 3])
    status_filter = f1.multiselect("Status", STATUS_ORDER, default=[])
    pic_filter = f2.multiselect("Vessel PIC", sorted({p for p in df["pic"] if p}), default=[])
    search = f3.text_input("Search vessel or IMO", "")

    view = df.copy()
    if status_filter:
        view = view[view["status"].isin(status_filter)]
    if pic_filter:
        view = view[view["pic"].isin(pic_filter)]
    if search.strip():
        needle = search.strip().lower()
        view = view[
            view["vessel"].str.lower().str.contains(needle, regex=False)
            | view["imo"].astype("string").fillna("").str.contains(needle, regex=False)
        ]

    sort_options = {
        "Priority score": ("score", False), "Status": ("status_rank", True),
        "Vessel name": ("vessel", True), "Years since drydock": ("dd_age_years", False),
        "Months since cleaning": ("months_since_clean", False),
        "First ETA in window": ("first_eta", True),
    }
    sort_by = st.selectbox(
        "Sort by", list(sort_options), index=1,
        help="Status first keeps the prime candidates at the top. Sorting by priority score "
             "instead mixes in vessels that fail a gate — the score measures how overdue a "
             "vessel is, not whether it qualifies.",
    )
    sort_column, ascending = sort_options[sort_by]
    view = view.sort_values(sort_column, ascending=ascending, na_position="last")

    st.caption("%d of %d vessels" % (len(view), len(df)))
    pinnable_table(
        view,
        ["vessel", "status", "score", "imo", "loa", "pic", "last_drydock", "dd_age_years",
         "next_drydock", "last_clean", "cleaning_age", "job_count", "hubs",
         "first_eta", "calls_in_window", "why"],
        key="fleet_table", height=620,
    )

    with st.expander("What the statuses mean"):
        for status in STATUS_ORDER:
            st.markdown("%s &nbsp;%s" % (status_pill(status), STATUS_HELP[status]),
                        unsafe_allow_html=True)

    st.download_button(
        "⬇️ Download this view",
        brief_mod.to_excel({"Fleet": brief_mod.build_table(view)}),
        file_name=brief_mod.filename("CLS_fleet", settings, "xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================================ 5. vessel

with tabs[4]:
    pick = st.selectbox("Vessel", df["vessel"].sort_values().tolist(),
                        index=None, placeholder="Choose a vessel...")
    if not pick:
        st.info("Pick a vessel to see its full history, schedule and contacts.")
    else:
        row = df[df["vessel"] == pick].iloc[0]
        head, action = st.columns([4, 1])
        with head:
            st.markdown("### %s" % row["vessel"])
            st.markdown(status_pill(row["status"]), unsafe_allow_html=True)
            st.caption(row["why"])
        with action:
            on = st.checkbox("⭐ Highlight", value=bool(row["starred"]), key="detail_star")
            if on != bool(row["starred"]):
                store.toggle_watch(st.session_state.watchlist, pick, on)
                st.rerun()

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("IMO", "—" if pd.isna(row["imo"]) else "%d" % row["imo"])
        m2.metric("LOA", "—" if pd.isna(row["loa"]) else "%.1f m" % row["loa"])
        m3.metric("Years since drydock",
                  "—" if pd.isna(row["dd_age_years"]) else "%.1f" % row["dd_age_years"])
        m4.metric("Months since cleaning",
                  "never" if row["never_cleaned"] else
                  ("—" if pd.isna(row["months_since_clean"]) else "%.0f" % row["months_since_clean"]))

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### Key dates")
            st.write(pd.DataFrame([
                ("Last drydocking", fmt_date(row["last_drydock"])),
                ("Next drydock", fmt_date(row["next_drydock"])),
                ("Last CLS cleaning", "never" if row["never_cleaned"] else fmt_date(row["last_clean"])),
                ("Recorded in", row["last_clean_source"] or "—"),
                ("Next recommended cleaning", fmt_date(row["clean_next_rec"])),
                ("Cleaning already booked", fmt_date(row["next_booked_clean"])),
                ("Last inspection", fmt_date(row["last_inspection"])),
            ], columns=["", "Date"]).set_index(""))
        with c2:
            st.markdown("##### Commercial & contact")
            contact = wb.contacts[wb.contacts["key"] == row["key"]]
            extra = contact.iloc[0] if not contact.empty else None
            st.write(pd.DataFrame([
                ("Vessel PIC", row["pic"] or "—"),
                ("PIC email", row["pic_email"] or "—"),
                ("Group email", row["group_email"] or "—"),
                ("Vessel email", (extra["vessel_email"] if extra is not None else "") or "—"),
                ("Agency", (extra["agency"] if extra is not None else "") or "—"),
                ("Call cycle", "%d days" % extra["call_cycle_days"]
                 if extra is not None and pd.notna(extra["call_cycle_days"]) else "—"),
                ("Est. cleaning value", fmt_money(row["est_value_usd"])),
            ], columns=["", "Value"]).set_index(""))

        if row["alerts"]:
            st.markdown("##### Flags")
            for flag in row["alerts"]:
                st.markdown("- %s" % flag)

        st.markdown("##### Scheduled calls")
        mine = wb.calls[wb.calls["key"] == row["key"]].sort_values("seq")
        if mine.empty:
            st.caption("No scheduled calls on file.")
        else:
            st.dataframe(
                to_display(mine)[["seq", "hub", "terminal", "eta", "etd"]],
                column_config={
                    "seq": st.column_config.NumberColumn("Call", width="small"),
                    "hub": st.column_config.TextColumn("Hub"),
                    "terminal": st.column_config.TextColumn("Terminal"),
                    "eta": st.column_config.DateColumn("ETA", format="DD MMM YYYY"),
                    "etd": st.column_config.DateColumn("ETD", format="DD MMM YYYY"),
                },
                hide_index=True, use_container_width=True,
            )

        st.markdown("##### Job history")
        jobs = wb.orders[wb.orders["key"] == row["key"]].sort_values("eta", ascending=False)
        if jobs.empty:
            st.caption("No jobs recorded on the order log.")
        else:
            st.dataframe(
                to_display(jobs)[["eta", "scope", "hub", "status", "fouling", "report", "invoiced", "remarks"]],
                column_config={
                    "eta": st.column_config.DateColumn("Date", format="DD MMM YYYY"),
                    "scope": st.column_config.TextColumn("Scope"),
                    "hub": st.column_config.TextColumn("Hub", width="small"),
                    "status": st.column_config.TextColumn("Status", width="small"),
                    "fouling": st.column_config.NumberColumn("Fouling", width="small"),
                    "report": st.column_config.TextColumn("Report", width="small"),
                    "invoiced": st.column_config.TextColumn("Invoiced", width="small"),
                    "remarks": st.column_config.TextColumn("Remarks", width="large"),
                },
                hide_index=True, use_container_width=True,
            )

        if row["remarks"] or row["recommend_action"]:
            st.markdown("##### Workbook notes")
            for note in (row["recommend_action"], row["remarks"]):
                if note:
                    st.markdown("- %s" % note)


# ============================================================================ 6. data quality

with tabs[5]:
    st.markdown("#### 🧪 Data quality")
    st.caption(
        "Every issue here weakens the targeting. Fixing them in the source workbook and "
        "re-uploading is the fastest way to improve the candidate list."
    )

    recovered = df[df["drydock_source"] == DD_FROM_CONTACT]
    unresolved = df[df["drydock_source"] == DD_UNKNOWN]
    dup_imo = df[df["imo"].notna() & df["imo"].duplicated(keep=False)].sort_values("imo")
    unmatched = wb.orders[~wb.orders["matched"]] if not wb.orders.empty else pd.DataFrame()
    split_jobs = df[df["last_clean_source"].str.contains("Invoice", na=False)]

    q1, q2, q3, q4, q5 = st.columns(5)
    q5.metric("Calls with no port", len(hubless_calls),
              help="A scheduled arrival with no hub name cannot be matched to a service hub, "
                   "so it is excluded from the rule.")
    q1.metric("Drydock dates recovered", len(recovered),
              help="The Fleet Schedule value was blank or in the future, so the contact "
                   "sheet was used instead.")
    q2.metric("Drydock still unknown", len(unresolved))
    q3.metric("Duplicate IMO", len(dup_imo))
    q4.metric("Unmatched order-log rows", len(unmatched))

    if wb.job_checksum:
        counted = {
            "UWHC": int(wb.orders["scope"].str.upper().str.contains("UWHC", na=False).sum()),
            "UWI": int(wb.orders["scope"].str.upper().str.fullmatch(r"\s*UWI\s*", na=False).sum()),
        }
        mismatch = {k: (v, counted.get(k)) for k, v in wb.job_checksum.items()
                    if k in counted and counted[k] != v}
        if mismatch:
            st.error(
                "The summary sheet says this dump should hold "
                + ", ".join("%d %s" % (v[0], k) for k, v in mismatch.items())
                + " jobs, but the Order Log contains "
                + ", ".join("%d" % v[1] for v in mismatch.values())
                + ". The job history is incomplete, so some vessels will look more overdue "
                  "than they are."
            )
        else:
            st.success(
                "Order Log matches the summary sheet ("
                + ", ".join("%d %s" % (v, k) for k, v in wb.job_checksum.items()) + ")."
            )

    if not hubless_calls.empty:
        st.markdown("##### Scheduled calls with no port name")
        st.caption(
            "These arrivals fall inside the window but carry no hub, so they cannot be matched "
            "against the CLS service hubs and are excluded. Several hub cells in the source "
            "workbook are formulas referring to the first call; an export that drops cached "
            "formula results empties them."
        )
        st.dataframe(
            to_display(hubless_calls)[["vessel", "seq", "terminal", "eta", "etd"]],
            column_config={
                "vessel": TABLE_COLUMNS["vessel"],
                "seq": st.column_config.NumberColumn("Call", width="small"),
                "terminal": st.column_config.TextColumn("Terminal", width="small"),
                "eta": st.column_config.DateColumn("ETA", format="DD MMM YYYY"),
                "etd": st.column_config.DateColumn("ETD", format="DD MMM YYYY"),
            },
            hide_index=True, use_container_width=True,
        )

    if not recovered.empty:
        st.markdown("##### Drydock dates recovered from the contact sheet")
        st.caption(
            "These %d vessels have no usable 'Last Dry Docking' value on the Fleet Schedule — it "
            "is either blank or dated in the future, because the two drydock columns hold the "
            "next special survey (+5 years) and the intermediate survey (+3 years) rather than "
            "the last docking. The contact sheet's 'LAST CLEANING' column holds the real "
            "drydocking date for them, corroborated by a roughly five-year gap to the next "
            "value and by not colliding with a known cleaning, so it is used instead. "
            % len(recovered)
            + ("This is switched on; untick it in the sidebar to see the strict reading."
               if settings.use_contact_drydock_fallback
               else "This is switched OFF, so these vessels are unclassified.")
        )
        st.dataframe(
            to_display(recovered)[["vessel", "sheet_last_drydock", "last_drydock",
                                   "dd_age_years", "status", "first_eta"]],
            column_config=TABLE_COLUMNS | {
                "sheet_last_drydock": st.column_config.DateColumn(
                    "On the Fleet Schedule", format="DD MMM YYYY"),
                "last_drydock": st.column_config.DateColumn(
                    "Used instead", format="DD MMM YYYY"),
            },
            hide_index=True, use_container_width=True,
        )

    if not unresolved.empty:
        st.markdown("##### No last-drydocking date at all")
        st.caption(
            "No sheet gives a usable drydocking date for these vessels — either nothing is on "
            "file, or the contact sheet's value could not be corroborated (it clashes with a "
            "recorded cleaning, or its follow-on date is not a survey interval apart). They "
            "cannot be assessed against the drydock gate; fill the date in and they will be."
        )
        st.dataframe(
            to_display(unresolved)[["vessel", "imo", "status", "calls_in_window",
                                    "first_eta", "last_clean"]],
            column_config=TABLE_COLUMNS, hide_index=True, use_container_width=True,
        )

    if not split_jobs.empty:
        st.markdown("##### Cleanings that ran over two port calls")
        st.caption(
            "The Order Log records only the first leg of these jobs and marks it 'Partial'. "
            "The completion date comes from the Invoice sheet — without it these vessels would "
            "look months more overdue than they are, and could be re-pitched by mistake."
        )
        st.dataframe(
            to_display(split_jobs)[["vessel", "last_clean", "cleaning_age",
                                    "last_clean_source", "status"]],
            column_config=TABLE_COLUMNS | {
                "last_clean_source": st.column_config.TextColumn("Sources used", width="medium"),
            },
            hide_index=True, use_container_width=True,
        )

    if not dup_imo.empty:
        st.markdown("##### Duplicate IMO numbers")
        st.caption("An IMO number is unique to a hull, so one of these rows is wrong.")
        st.dataframe(dup_imo[["vessel", "imo", "loa", "pic"]],
                     column_config=TABLE_COLUMNS, hide_index=True, use_container_width=True)

    if not unmatched.empty:
        st.markdown("##### Order-log rows that match no vessel")
        st.caption("Their cleaning history is not being counted against any vessel.")
        st.dataframe(to_display(unmatched)[["vessel", "scope", "eta", "hub", "status"]],
                     hide_index=True, use_container_width=True)

    if wb.issues:
        st.markdown("##### Parsing notes")
        issues = pd.DataFrame([{
            "Severity": i.severity, "Sheet": i.sheet, "Vessel": i.vessel,
            "Field": i.field, "Detail": i.detail, "Cell value": i.raw,
        } for i in wb.issues])
        counts = issues["Severity"].value_counts().to_dict()
        st.caption(" · ".join("%d %s" % (v, k) for k, v in counts.items()))
        st.dataframe(issues, hide_index=True, use_container_width=True, height=300)
    else:
        st.success("Every cell parsed cleanly.")


# ============================================================================ 7. data & updates

with tabs[6]:
    st.markdown("#### 📂 Update the data")
    st.caption(
        "Drop in a new export of the fleet workbook. Columns are matched by their header text, "
        "so a dump with moved or renamed columns still loads. The previous files are kept, so "
        "you can switch back at any time from the sidebar."
    )

    upload = st.file_uploader("New fleet workbook (.xlsx)", type=["xlsx", "xlsm", "xls"],
                              key="uploader")
    if upload is not None:
        raw = upload.getvalue()
        saved, note = store.save_upload(raw, upload.name)
        try:
            fresh = load_workbook(saved)
        except Exception as exc:                                    # noqa: BLE001
            if note is None:
                store.delete_data_file(saved)
            st.error(
                "That file could not be read, so nothing was changed.\n\n**%s**\n\n"
                "Check that it still has a fleet-schedule sheet with a vessel-name column." % exc
            )
        else:
            if note:
                st.info(note)
            else:
                st.success(
                    "Loaded **%s** — %d vessels, %d scheduled calls, %d past jobs."
                    % (upload.name, len(fresh.fleet), len(fresh.calls), len(fresh.orders))
                )
            if fresh.job_checksum:
                counted = {
                    "UWHC": int(fresh.orders["scope"].str.upper()
                                .str.contains("UWHC", na=False).sum()),
                    "UWI": int(fresh.orders["scope"].str.upper()
                               .str.fullmatch(r"\s*UWI\s*", na=False).sum()),
                }
                short = {k: (v, counted[k]) for k, v in fresh.job_checksum.items()
                         if k in counted and counted[k] != v}
                if short:
                    st.warning(
                        "This dump's summary sheet says it should hold "
                        + ", ".join("%d %s" % (v[0], k) for k, v in short.items())
                        + " jobs, but its Order Log has "
                        + ", ".join("%d" % v[1] for v in short.values())
                        + ". Some cleaning history is missing, so vessels may look more "
                          "overdue than they are."
                    )

            errs = [i for i in fresh.issues if i.severity == "error"]
            warns = [i for i in fresh.issues if i.severity == "warning"]
            if errs or warns:
                with st.expander("%d parsing warning(s) in the new file" % (len(errs) + len(warns))):
                    st.dataframe(pd.DataFrame([{
                        "Severity": i.severity, "Vessel": i.vessel, "Field": i.field,
                        "Detail": i.detail, "Cell value": i.raw} for i in errs + warns],
                    ), hide_index=True, use_container_width=True)

            others = [f for f in store.list_data_files() if f.path != saved]
            if others:
                st.markdown("##### What changed since %s" % others[0].original_name)
                try:
                    delta = store.diff_workbooks(load_cached(others[0].path), fresh)
                except Exception as exc:                            # noqa: BLE001
                    st.caption("Could not compare with the previous file: %s" % exc)
                else:
                    d1, d2, d3, d4 = st.columns(4)
                    d1.metric("Vessels added", len(delta["added"]))
                    d2.metric("Vessels removed", len(delta["removed"]))
                    d3.metric("Date changes", len(delta["changes"]))
                    d4.metric("Schedule changes", len(delta["schedule"]))
                    for title, frame in (
                        ("New vessels", delta["added"]), ("Removed vessels", delta["removed"]),
                        ("Changed dates", delta["changes"]),
                        ("Changed port calls", delta["schedule"]),
                    ):
                        if not frame.empty:
                            with st.expander("%s (%d)" % (title, len(frame))):
                                st.dataframe(frame, hide_index=True, use_container_width=True)

            if st.button("Use this file", type="primary"):
                st.cache_data.clear()
                st.rerun()

    st.divider()
    st.markdown("##### Stored workbooks")
    for data_file in store.list_data_files():
        c1, c2, c3, c4 = st.columns([4, 2, 2, 1])
        is_active = data_file.path == active.path
        c1.markdown(("**%s**" if is_active else "%s") % data_file.original_name
                    + ("  ·  *in use*" if is_active else ""))
        c2.caption(data_file.uploaded.strftime("%d %b %Y %H:%M"))
        c3.caption("%d KB" % data_file.size_kb)
        if not is_active and c4.button("Remove", key="rm_%s" % data_file.path.name):
            store.delete_data_file(data_file.path)
            st.cache_data.clear()
            st.rerun()

    st.divider()
    with st.expander("How the workbook is read"):
        st.markdown(
            "- **Sheets** are found by name, tolerating renames "
            "(`Fleet Schedule`, `Order Log`, `Fleet Contact Info`, `Price list`).\n"
            "- **Columns** are matched on header text, not position.\n"
            "- **Dates** are read whether they are real dates, Excel serials or text such as "
            "`13/7/2026`; day-first versus month-first is decided per column from the "
            "unambiguous values in it.\n"
            "- **Cleaning history** merges the Fleet Schedule's confirmed-cleaning column with "
            "the Order Log, joined on a normalised vessel name so spelling differences still "
            "match.\n"
            "- **The contact sheet's own cleaning columns are ignored** — in the source "
            "workbook they often hold drydock dates and contradict the Fleet Schedule. Only "
            "its contact details, terminals and call cycle are used.\n"
        )
        st.json({
            "sheets matched": wb.sheet_map,
            "cleaning interval (months)": wb.cleaning_interval_months,
            "inspection interval (months)": wb.inspection_interval_months,
            "workbook as-at": str(wb.as_at),
            "price list": {k: v for k, v in wb.prices.items() if k != "hull_bands"},
        })
