"""The four charts that earn their place.

Colour discipline: only the two series the user acts on carry a hue (prime candidates in blue,
near-misses in orange). Everything else is a recessive grey backdrop rather than a third series,
which keeps every chart inside the validated two-hue case and keeps attention where it belongs.
Each series is labelled in a legend and, on the timeline, directly on the axis, so identity is
never colour-alone.
"""
from __future__ import annotations

import datetime as _dt

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from .rules import PRIME, WATCH, Settings

BLUE = "#2a78d6"        # categorical slot 1
ORANGE = "#eb6834"      # categorical slot 2
MUTED = "#898781"       # muted ink - the backdrop, not a series
BASELINE = "#c3c2b7"
INK = "#0b0b0b"
INK_2 = "#52514e"
SURFACE = "#fcfcfb"

FOCUS_LABEL = {PRIME: "Prime candidate", WATCH: "Worth a look"}
FOCUS_COLOUR = {PRIME: BLUE, WATCH: ORANGE}


def _base_layout(fig: go.Figure, height: int, title: str = "") -> go.Figure:
    if title:
        fig.update_layout(title=dict(text=title, font=dict(size=15, color=INK)))
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=8, t=40 if title else 12, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="system-ui, -apple-system, Segoe UI, sans-serif",
                  size=12, color=INK_2),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0,
                    bgcolor="rgba(0,0,0,0)"),
        legend_title_text="",
        hoverlabel=dict(bgcolor=SURFACE, bordercolor=BASELINE,
                        font=dict(color=INK, size=12)),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(195,194,183,0.45)", gridwidth=1,
                     zeroline=False, linecolor=BASELINE, ticks="outside",
                     tickcolor=BASELINE, tickfont=dict(color=MUTED))
    fig.update_yaxes(showgrid=True, gridcolor="rgba(195,194,183,0.45)", gridwidth=1,
                     zeroline=False, linecolor=BASELINE, ticks="outside",
                     tickcolor=BASELINE, tickfont=dict(color=MUTED))
    return fig


# ---------------------------------------------------------------------------- targeting map


def targeting_map(df: pd.DataFrame, settings: Settings, label_limit: int = 8) -> go.Figure:
    """Years since drydock (x) against months of fouling exposure (y).

    The y-axis is months since the last cleaning, or - for a hull that has never been cleaned -
    months since it came out of drydock, because that is when the fouling started. Plotting
    never-cleaned vessels at a single "never" ceiling would stack thirty of them on one line and
    hide the thing that separates them.

    Never-cleaned hulls carry a diamond marker as well as their position, so that distinction
    does not rest on the axis alone. Colour comes from the evaluated status rather than from the
    axes, so a point can never contradict the table.
    """
    data = df[df["calls_in_window"] > 0].copy()
    if data.empty:
        return _base_layout(go.Figure(), 420, "No vessels call a CLS hub in this window")

    def exposure(row):
        if not row["never_cleaned"]:
            return row["months_since_clean"]
        return None if pd.isna(row["dd_age_years"]) else row["dd_age_years"] * 12

    data["y"] = data.apply(exposure, axis=1)
    data["y_label"] = data.apply(
        lambda r: ("never cleaned - %.0f months since drydock" % r["y"])
        if r["never_cleaned"] and pd.notna(r["y"])
        else ("never cleaned" if r["never_cleaned"] else "%.0f months since cleaning" % r["y"]),
        axis=1,
    )
    data["x"] = data["dd_age_years"]
    data["x_label"] = data["dd_age_years"].apply(
        lambda v: "no drydock date" if pd.isna(v) else "%.1f yrs" % v
    )
    data = data.dropna(subset=["x", "y"])
    if data.empty:
        return _base_layout(go.Figure(), 420, "Not enough dates to plot")

    x_max, y_max = float(data["x"].max()), float(data["y"].max())

    fig = go.Figure()

    # The region where both date thresholds are met.
    fig.add_shape(
        type="rect", x0=settings.min_dd_age_years, x1=x_max * 1.15 + 1,
        y0=settings.min_months_since_clean, y1=y_max * 1.15 + 4,
        fillcolor="rgba(42,120,214,0.06)", line=dict(width=0), layer="below",
    )
    fig.add_vline(x=settings.min_dd_age_years, line=dict(color=BASELINE, width=2, dash="dot"))
    fig.add_hline(y=settings.min_months_since_clean, line=dict(color=BASELINE, width=2, dash="dot"))
    fig.add_annotation(
        x=settings.min_dd_age_years, y=y_max * 1.15 + 4, yanchor="top", xanchor="left",
        text="  %.1f yrs since drydock" % settings.min_dd_age_years,
        showarrow=False, font=dict(size=11, color=MUTED),
    )
    fig.add_annotation(
        x=x_max * 1.15 + 1, y=settings.min_months_since_clean, xanchor="right", yanchor="bottom",
        text="%.0f months of fouling  " % settings.min_months_since_clean,
        showarrow=False, font=dict(size=11, color=MUTED),
    )
    fig.add_annotation(
        x=0, y=1.0, xref="paper", yref="paper", xanchor="left", yanchor="bottom",
        text="Diamonds have never been cleaned, so they clear the cleaning threshold "
             "wherever they sit on the vertical axis.",
        showarrow=False, font=dict(size=10, color=MUTED),
    )

    hover = ("<b>%{customdata[0]}</b><br>%{customdata[1]}"
             "<br>Drydock: %{customdata[2]} ago"
             "<br>%{customdata[3]}<extra></extra>")
    fields = ["vessel", "status", "x_label", "y_label"]

    backdrop = data[~data["status"].isin(FOCUS_COLOUR)]
    for never, symbol in ((False, "circle"), (True, "diamond")):
        part = backdrop[backdrop["never_cleaned"] == never]
        if part.empty:
            continue
        fig.add_trace(go.Scatter(
            x=part["x"], y=part["y"], mode="markers", name="Other vessels calling",
            legendgroup="other", showlegend=not never,
            marker=dict(size=9, color=MUTED, opacity=0.5, symbol=symbol,
                        line=dict(width=2, color=SURFACE)),
            customdata=part[fields], hovertemplate=hover,
        ))

    # Only the strongest few carry an inline label; the rest would collide into a smear.
    labelled = set(
        data[data["status"].isin(FOCUS_COLOUR)].nlargest(label_limit, "score")["vessel"]
    )
    for status, colour in FOCUS_COLOUR.items():
        for never, symbol in ((False, "circle"), (True, "diamond")):
            part = data[(data["status"] == status) & (data["never_cleaned"] == never)]
            if part.empty:
                continue
            fig.add_trace(go.Scatter(
                x=part["x"], y=part["y"], mode="markers+text", name=FOCUS_LABEL[status],
                legendgroup=status, showlegend=not never,
                marker=dict(size=13, color=colour, symbol=symbol,
                            line=dict(width=2, color=SURFACE)),
                text=[v.title() if v in labelled else "" for v in part["vessel"]],
                textposition="middle right", textfont=dict(size=10, color=INK_2),
                cliponaxis=False,
                customdata=part[fields], hovertemplate=hover,
            ))

    fig = _base_layout(fig, 460)
    fig.update_xaxes(title_text="Years since last drydocking", range=[-0.3, x_max * 1.15 + 1])
    fig.update_yaxes(title_text="Months of fouling exposure", range=[-2, y_max * 1.15 + 4])
    return fig


# ---------------------------------------------------------------------------- port-call timeline


def call_timeline(df: pd.DataFrame, settings: Settings, limit: int = 28) -> go.Figure:
    """Every scheduled CLS-hub call inside the window, as a berth-window Gantt."""
    bars = []
    for row in df.itertuples():
        for call in (row.window_calls or []):
            eta = call["eta"]
            etd = call["etd"] or eta
            if etd < eta:
                etd = eta
            bars.append({
                "vessel": row.vessel,
                "status": row.status,
                "start": _dt.datetime.combine(eta, _dt.time(6, 0)),
                "finish": _dt.datetime.combine(etd, _dt.time(18, 0)),
                "hub": call["hub"],
                "terminal": call["terminal"],
                "score": row.score,
                "eta_text": eta.strftime("%d %b"),
                "etd_text": etd.strftime("%d %b"),
            })
    if not bars:
        return _base_layout(go.Figure(), 320, "No scheduled calls in this window")

    frame = pd.DataFrame(bars)
    # Keep the chart readable: the most interesting vessels first, then by arrival.
    rank = frame.groupby("vessel")["score"].max().sort_values(ascending=False)
    focus = [v for v in rank.index
             if frame.loc[frame["vessel"] == v, "status"].iloc[0] in FOCUS_COLOUR]
    rest = [v for v in rank.index if v not in focus]
    keep = (focus + rest)[:limit]
    frame = frame[frame["vessel"].isin(keep)]
    order = frame.groupby("vessel")["start"].min().sort_values(ascending=False).index.tolist()

    frame["series"] = frame["status"].map(FOCUS_LABEL).fillna("Other vessels calling")
    colours = {FOCUS_LABEL[PRIME]: BLUE, FOCUS_LABEL[WATCH]: ORANGE, "Other vessels calling": MUTED}

    fig = px.timeline(
        frame, x_start="start", x_end="finish", y="vessel", color="series",
        color_discrete_map=colours,
        category_orders={"vessel": order,
                         "series": [FOCUS_LABEL[PRIME], FOCUS_LABEL[WATCH], "Other vessels calling"]},
        custom_data=["hub", "terminal", "eta_text", "etd_text", "status"],
    )
    fig.update_traces(
        marker_line_color=SURFACE, marker_line_width=2,
        hovertemplate=("<b>%{y}</b><br>%{customdata[4]}"
                       "<br>%{customdata[0]} %{customdata[1]}"
                       "<br>ETA %{customdata[2]} - ETD %{customdata[3]}<extra></extra>"),
    )
    height = max(320, 26 * len(order) + 90)
    fig = _base_layout(fig, height)
    fig.update_yaxes(title_text="", tickfont=dict(size=11, color=INK_2), showgrid=False)
    fig.update_xaxes(
        title_text="", dtick=86400000 * 2, tickformat="%d %b",
        range=[_dt.datetime.combine(settings.window_start, _dt.time(0, 0)),
               _dt.datetime.combine(settings.window_end, _dt.time(23, 59))],
    )
    for trace in fig.data:
        if trace.name == "Other vessels calling":
            trace.opacity = 0.5
    return fig


# ---------------------------------------------------------------------------- funnel


def funnel(stages: pd.DataFrame) -> go.Figure:
    """How the fleet narrows down, one gate at a time. Single series, so no legend."""
    if stages.empty:
        return _base_layout(go.Figure(), 220)
    data = stages.iloc[::-1]
    fig = go.Figure(go.Bar(
        x=data["vessels"], y=data["stage"], orientation="h",
        marker=dict(color=BLUE, line=dict(width=0)),
        text=data["vessels"], textposition="outside",
        textfont=dict(size=12, color=INK_2), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x} vessels<extra></extra>",
    ))
    fig = _base_layout(fig, 40 * len(data) + 70)
    fig.update_traces(marker_cornerradius=4)
    fig.update_xaxes(title_text="", showgrid=True, range=[0, data["vessels"].max() * 1.18])
    fig.update_yaxes(title_text="", showgrid=False, tickfont=dict(size=12, color=INK_2))
    return fig


# ---------------------------------------------------------------------------- forward load


def call_load_by_month(wb, df: pd.DataFrame, months: int = 9,
                       today: _dt.date | None = None) -> go.Figure:
    """Scheduled CLS-hub calls per month, split by whether the vessel is due a cleaning.

    Answers the planning question the targeting tab cannot: which of the coming months has both
    the port calls and the overdue vessels to make a campaign worth running.
    """
    today = today or _dt.date.today()
    if wb.calls.empty:
        return _base_layout(go.Figure(), 280)

    due = set(df.loc[df["gate_dd"] & df["gate_clean"], "key"])
    calls = wb.calls.dropna(subset=["eta"]).copy()
    calls = calls[calls["eta"].map(lambda d: d is not None and d >= today.replace(day=1))]
    if calls.empty:
        return _base_layout(go.Figure(), 280, "No future port calls on file")

    calls["month"] = calls["eta"].map(lambda d: _dt.date(d.year, d.month, 1))
    calls["due"] = calls["key"].isin(due)
    cutoff = sorted(calls["month"].unique())[:months]
    calls = calls[calls["month"].isin(cutoff)]

    grouped = calls.groupby(["month", "due"]).size().unstack(fill_value=0)
    for column in (True, False):
        if column not in grouped.columns:
            grouped[column] = 0
    labels = [m.strftime("%b %Y") for m in grouped.index]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=labels, y=grouped[True], name="Vessel is due a cleaning",
        marker=dict(color=BLUE, line=dict(width=2, color=SURFACE)),
        hovertemplate="<b>%{x}</b><br>%{y} calls by vessels due a cleaning<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        x=labels, y=grouped[False], name="Not due",
        marker=dict(color=MUTED, opacity=0.5, line=dict(width=2, color=SURFACE)),
        hovertemplate="<b>%{x}</b><br>%{y} other calls<extra></extra>",
    ))
    fig.update_layout(barmode="stack", bargap=0.35)
    fig = _base_layout(fig, 320)
    fig.update_traces(marker_cornerradius=4)
    fig.update_xaxes(title_text="", showgrid=False)
    fig.update_yaxes(title_text="Scheduled CLS-hub calls")
    return fig
