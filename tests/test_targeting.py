"""Regression tests for the targeting rule engine.

Run from the project root:   python -m pytest cls_dashboard/tests -q
or without pytest:           python cls_dashboard/tests/test_targeting.py

The fixture is the September 2026 Sinokor dump shipped in ``data/``, evaluated as at
2026-09-22 for an October 2026 window. The two acceptance cases the user gave by name
(HEUNG-A YOUNG in, NINGBO TRADER out) are asserted directly.
"""
from __future__ import annotations

import datetime as _dt
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pytest  # noqa: E402

from cls.dates import coerce_date, infer_dayfirst, parse_date_column  # noqa: E402
from cls.loader import load_workbook  # noqa: E402
from cls.names import match, normalise  # noqa: E402
from cls.rules import (  # noqa: E402
    BOOKED, DD_FROM_CONTACT, DD_FROM_FLEET, DD_UNKNOWN, EARLY_DD, NO_DD, PRIME, RECENT,
    STATUS_ORDER, WATCH,
    Settings, estimate_value, evaluate, gate_funnel, month_window, summarise,
)

TODAY = _dt.date(2026, 9, 22)
# The client workbook is deliberately not in the repo (see .gitignore), so on a fresh
# clone these tests skip with an explanation rather than erroring out.
_WORKBOOKS = sorted((ROOT / "data").glob("*.xlsx"))
WORKBOOK = _WORKBOOKS[0] if _WORKBOOKS else None

pytestmark = pytest.mark.skipif(
    WORKBOOK is None,
    reason="No fleet workbook in data/ - drop an .xlsx export there to run these tests.",
)


@pytest.fixture(scope="module")
def wb():
    return load_workbook(WORKBOOK)


@pytest.fixture(scope="module")
def settings():
    start, end = month_window(2026, 10)
    return Settings(today=TODAY, window_start=start, window_end=end,
                    window_label="October 2026")


@pytest.fixture(scope="module")
def df(wb, settings):
    return evaluate(wb, settings)


def replace_calls(wb, calls):
    import copy
    clone = copy.copy(wb)
    clone.calls = calls
    return clone


def row_for(df, vessel):
    match_rows = df[df["vessel"] == vessel]
    assert not match_rows.empty, "%s missing from the evaluated fleet" % vessel
    return match_rows.iloc[0]


# ---------------------------------------------------------------------------- date parsing


def test_text_dates_are_read_day_first_on_the_fleet_sheet():
    # "13/7/2026" in the Fleet Schedule means 13 July, not 7 December.
    assert parse_date_column(["13/7/2026", "21/6/2026", "2/3/2026"]) == [
        _dt.date(2026, 7, 13), _dt.date(2026, 6, 21), _dt.date(2026, 3, 2),
    ]


def test_text_dates_are_read_month_first_on_the_contact_sheet():
    # The same a/b/yyyy shape, but this column's unambiguous values prove month-first.
    assert parse_date_column(["1/24/2026", "2/13/2026", "2/3/2026"]) == [
        _dt.date(2026, 1, 24), _dt.date(2026, 2, 13), _dt.date(2026, 2, 3),
    ]


def test_dotted_and_korean_dates():
    assert coerce_date("2022.10.31") == _dt.date(2022, 10, 31)
    assert coerce_date("2028년 3월(차기입거)") == _dt.date(2028, 3, 1)


def test_non_dates_are_rejected():
    for value in ("Not planned", "X", "확인불가", "", None, 9963152):
        assert coerce_date(value) is None, value


def test_imo_numbers_are_never_mistaken_for_excel_serials():
    assert coerce_date(9963152) is None
    assert coerce_date(46000) == _dt.date(2025, 12, 9)   # a plausible serial still converts


def test_dayfirst_defaults_to_day_first_when_undecidable():
    assert infer_dayfirst(["1/2/2026", "3/4/2026"]) is True


# ---------------------------------------------------------------------------- name joining


@pytest.mark.parametrize("written,canonical", [
    ("Heung A Young", "HEUNG-A YOUNG"),
    ("Toyoma Trader", "TOYAMA TRADER"),
    ("Volostchny Voyager", "VOSTOCHNY VOYAGER"),
    ("Hakata Expresss", "HAKATA EXPRESS"),
    ("Heaung A Akita", "HEUNG-A AKITA"),
    ("SAWASDEE RIGEL ", "SAWASDEE RIGEL"),
])
def test_misspellings_join_to_the_canonical_vessel(written, canonical, wb):
    index = {normalise(v): v for v in wb.vessel_names}
    assert match(written, index) == canonical


def test_every_order_log_row_joins_to_a_vessel(wb):
    unmatched = wb.orders.loc[~wb.orders["matched"], "vessel"].tolist()
    assert unmatched == [], "order-log rows with no matching vessel: %s" % unmatched


def test_only_non_fleet_invoices_are_unmatched(wb):
    unmatched = set(wb.invoices.loc[~wb.invoices["matched"], "vessel"])
    assert unmatched <= {"DUBAI ENERGY"}, unmatched


# ---------------------------------------------------------------------------- loading


def test_workbook_shape(wb):
    assert len(wb.fleet) == 73
    assert len(wb.orders) == 43
    assert wb.as_at == TODAY
    assert wb.cleaning_interval_months == 12
    assert wb.inspection_interval_months == 6
    assert wb.sheet_map["fleet"] == "Fleet Schedule"


def test_order_log_matches_the_summary_sheet_checksum(wb):
    """Sheet1 states how many jobs the dump holds; a short Order Log would silently
    make vessels look overdue."""
    cleanings = wb.orders["scope"].str.upper().str.contains("UWHC").sum()
    inspections = wb.orders["scope"].str.upper().str.fullmatch(r"\s*UWI\s*").sum()
    assert wb.job_checksum == {"UWHC": 40, "UWI": 3}
    assert (cleanings, inspections) == (40, 3)


def test_price_list_is_read_from_the_workbook(wb):
    assert wb.prices["inspection"] == 1500
    assert wb.prices["propeller"] == 4500
    assert wb.prices["second_mobilisation"] == 3000
    assert estimate_value(141.0, wb.prices) == 16000     # HEUNG-A YOUNG, 100-150 m band
    assert estimate_value(195.58, wb.prices) == 20000    # YEOSU VOYAGER, 151-200 m band
    assert estimate_value(272.21, wb.prices) == 22000    # MUMBAI BRIDGE, 251-300 m band


@pytest.mark.parametrize("loa,expected", [
    (90, 16000),      # shorter than every band
    (150, 16000), (151, 20000),
    (150.5, 16000),   # in the one-metre gap: nearest band, not the largest
    (200.5, 20000), (250.5, 20000), (300.5, 22000),
    (400, 24000),     # longer than every band
])
def test_price_bands_do_not_tile_so_gaps_take_the_nearest_band(loa, expected, wb):
    assert estimate_value(loa, wb.prices) == expected


def test_contact_sheet_enrichment_is_populated(wb):
    assert (wb.contacts["vessel_email"] != "").sum() >= 70
    assert wb.contacts["alt_drydock"].notna().sum() >= 70


# ---------------------------------------------------------------------------- the user's rule


def test_heung_a_young_is_a_prime_candidate(df):
    """The user's own worked example: drydocked 2024, never cleaned, calls Busan in October."""
    row = row_for(df, "HEUNG-A YOUNG")
    assert row["status"] == PRIME
    assert bool(row["never_cleaned"])
    assert row["last_drydock"] == _dt.date(2024, 8, 30)
    assert row["gate_dd"] and row["gate_clean"] and row["gate_call"]
    assert row["calls_in_window"] == 2
    assert row["hubs"] == "Busan"


def test_ningbo_trader_is_excluded_for_being_too_soon_after_drydock(df):
    """The user's counter-example: drydocked 2026, so not worth cleaning despite calling."""
    row = row_for(df, "NINGBO TRADER")
    assert row["status"] == EARLY_DD
    assert bool(row["gate_call"])            # it does call in October
    assert not bool(row["gate_dd"])          # but the drydock gate stops it
    assert row["last_drydock"] == _dt.date(2026, 7, 10)


def test_every_vessel_lands_in_exactly_one_bucket(df):
    assert set(df["status"]) <= set(STATUS_ORDER)
    assert df["status"].notna().all()
    assert sum(summarise(df)[s] for s in STATUS_ORDER) == len(df) == 73


def test_gates_are_consistent_with_status(df):
    for _, row in df.iterrows():
        if row["status"] == PRIME:
            assert row["gate_dd"] and row["gate_clean"] and row["gate_call"], row["vessel"]
        if row["status"] == "Not calling in window":
            assert not row["gate_call"], row["vessel"]


def test_prime_list_is_exactly_the_expected_eight(df):
    assert set(df.loc[df["status"] == PRIME, "vessel"]) == {
        "ATLANTIC SOUTH", "SAWASDEE VEGA", "SAWASDEE MIMOSA", "INCHEON VOYAGER",
        "SINOKOR NIIGATA", "YEOSU VOYAGER", "HEUNG-A YOUNG", "TOYAMA TRADER",
    }


def test_never_cleaned_vessels_outrank_recently_cleaned_ones(df):
    prime = df[df["status"] == PRIME]
    never = prime[prime["never_cleaned"]]["score"]
    cleaned = prime[~prime["never_cleaned"]]["score"]
    assert never.min() > cleaned.max()


# ------------------------------------------------- the two defects the workbook hides


def test_split_job_completion_comes_from_the_invoice(df):
    """QINGDAO VOYAGER's Order Log entry is a 'Partial' dated 21 Jan 2026, but the job ran
    over two calls and was billed out on 25 Mar 2026 (10k + 13k = the 20k band plus the 3k
    second-mobilisation fee). Reading the Order Log alone would wrongly re-pitch it."""
    row = row_for(df, "QINGDAO VOYAGER")
    assert row["last_clean"] == _dt.date(2026, 3, 25)
    assert row["status"] == RECENT
    assert row["status"] != PRIME
    assert "Invoice" in row["last_clean_source"]


@pytest.mark.parametrize("vessel,finished", [
    ("OSAKA VOYAGER", _dt.date(2026, 4, 3)),
    ("SAWASDEE INCHEON", _dt.date(2026, 4, 15)),
])
def test_other_split_jobs_also_complete_from_the_invoice(df, vessel, finished):
    assert row_for(df, vessel)["last_clean"] == finished


def test_drydock_falls_back_to_the_contact_sheet_when_the_fleet_value_is_unusable(df):
    """Five vessels carry the +5y special survey in the 'last drydocking' column. Swapping
    the two drydock columns does not recover the real date; the contact sheet does."""
    row = row_for(df, "SAWASDEE VEGA")
    assert row["sheet_last_drydock"] == _dt.date(2027, 10, 30)   # in the future, unusable
    assert row["last_drydock"] == _dt.date(2022, 10, 31)         # recovered
    assert row["drydock_source"] == DD_FROM_CONTACT
    assert row["status"] == PRIME


def test_the_fallback_is_only_used_when_needed(df):
    used = df[df["drydock_source"] == DD_FROM_CONTACT]
    assert len(used) == 7
    assert (df["drydock_source"] == DD_FROM_FLEET).sum() == 64
    # Never applied over a usable Fleet Schedule value.
    for _, row in used.iterrows():
        assert row["sheet_last_drydock"] is None or row["sheet_last_drydock"] > TODAY


def test_the_fallback_refuses_a_value_that_is_really_a_cleaning(df):
    """SAWASDEE CAPELLA's contact-sheet value is 2026-01-06 - the day after an Order Log
    cleaning on 2026-01-05 that was invoiced on 2026-01-09. Taking it as a drydocking would
    exclude the vessel for being 'too soon after drydock' three days before it was cleaned."""
    row = row_for(df, "SAWASDEE CAPELLA")
    assert row["drydock_source"] == DD_UNKNOWN
    assert row["status"] == NO_DD
    assert row["last_clean"] == _dt.date(2026, 1, 5)


def test_the_fallback_requires_a_survey_shaped_interval(df):
    """HOCHIMINH VOYAGER's contact pair is 2.98 years apart, not the ~5 years of a special
    survey, so it is not trusted as a drydocking either."""
    assert row_for(df, "HOCHIMINH VOYAGER")["drydock_source"] == DD_UNKNOWN


def test_disabling_the_fallback_demotes_those_vessels(wb, settings):
    strict = evaluate(wb, Settings(**{**settings.__dict__, "use_contact_drydock_fallback": False}))
    assert row_for(strict, "SAWASDEE VEGA")["status"] != PRIME
    assert row_for(strict, "HEUNG-A YOUNG")["status"] == PRIME     # unaffected


def test_a_future_priced_invoice_line_counts_as_booked(df):
    """AKITA TRADER has a priced job on 9 Oct 2026 that is not yet invoiced."""
    row = row_for(df, "AKITA TRADER")
    assert row["next_booked_clean"] == _dt.date(2026, 10, 9)
    assert row["status"] == BOOKED


# ---------------------------------------------------------------------------- near misses


def test_near_misses_carry_the_date_they_would_qualify(df):
    watch = df[df["status"] == WATCH]
    assert not watch.empty
    assert watch["qualifies_on"].notna().all()
    assert row_for(df, "MUMBAI BRIDGE")["qualifies_on"] == _dt.date(2026, 9, 23)


def test_measuring_at_the_port_call_admits_vessels_that_mature_mid_month(wb, settings):
    """A vessel calling on the 29th has five more weeks of coating age than one calling
    on the 1st, so the two readings should differ - and only in the permissive direction."""
    at_call = evaluate(wb, Settings(**{**settings.__dict__, "measure_at_call": True}))
    base_prime = set(evaluate(wb, settings).loc[lambda d: d["status"] == PRIME, "vessel"])
    call_prime = set(at_call.loc[at_call["status"] == PRIME, "vessel"])
    assert base_prime <= call_prime
    assert {"MUMBAI BRIDGE", "TIANJIN VOYAGER"} <= call_prime


# ---------------------------------------------------------------------------- thresholds


def test_loosening_the_drydock_gate_only_adds_vessels(wb, settings):
    base = set(evaluate(wb, settings).loc[lambda d: d["gate_dd"], "vessel"])
    loose = set(evaluate(wb, Settings(**{**settings.__dict__, "min_dd_age_years": 1.0}))
                .loc[lambda d: d["gate_dd"], "vessel"])
    assert base < loose


def test_tightening_the_cleaning_gate_only_removes_vessels(wb, settings):
    base = set(evaluate(wb, settings).loc[lambda d: d["gate_clean"], "vessel"])
    tight = set(evaluate(wb, Settings(**{**settings.__dict__, "min_months_since_clean": 18.0}))
                .loc[lambda d: d["gate_clean"], "vessel"])
    assert tight < base


def test_restricting_hubs_to_singapore_shrinks_the_calling_set(wb, settings):
    sg = evaluate(wb, Settings(**{**settings.__dict__, "hubs": ("singapore",)}))
    assert 0 < sg["gate_call"].sum() < evaluate(wb, settings)["gate_call"].sum()
    assert set(sg.loc[sg["gate_call"], "hubs"]) == {"Singapore"}


def test_multi_word_hub_names_still_match(wb, settings):
    """The loader strips spaces out of hub keys, so the sidebar's labels must be normalised the
    same way - otherwise a two-word port like Port Klang would silently match nothing."""
    from cls.names import normalise
    renamed = wb.calls.copy()
    renamed.loc[renamed["hub_key"] == "singapore", "hub"] = "Port Klang"
    renamed.loc[renamed["hub"] == "Port Klang", "hub_key"] = normalise("Port Klang")
    patched = replace_calls(wb, renamed)
    picked = evaluate(patched, Settings(**{**settings.__dict__,
                                           "hubs": (normalise("Port Klang"),)}))
    assert picked["gate_call"].sum() > 0


def test_funnel_is_monotonically_narrowing(df, settings):
    counts = gate_funnel(df, settings)["vessels"].tolist()
    assert counts == sorted(counts, reverse=True)
    assert counts[0] == 73


def test_window_with_no_calls_yields_no_candidates(wb, settings):
    start, end = month_window(2027, 3)
    quiet = evaluate(wb, Settings(**{**settings.__dict__, "window_start": start,
                                     "window_end": end, "window_label": "March 2027"}))
    assert (quiet["status"] == PRIME).sum() == 0
    assert len(quiet) == 73


# ---------------------------------------------------------------------------- explanations


@pytest.mark.parametrize("min_dd", [0.0, 2.0, 5.0])
@pytest.mark.parametrize("min_months", [0, 8, 24])
@pytest.mark.parametrize("hubs", [(), ("busan",), ("singapore",), ("busan", "singapore")])
def test_invariants_hold_across_the_whole_settings_grid(wb, settings, min_dd, min_months, hubs):
    """Whatever the user does with the sliders: nobody is lost, nobody is double-counted,
    and the funnel can only ever narrow."""
    tuned = evaluate(wb, Settings(**{**settings.__dict__, "min_dd_age_years": min_dd,
                                     "min_months_since_clean": float(min_months),
                                     "hubs": hubs}))
    assert len(tuned) == 73
    counts = summarise(tuned)
    assert sum(counts[s] for s in STATUS_ORDER) == 73
    stages = gate_funnel(tuned, settings)["vessels"].tolist()
    assert stages == sorted(stages, reverse=True)


def test_degenerate_fleets_do_not_crash(wb, settings):
    import copy
    for size in (0, 1):
        small = copy.copy(wb)
        small.fleet = wb.fleet.head(size)
        small.calls = wb.calls[wb.calls["key"].isin(set(small.fleet["key"]))]
        result = evaluate(small, settings)
        assert len(result) == size
        assert summarise(result)["fleet"] == size


def test_every_vessel_has_a_readable_explanation(df):
    assert df["why"].str.len().min() > 20
    assert row_for(df, "HEUNG-A YOUNG")["why"].startswith("never cleaned by CLS")
    assert "2 calls in October 2026" in row_for(df, "HEUNG-A YOUNG")["why"]


def test_contact_sourced_drydock_is_marked_in_the_explanation(df):
    assert "*" in row_for(df, "SAWASDEE VEGA")["why"]


def test_pipeline_value_matches_the_price_list(df, wb):
    prime = df[df["status"] == PRIME]
    assert summarise(df)["pipeline_usd"] == prime["est_value_usd"].sum() == 144000


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
