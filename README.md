# CLS Fleet Targeting

Which Sinokor vessels should we pitch a hull cleaning to, for a given month?

> **The fleet workbook is not in this repository.** It holds Sinokor's fleet, vessel PIC email
> addresses, invoice amounts and pricing, so `data/` is gitignored. Put an export in `data/`, or
> just start the app and drop the file on the upload prompt.

Double-click **`run_dashboard.bat`**, or from a terminal:

```bash
python -m streamlit run app.py
```

---

## The rule

Three gates, all adjustable from the sidebar:

| Gate | Default | Meaning |
|---|---|---|
| **Drydock** | 2+ years since the last drydocking | Below that the coating is still working, so cleaning is wasted spend |
| **Cleaning** | not cleaned in the last 8 months | Never cleaned always passes |
| **Port call** | a scheduled ETA at Busan or Singapore inside the target month | Nothing can be sold without a berth window |

Every vessel lands in exactly one bucket, so the buckets always add up to the fleet size:

**Prime candidate** · **Worth a look** (misses one gate narrowly, with the date it would qualify) ·
**Already booked** · **Cleaned recently** · **Too soon after drydock** · **No drydock date** ·
**Not calling in window**

Ranking inside a bucket is a 0–100 priority score built from three visible parts: months overdue
for cleaning (0–50), years past drydock beyond the threshold (0–30), and how easy the vessel is to
reach — number of calls and length of berth window (0–20). A never-cleaned hull scores the full
cleaning component, which is why HEUNG-A YOUNG and YEOSU VOYAGER sit above TOYAMA TRADER.

---

## Two things the workbook hides

Both were found by cross-checking sheets against each other, and both change the answer.

### 1. Nine vessels have no usable drydock date

For five of them the "Last Dry Docking" column holds the **next special survey** (true drydock
+ 5 years) and "Next Dry Dock" holds the **intermediate survey** (+ 3 years). Swapping the two
columns does *not* recover the real date — it just yields another future date.

The contact sheet's `LAST CLEANING` column is mislabelled and actually holds the drydocking date.
The Korean header on the column beside it, 차기입거, means *next docking*, and adding five years to
it reproduces the Fleet Schedule's value to within a few days:

| Vessel | Fleet Schedule says | Contact sheet | + 5 years |
|---|---|---|---|
| SAWASDEE SPICA | 2028-05-03 | 2023-05-04 | 2028-05-04 (0 days out) |
| SAWASDEE DENEB | 2028-08-24 | 2023-08-28 | 2028-08-28 (4 days out) |
| SAWASDEE VEGA | 2027-10-30 | 2022-10-31 | 2027-10-31 (1 day out) |

So the contact sheet is used as a fallback — but only where the workbook corroborates it, because
that column is not *always* a drydocking. It must have a follow-on value roughly five years later
(a survey interval), and it must not coincide with a cleaning recorded elsewhere. SAWASDEE
CAPELLA fails the second test: its value is 2026-01-06, the day after an Order Log cleaning on
2026-01-05 that was invoiced on 2026-01-09. Taking it at face value would exclude the vessel for
being "too soon after drydock" three days before it was cleaned.

Seven vessels are recovered this way; two (SAWASDEE CAPELLA, HOCHIMINH VOYAGER) are honestly
reported as having no drydock date rather than given a fabricated one. **The recovery puts
SAWASDEE VEGA and SAWASDEE MIMOSA into the October list** — both 3+ years past drydock, never
cleaned, calling Busan. Untick the sidebar option to see the strict reading; every affected
vessel is on the Data quality tab and marked with an asterisk in its explanation.

### 2. Some cleanings run over two port calls

A Busan North Port job is often split across two calls. The Order Log records only the first leg
and marks it `Partial`; the **Invoice** sheet carries the date the work actually finished, billed
as the LOA band plus a 3,000 USD second-mobilisation fee.

QINGDAO VOYAGER's Order Log entry reads 21 Jan 2026 — eight months ago, which would put it on the
October pitch list. Its invoices read 10,000 on 25 Jan and 13,000 on 25 Mar: the job finished
**25 March**, five months ago. Without the invoice, the dashboard would have sent you to the
client to re-sell a job Sinokor had already paid for. The same applies to OSAKA VOYAGER and
SAWASDEE INCHEON.

Only jobs the Order Log actually marks `Partial` are extended this way, and only as far as the
vessel's next job — otherwise an ordinary invoice for a finished cleaning, or the billing for a
later campaign, gets folded in and the vessel looks fresher than it is. Separately, a substantial
invoice with no Order Log row at all is counted as a cleaning in its own right, which picks up
five more recent cleanings than the Order Log knows about.

---

## Updating the data

**Data & updates** tab → drop in a new export. The previous files are kept, so you can switch
back from the sidebar at any time.

The loader is built to survive a changed dump:

- **Sheets** are found by name, tolerating renames and the trailing space in `Montly Vessel Schedule `.
- **Columns** are matched on header text, not position, so inserting a column does not break it.
- **Port-call blocks** repeat: a dump with one, two or five scheduled calls per vessel all work.
- **Dates** are read whether they are real dates, Excel serials, `13/7/2026`, `2022.10.31` or
  `2028년 3월(차기입거)`. Day-first versus month-first is decided **per column** from the
  unambiguous values in it — the Fleet Schedule writes day-first and the contact sheet writes
  month-first, and both parse correctly.
- **Vessel names** join on a normalised key, so `Toyoma Trader`, `Volostchny Voyager`,
  `Hakata Expresss` and `Heaung A Akita` all match. All 43 order-log rows join today. The fuzzy
  step is deliberately timid: this fleet contains HAKATA VOYAGER *and* JAKARTA VOYAGER, so a
  match is only accepted when no runner-up is close behind it. An unmatched row is reported;
  crediting one vessel's cleaning to another would not be.
- **A bad file changes nothing.** It is parsed before being adopted; if it cannot be read, the
  current data stays in place and the error is shown.
- **Silence is never the failure mode.** The Targeting tab carries a banner whenever the
  workbook raised any parsing problem. A date column whose day-first/month-first convention
  nothing settles, a non-numeric LOA, a missing Order Log or Invoice sheet, a scheduled call
  with no port name, a duplicated vessel row — each is reported rather than quietly absorbed.

After upload you get a diff against the previous dump — vessels added or removed, changed drydock
and cleaning dates, changed port calls — and a check against the summary sheet's own job count, so
a truncated Order Log is caught rather than silently making vessels look overdue.

---

## Highlighting vessels for the client

Tick the ⭐ beside any vessel, on any table, or pick them in the sidebar. Highlighted vessels are
starred everywhere, pinned as cards at the top of the Targeting tab, and survive a restart.

The **Highlights & brief** tab turns them into a ready-to-send email — vessel, IMO, LOA, last
drydocking, cleaning history, every October call with terminal and dates, the PIC and their
address, and a recommendation — plus Excel and CSV exports. Per-vessel notes you add there appear
in the brief.

---

## Layout

```
app.py                 the dashboard: seven tabs
cls/dates.py           date coercion, per-column day-first inference
cls/names.py           vessel-name normalisation and fuzzy joining
cls/loader.py          workbook -> tidy dataframes, defensively
cls/rules.py           the gates, the status taxonomy, the score, the explanations
cls/charts.py          the four charts
cls/brief.py           client email, Excel and CSV exports
cls/store.py           stored workbooks, watchlist, dump-to-dump diff
data/                  uploaded workbooks and watchlist.json
tests/test_targeting.py  the rule engine
tests/test_robustness.py the ways a new dump can differ
```

## Tests

```bash
python -m pytest tests -q
```

124 tests, in two files. `test_targeting.py` asserts the user's two worked examples by name — HEUNG-A YOUNG is a prime candidate,
NINGBO TRADER is excluded for being too soon after drydock — plus the exact expected October list,
both workbook defects above, the date and name parsing, and that loosening a threshold can only
ever add vessels.

It also sweeps the whole settings grid (every drydock threshold × cleaning threshold × hub
selection) and checks three invariants hold throughout: every vessel lands in exactly one bucket,
the buckets always sum to the fleet size, and the funnel can only ever narrow.

`test_robustness.py` mutates the workbook the way a new dump might differ and asserts the
dashboard either still gets the right answer or says loudly that it cannot: the fleet sheet
renamed, port-call columns reordered, a third call block added, header spellings like
`Last DryDocking`, an undecidable date column, missing history sheets, duplicated vessel rows,
a corrupt watchlist, and two uploads in the same second.
