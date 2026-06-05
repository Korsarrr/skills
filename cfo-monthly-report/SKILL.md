---
name: cfo-monthly-report
description: "Generate the polished BlazingCDN MONTHLY financial report as a formatted Word (.docx) — title page, KPI cards, data-integrity callouts, formatted tables with status colors, and three charts (revenue trend, top-10 clients, concentration donut). Use on the 5th-of-month financial cycle once the HubSpot Deals Worker (revenue JSON) and Subscription Tracker Worker (expense JSON) have returned data and the CFO has assembled input.json. Covers sections A–F + SaaS metrics summary. Do NOT use for quarterly reports (use cfo-quarterly-report) or for the data pull itself (that is the workers' job)."
---

# CFO Monthly Financial Report

Deterministic render: same code → identical design every month. Only the numbers in `input.json` change.
The layout is frozen in `scripts/build.js`; do not re-derive it. `assets/template_reference.docx` is a visual
reference only — it is not the mechanism.

## When to run
On the **5th-of-month** financial cycle, AFTER the billing batch has loaded and both data workers have returned:
- HubSpot Deals Worker → revenue data
- Subscription Tracker Worker → expense data

## What the CFO does (the only manual part)
1. Assemble `input.json` from the two workers' JSON + add the qualitative `cfo_narrative`
   (bottom_line, notes, flags). See `schema/example_input.json` for the exact contract.
   - Revenue numbers come from the HubSpot Deals Worker; expense lines from the Subscription Tracker Worker.
   - Set `meta.source_portal` to the confirmed **production** HubSpot Hub ID (see note below).
2. Run the skill (one command, below).
3. Review the output, then save to Google Drive "CFO" and post the link to Hermes.

The CFO writes only judgement (the `cfo_narrative` block). All math (ARR/MRR, YoY, T3M + backfill estimate,
net result, margins, tool-spend %, concentration, integrity-flag auto-detection, metric statuses/colors) is done
by `compute.py` — single source of truth, matching the CFO metric definitions.

## Run
```bash
python3 scripts/make_report.py <input.json> <output.docx>
```
Pipeline: `compute.py` (input.json → computed.json) → `charts.py` (3 PNGs) → `build.js` (render) → docx validate.
Requires `matplotlib` (python) and `docx` (node) — both present in the agent env; export
`NODE_PATH=$(npm root -g)` is handled by the orchestrator.

## Token economy
- Workers do the heavy data pull and return compact JSON; raw deal rows never enter the CFO's context.
- Charts and tables are built by code, not by model reasoning — no per-month re-derivation of layout.
- This skill loads only for monthly reports; quarterly logic lives in a separate skill so it never bloats context.

## Files
- `scripts/compute.py` — metric math + formatting + integrity-flag detection → computed.json
- `scripts/charts.py` — 3 charts from computed.json
- `scripts/build.js` — Word renderer (frozen layout)
- `scripts/make_report.py` — orchestrator + validation
- `schema/example_input.json` — the input contract (May 2026 worked example)
- `assets/template_reference.docx` — visual reference only

## Section C — Top 20 customers
The customer table lists the **top 20 clients** with four comparisons per client so new logos and patterns
surface at a glance: **This mo · Prev mo (previous complete billed month) · MoM% (colored \u25B2green/\u25BCred) · T3M ·
Same mo last year**. Net-new logos (YoY) are flagged \u2605 and colored. If the literal previous month is a gap,
the Prev/MoM cells read "\u2014" and a footnote names the skipped month. **Concentration stays Top-10**
(top-10 / top-1 share, donut) \u2014 that risk metric is unchanged.

## Integrity behaviour (automatic)
- Any $0 month in the series is auto-flagged as a probable missing batch; all trailing metrics spanning it
  are marked PROVISIONAL.
- If `expenses.full_pl` is false, margins are labelled "net of tracked OpEx," not GAAP.
- If `meta.pre_batch` is true, the report is labelled provisional with the close date.

## NOTE — production Hub ID
The example uses a `<CONFIRM_PROD_HUB_ID>` placeholder. There is an open ambiguity in the fleet
(143144902 vs 145006611). Confirm the single canonical **production** Hub ID before running, and put it in
`meta.source_portal`. A report keyed to the wrong portal is a data-integrity failure.
