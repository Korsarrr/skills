---
name: cfo-quarterly-report
description: "Generate the polished BlazingCDN QUARTERLY financial report as a formatted Word (.docx) — same house style as the monthly report but quarter-framed, with QoQ + YoY-by-quarter comparisons and the quarter-only metrics promoted to first-class: NRR, GRR, Magic Number, Rule of 40. Use at quarter close once the data workers have returned quarterly aggregates and the CFO has assembled the quarterly input.json. Do NOT use for monthly reports (use cfo-monthly-report) or for the data pull itself (that is the workers' job)."
---

# CFO Quarterly Financial Report

Same deterministic engine and house style as the monthly skill, separate so quarterly logic never loads on a
monthly run (and vice-versa). Layout frozen in `scripts/build.js`; only `input.json` numbers change.

## When to run
At **quarter close**, after the data workers return quarterly aggregates:
- HubSpot Deals Worker → quarterly revenue series, top clients by quarter, quarterly decomposition
  (including `decomposition.starting_recurring` so NRR/GRR can be computed)
- Subscription Tracker Worker → quarter expense lines
- Optionally `inputs.prior_q_sm_usd` (prior-quarter S&M spend) so the Magic Number computes

## What the CFO does
1. Assemble quarterly `input.json` (see `schema/example_input.json` — Q1 2026 worked example) + add `cfo_narrative`.
2. `python3 scripts/make_report.py <input.json> <output.docx>`
3. Review → save to Google Drive "CFO" → post to Hermes (who packages for the CEO / board).

## What differs from monthly
- Trend chart and tables are by **quarter**.
- Section B is **QoQ** (this quarter vs prior quarter) plus **same quarter last year**.
- Metrics table computes **NRR / GRR** (from the decomposition base), **Magic Number** (net-new ARR ÷ prior-Q S&M),
  and frames **Rule of 40** (FCF margin remains an OpEx-only proxy until a full P&L exists).
- ARR is the quarter-exit run-rate (exit MRR × 12).

## Run / files / token economy / Hub-ID note
Identical to the monthly skill (see its SKILL.md). Compute math lives in `scripts/compute.py`;
CFO supplies only `cfo_narrative`.
