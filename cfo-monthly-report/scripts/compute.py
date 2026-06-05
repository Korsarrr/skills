#!/usr/bin/env python3
"""
cfo-monthly-report / compute.py
Input : input.json  (assembled by CFO from HubSpot Deals Worker + Subscription Tracker Worker)
Output: computed.json (render-ready: numbers formatted, statuses + colors derived, integrity flags auto-detected)
Pure math + formatting. No tool calls. Deterministic.
"""
import json, sys, statistics

# palette (must match build.js)
NAVY="16365C"; ORANGE="E8590C"; GREEN="2E7D32"; AMBER="B26A00"; RED="C62828"; GREY="6B7280"

def usd(x):  return f"${x:,.2f}"
def usd0(x): return f"${x:,.0f}"
def usdk(x): return f"${x/1000:.1f}K"
def pct(x):  return f"{x*100:.1f}%"
def signed_usd(x): return ("+" if x>=0 else "\u2212") + f"${abs(x):,.2f}"
def signed_pct(x): return ("+" if x>=0 else "\u2212") + f"{abs(x)*100:.1f}%"

def main(inp_path, out_path):
    d = json.load(open(inp_path, encoding="utf-8"))
    meta = d["meta"]; rev = d["revenue"]; exp = d["expenses"]
    narr = d.get("cfo_narrative", {})
    fx = meta.get("fx_eur_usd", 1.08)

    series = rev["monthly_series"]                       # list of {month,count,revenue}
    by_month = {m["month"]: m for m in series}
    focal = rev["focal_month"]
    fv = by_month[focal]["revenue"]
    fcount = by_month[focal]["count"]

    # ---- integrity flags (auto-detect zero-revenue months in series) ----
    zero_months = [m["month"] for m in series if m["revenue"] == 0]
    flags = []
    for z in zero_months:
        flags.append(f"{z} = $0 in source data \u2014 almost certainly a missing billing batch, not a zero-revenue month. "
                     "Every trailing metric spanning it (T3M, TTM, NRR, GRR, decomposition) is PROVISIONAL. "
                     "Clean single-month and YoY figures are unaffected. Action: backfill the batch (owner: data/RevOps).")
    if not exp.get("full_pl", False):
        flags.append("Expense files are not a full P&L \u2014 they cover SaaS tooling, ads and AI-API only; no CDN delivery COGS, "
                     "no payroll. True gross/operating margin and runway cannot be computed yet. Margins below are "
                     "\u201Cnet of tracked OpEx,\u201D not GAAP.")
    if meta.get("pre_batch", False):
        flags.append(f"{meta['period_label']} is pre-batch / provisional \u2014 formal close is {meta.get('close_date','the 5th')}; "
                     f"figures reflect deal data present as of {meta.get('date')}. "
                     f"FX: expenses in EUR converted at assumed EUR\u2192USD {fx} (confirm settlement rate).")
    flags += narr.get("extra_integrity_flags", [])
    has_gap = len(zero_months) > 0

    # ---- expenses ----
    opex_rows=[]; eur_total=0.0
    for ln in exp["lines"]:
        eur=ln["eur"]; eur_total+=eur
        opex_rows.append([f'{ln["category"]} ({ln["detail"]})', f"\u20AC{eur:,.2f}", usd(eur*fx)])
    usd_total = eur_total*fx
    opex_rows.append({"total":True,"cells":["Total tracked OpEx", f"\u20AC{eur_total:,.2f}", usd(usd_total)]})

    net = fv - usd_total
    contrib = net/fv if fv else 0

    # ---- plan split ----
    rec = rev["plan_split"]["recurring"]; payg = rev["plan_split"]["payg"]
    tot_split = rec["revenue"]+payg["revenue"]
    plan_rows=[
        [f'Recurring / subscription-like ({rec["clients"]} clients)', usd(rec["revenue"]),
         pct(rec["revenue"]/tot_split) if tot_split else "0%"],
        [f'One-off / PAYG-like ({payg["clients"]} clients)', usd(payg["revenue"]),
         pct(payg["revenue"]/tot_split) if tot_split else "0%"],
    ]

    # ---- YoY (same calendar month, prior year) ----
    y,mn = focal.split("-"); prior_my = f"{int(y)-1}-{mn}"
    yoy=None
    if prior_my in by_month:
        pv=by_month[prior_my]["revenue"]; yoy={"prior":pv,"abs":fv-pv,"pct":(fv-pv)/pv if pv else 0}

    # ---- T3M vs prior T3M ----
    months=[m["month"] for m in series]
    idx=months.index(focal)
    t3m_window=series[idx-2:idx+1]
    prior_window=series[idx-5:idx-2]
    t3m=sum(m["revenue"] for m in t3m_window)
    pt3m=sum(m["revenue"] for m in prior_window)
    t3m_change=(t3m-pt3m)/pt3m if pt3m else 0
    # backfill estimate if a zero month sits in the t3m window
    complete=[m["revenue"] for m in series if m["revenue"]>0]
    trailing_avg=statistics.mean(complete[-6:]) if complete else 0
    gap_in_window=any(m["revenue"]==0 for m in t3m_window)
    t3m_bf = t3m + sum(trailing_avg for m in t3m_window if m["revenue"]==0)
    t3m_bf_change=(t3m_bf-pt3m)/pt3m if pt3m else 0

    # ---- ARR ----
    arr_current=fv*12
    smoothed_months=[m["revenue"] for m in series[max(0,idx-3):idx+1] if m["revenue"]>0]
    mrr_sm=statistics.mean(smoothed_months) if smoothed_months else fv
    arr_sm=mrr_sm*12

    # ---- tool spend % ----
    rec_tool_eur=exp.get("recurring_tool_eur", eur_total)
    tool_usd=rec_tool_eur*fx
    tool_pct=tool_usd/fv if fv else 0

    # ---- KPI cards ----
    kpis=[
        {"value":usdk(fv),"label":f"MRR ({'provisional' if meta.get('pre_batch') else 'actual'})",
         "status":f"ARR \u2248 {usdk(arr_sm)} smoothed","color":NAVY},
        {"value":signed_pct(yoy["pct"]) if yoy else "n/a","label":f"Revenue YoY ({meta['period_label'].split()[0]})",
         "status":"Clean \u2014 trust","color":GREEN},
        {"value":pct(tool_pct),"label":"Tool spend / revenue",
         "status":("\u2713 Under 5% target" if tool_pct<0.05 else "\u26A0 Over 5% target"),
         "color":GREEN if tool_pct<0.05 else AMBER},
        {"value":pct(rev["concentration"]["top10_share"]),"label":"Top-10 concentration",
         "status":"\u26A0 High risk" if rev["concentration"]["top10_share"]>0.70 else "OK","color":AMBER},
    ]

    # ---- top clients (Top-N: this mo / prev complete mo / MoM / T3M / YoY same mo) ----
    prev_literal = months[idx-1] if idx>0 else None
    prev_is_gap = (prev_literal in zero_months) if prev_literal else False
    top_prev_note = ("* Prev mo. = previous complete billed month; "
                     f"{prev_literal} was skipped (missing batch).") if prev_is_gap else ""
    top_rows=[]
    for i,c in enumerate(rev["top_clients"],1):
        label=f'{i}  {c["name"]}' + ("  \u2605 new" if c.get("new_logo") else "")
        pm=c.get("prev_month",0) or 0
        if pm>0:
            mom_v=(c["focal"]-pm)/pm
            mom=signed_pct(mom_v); mom_color=GREEN if mom_v>=0 else RED
            prev_disp=usd(pm); prev_color=None
        else:
            mom="\u2014"; mom_color=GREY; prev_disp="\u2014"; prev_color=GREY
        top_rows.append({"new":c.get("new_logo",False),"label":label,
            "this":usd(c["focal"]),"prev":prev_disp,"prev_color":prev_color,
            "mom":mom,"mom_color":mom_color,"t3m":usd(c["t3m"]),
            "yoy":usd(c.get("yoy_same_month",0)),"yoy_color":GREY if c.get("yoy_same_month",0)==0 else None})

    # ---- decomposition ----
    dc=rev["decomposition"]
    decomp_rows=[
        {"label":"New","color":GREEN,"amt":signed_usd(dc["new"]["amount"]),"n":dc["new"]["clients"]},
        {"label":"Expansion","color":GREEN,"amt":signed_usd(dc["expansion"]["amount"]),"n":dc["expansion"]["clients"]},
        {"label":"Contraction","color":RED,"amt":signed_usd(dc["contraction"]["amount"]),"n":dc["contraction"]["clients"]},
        {"label":"Churn (full logo loss)","color":GREEN if dc["churn"]["clients"]==0 else RED,
         "amt":signed_usd(dc["churn"]["amount"]),"n":dc["churn"]["clients"]},
    ]
    net_change=dc["new"]["amount"]+dc["expansion"]["amount"]+dc["contraction"]["amount"]+dc["churn"]["amount"]

    # ---- SaaS metrics statuses ----
    def st(text,color,bold=False): return {"t":text,"color":color,"bold":bold}
    gapamber = lambda t: st(t, AMBER)
    metrics=[
        ["MRR", f"{usdk(fv)} prov. / {usdk(mrr_sm)} smoothed", "ARR/12", st("Tracking",GREY)],
        ["ARR", f"~{usdk(arr_current)} floor / ~{usdk(arr_sm)} smoothed", "10\u00D7 baseline", st("Tracking",GREY)],
        ["Net new ARR (mo)", "provisional (decomposition)", "(10\u00D7\u22121)/12",
            st("Blocked \u2014 gap",RED) if has_gap else st("Tracking",GREY)],
        ["NRR (T3M proxy)", narr.get("nrr_text","see decomposition")+("  \u26A0 distorted" if has_gap else ""),
            "\u2265110%", gapamber("Re-run post-backfill") if has_gap else st("Tracking",GREY)],
        ["GRR (T3M proxy)", narr.get("grr_text","see decomposition")+("  \u26A0 distorted" if has_gap else ""),
            "\u226595%", gapamber("Re-run post-backfill") if has_gap else st("Tracking",GREY)],
        ["CAC by channel", "not attributable yet", "<18mo payback", gapamber("Needs channel tags")],
        ["Magic Number", "n/a (needs prior-Q S&M)", "\u22651.0", gapamber("Needs S&M base")],
        ["Rule of 40", "n/a (needs FCF margin)", "\u226540", gapamber("Needs full P&L")],
        ["Tool spend / revenue", pct(tool_pct), "<5%",
            st("\u2713 Pass",GREEN,True) if tool_pct<0.05 else st("\u26A0 Over",AMBER,True)],
        ["Concentration top-10", pct(rev["concentration"]["top10_share"]), "(watch)",
            st("\u26A0 High",AMBER,True) if rev["concentration"]["top10_share"]>0.70 else st("OK",GREEN,True)],
    ]

    conc=rev["concentration"]
    computed={
        "meta":meta,
        "kpis":kpis,
        "integrity_flags":flags,
        "bottom_line":narr.get("bottom_line",[]),
        "section_a":{
            "revenue_headline":usd(fv),
            "revenue_sub":f"{fcount} billing events  \u2022  {rev.get('unique_clients','?')} unique paying clients",
            "plan_rows":plan_rows,
            "plan_note":narr.get("plan_note","Recommendation: add a plan_type deal property so the split is exact, not heuristic."),
            "opex_rows":opex_rows,
            "net_eq":f"{usd(fv)} \u2212 {usd(usd_total)} = ",
            "net_val":usd(net),
            "net_sub":f"{pct(contrib)} contribution after tracked tooling/marketing",
            "net_caveat":"This is not a gross or operating margin \u2014 CDN delivery COGS and payroll are absent from the data room.",
        },
        "section_b":{
            "yoy_pre":(f"{usd(fv)} vs {usd(yoy['prior'])}   \u2192   " if yoy else ""),
            "yoy_val":(f"{signed_usd(yoy['abs'])}  /  {signed_pct(yoy['pct'])}" if yoy else "n/a"),
            "t3m_rows":[
                [f"T3M ({' + '.join(m['month'][-2:]+f'={usdk(m['revenue'])}' for m in t3m_window)})", usd(t3m), ""],
                ["Prior T3M", usd(pt3m), {"t":signed_pct(t3m_change),"color":RED if t3m_change<0 else GREEN}],
            ] + ([{"total":True,"cells":["With backfill (\u2248 trailing avg)", "\u2248"+usdk(t3m_bf),
                    {"t":"\u2248"+signed_pct(t3m_bf_change),"color":GREEN if t3m_bf_change>=0 else RED}]}] if gap_in_window else []),
            "t3m_note":narr.get("t3m_note","The headline change is an artifact of the missing batch \u2014 do not act on it until the gap is loaded." if gap_in_window else ""),
            "decomp_rows":decomp_rows,
            "decomp_net":signed_usd(net_change),
            "decomp_material":dc.get("material_clients",len(rev["top_clients"])),
            "decomp_window":dc.get("window",""),
            "decomp_note":narr.get("decomp_note","Zero churn is the real signal \u2014 no clients fully lost. Re-run after backfill for a true decomposition."),
        },
        "section_c":{
            "top_rows":top_rows,
            "top_prev_note":top_prev_note,
            "top_note":narr.get("top_note",""),
            "conc_pre":f"Top-10 share {pct(conc['top10_share'])}",
            "conc_post":f"  \u2022  Top-1 share {pct(conc['top1_share'])} ({conc.get('top1_name','')})",
            "conc_reco":narr.get("concentration_note",
                "Loss of any single top account is material. Recommend a named-account QBR/retention program for the top 10 and a diversification target (top-10 < 70% by year-end)."),
            "hygiene":narr.get("hygiene_note",""),
        },
        "section_d":{
            "arr_rows":[
                ["Current (provisional)", usd0(fv), "\u2248"+usdk(arr_current)],
                {"total":True,"cells":["Smoothed (complete months, excl. gap)", usd0(mrr_sm), "\u2248"+usdk(arr_sm)]},
            ],
            "arr_reco":narr.get("arr_recommendation",
                f"Plan against ARR \u2248 {usdk(arr_sm)} (smoothed working run-rate); {usdk(arr_current)} is the conservative floor until the gap is backfilled and the month reconciled."),
            "pipeline_note":narr.get("pipeline_note",
                "Pipeline-to-ARR conversion: not computable \u2014 \u201CCurrent Clients Pipeline\u201D is a single-stage customer ledger, not a funnel. Action: stand up a sales pipeline with open stages (Lead \u2192 Qualified \u2192 Proposal \u2192 Won) to measure coverage/conversion."),
        },
        "section_e":{
            "cac_note":narr.get("cac_note",
                "CAC by channel: only paid ad spend is attributable. Not a real CAC until new-logo counts are channel-tagged and fully-loaded S&M cost is captured. Action: tag deals with acquisition channel; route S&M payroll into the data room."),
            "toolspend_note":narr.get("toolspend_note",
                f"Tool spend vs revenue: recurring tooling \u2248 {usd0(tool_usd)}/mo = {pct(tool_pct)} of MRR \u2014 under the <5% target."),
        },
        "section_f":{
            "runway_title":"Cash runway \u2014 not computable this cycle",
            "runway_lines":narr.get("runway_lines",[
                "Missing inputs: opening cash balance (held by YODA / banking) and fully-loaded monthly burn (payroll + CDN infrastructure COGS).",
                f"On tracked OpEx alone the business is cash-generative ({signed_usd(net)}), but that ignores the two largest real cost buckets (people, delivery infra) \u2014 so it is NOT a runway statement. Action: provide opening cash + payroll + infra COGS next cycle.",
            ]),
        },
        "metrics_rows":metrics,
        "chart_data":{
            "series":[{"label":m["month"],"value":m["revenue"]} for m in series],
            "record_month":max(series,key=lambda m:m["revenue"])["month"],
            "gap_months":zero_months,
            "focal_month":focal,
            "top_clients":[{"name":c["name"],"value":c["focal"],"new":c.get("new_logo",False)} for c in rev["top_clients"]],
            "concentration":{"top1":conc["top1_share"]*100,
                             "top2_10":(conc["top10_share"]-conc["top1_share"])*100,
                             "rest":(1-conc["top10_share"])*100,
                             "headline":pct(conc["top10_share"])},
        },
    }
    json.dump(computed, open(out_path,"w",encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"computed.json written ({len(flags)} integrity flags, {'GAP detected' if has_gap else 'no gap'})")

if __name__=="__main__":
    main(sys.argv[1] if len(sys.argv)>1 else "input.json",
         sys.argv[2] if len(sys.argv)>2 else "computed.json")
