#!/usr/bin/env python3
"""
cfo-quarterly-report / compute.py
Input : input.json (quarter-focused; assembled by CFO from the two data workers)
Output: computed.json (same render shape as monthly skill; quarter framing)
NRR/GRR/Magic Number/Rule of 40 are first-class here (quarterly cadence).
"""
import json, sys, statistics

NAVY="16365C"; ORANGE="E8590C"; GREEN="2E7D32"; AMBER="B26A00"; RED="C62828"; GREY="6B7280"
usd=lambda x:f"${x:,.2f}"; usd0=lambda x:f"${x:,.0f}"; usdk=lambda x:f"${x/1000:.1f}K"
pct=lambda x:f"{x*100:.1f}%"
def signed_usd(x): return ("+" if x>=0 else "\u2212")+f"${abs(x):,.2f}"
def signed_pct(x): return ("+" if x>=0 else "\u2212")+f"{abs(x)*100:.1f}%"

def main(inp_path, out_path):
    d=json.load(open(inp_path,encoding="utf-8"))
    meta=d["meta"]; rev=d["revenue"]; exp=d["expenses"]; narr=d.get("cfo_narrative",{})
    fx=meta.get("fx_eur_usd",1.08)

    qseries=rev["quarterly_series"]                       # [{quarter,revenue,count}]
    by_q={q["quarter"]:q for q in qseries}
    focal=rev["focal_quarter"]; fv=by_q[focal]["revenue"]; fcount=by_q[focal].get("count","?")
    zero_q=[q["quarter"] for q in qseries if q["revenue"]==0]
    has_gap=len(zero_q)>0

    # integrity flags
    flags=[]
    for z in zero_q:
        flags.append(f"{z} = $0 in source data \u2014 a missing batch, not a zero-revenue quarter; trailing metrics spanning it are PROVISIONAL. Action: backfill.")
    if not exp.get("full_pl",False):
        flags.append("Expense files are not a full P&L \u2014 SaaS tooling, ads and AI only; no CDN delivery COGS, no payroll. Rule-of-40 FCF margin is therefore an OpEx-only proxy, not GAAP.")
    if meta.get("pre_batch",False):
        flags.append(f"{meta['period_label']} closes on the batch of {meta.get('close_date','the 5th')}; figures are provisional as of {meta.get('date')}. FX EUR\u2192USD {fx}.")
    flags+=narr.get("extra_integrity_flags",[])

    # expenses (quarter total)
    opex_rows=[]; eur_total=0.0
    for ln in exp["lines"]:
        eur=ln["eur"]; eur_total+=eur
        opex_rows.append([f'{ln["category"]} ({ln["detail"]})', f"\u20AC{eur:,.2f}", usd(eur*fx)])
    usd_total=eur_total*fx
    opex_rows.append({"total":True,"cells":["Total tracked OpEx (quarter)", f"\u20AC{eur_total:,.2f}", usd(usd_total)]})
    net=fv-usd_total; contrib=net/fv if fv else 0

    rec=rev["plan_split"]["recurring"]; payg=rev["plan_split"]["payg"]; tot=rec["revenue"]+payg["revenue"]
    plan_rows=[
        [f'Recurring / subscription-like ({rec["clients"]} clients)', usd(rec["revenue"]), pct(rec["revenue"]/tot) if tot else "0%"],
        [f'One-off / PAYG-like ({payg["clients"]} clients)', usd(payg["revenue"]), pct(payg["revenue"]/tot) if tot else "0%"],
    ]

    # QoQ + YoY(quarter)
    quarters=[q["quarter"] for q in qseries]; idx=quarters.index(focal)
    prior=qseries[idx-1]["revenue"] if idx>0 else None
    qoq={"prior":prior,"abs":fv-prior,"pct":(fv-prior)/prior if prior else 0} if prior else None
    # same quarter last year: match "Qn" four positions back if present
    yoy=None
    if idx-4>=0:
        py=qseries[idx-4]["revenue"]; yoy={"prior":py,"abs":fv-py,"pct":(fv-py)/py if py else 0}

    # ARR: quarter-exit run-rate (exit MRR ~ focal/3) * 12
    exit_mrr=fv/3.0; arr_current=exit_mrr*12
    complete=[q["revenue"] for q in qseries if q["revenue"]>0]
    arr_sm=(statistics.mean(complete[-2:])/3*12) if complete else arr_current

    rec_tool_eur=exp.get("recurring_tool_eur", eur_total); tool_usd=rec_tool_eur*fx
    qtr_rev_for_pct=fv  # tool spend over quarter revenue
    tool_pct=(tool_usd)/qtr_rev_for_pct if qtr_rev_for_pct else 0

    conc=rev["concentration"]
    kpis=[
        {"value":usdk(fv),"label":f"Quarter revenue ({focal})","status":f"ARR \u2248 {usdk(arr_current)} run-rate","color":NAVY},
        {"value":signed_pct(qoq["pct"]) if qoq else "n/a","label":"QoQ growth","status":"vs prior quarter","color":GREEN if (qoq and qoq["pct"]>=0) else RED},
        {"value":signed_pct(yoy["pct"]) if yoy else "n/a","label":"YoY (quarter)","status":"vs same Q last year","color":GREEN},
        {"value":pct(conc["top10_share"]),"label":"Top-10 concentration","status":"\u26A0 High risk" if conc["top10_share"]>0.70 else "OK","color":AMBER},
    ]

    top_rows=[]
    for i,c in enumerate(rev["top_clients"],1):
        label=f'{i}  {c["name"]}'+("  \u2605 new" if c.get("new_logo") else "")
        top_rows.append({"new":c.get("new_logo",False),"label":label,"focal":usd(c["focal"]),
            "t3m":usd(c.get("prior_q",0)),"yoy":usd(c.get("yoy_same_q",0)),"yoy_zero":c.get("yoy_same_q",0)==0})

    dc=rev["decomposition"]
    decomp_rows=[
        {"label":"New","color":GREEN,"amt":signed_usd(dc["new"]["amount"]),"n":dc["new"]["clients"]},
        {"label":"Expansion","color":GREEN,"amt":signed_usd(dc["expansion"]["amount"]),"n":dc["expansion"]["clients"]},
        {"label":"Contraction","color":RED,"amt":signed_usd(dc["contraction"]["amount"]),"n":dc["contraction"]["clients"]},
        {"label":"Churn (full logo loss)","color":GREEN if dc["churn"]["clients"]==0 else RED,"amt":signed_usd(dc["churn"]["amount"]),"n":dc["churn"]["clients"]},
    ]
    net_change=dc["new"]["amount"]+dc["expansion"]["amount"]+dc["contraction"]["amount"]+dc["churn"]["amount"]

    # quarterly NRR / GRR from decomposition base (if provided)
    base=dc.get("starting_recurring")
    nrr=grr=None
    if base:
        nrr=(base+dc["expansion"]["amount"]+dc["contraction"]["amount"]+dc["churn"]["amount"])/base
        grr=(base+dc["contraction"]["amount"]+dc["churn"]["amount"])/base
    # Magic Number = net new ARR / prior-Q S&M  (if S&M provided)
    sm_prior=d.get("inputs",{}).get("prior_q_sm_usd")
    net_new_arr=net_change*4  # quarterly net-new annualized
    magic=(net_new_arr/sm_prior) if sm_prior else None

    def st(t,c,bold=False): return {"t":t,"color":c,"bold":bold}
    def stat_nrr():
        if nrr is None: return st("Needs base",AMBER)
        return st("\u2713 Pass" if nrr>=1.10 else "\u26A0 Below",GREEN if nrr>=1.10 else AMBER,True)
    def stat_grr():
        if grr is None: return st("Needs base",AMBER)
        return st("\u2713 Pass" if grr>=0.95 else "\u26A0 Below",GREEN if grr>=0.95 else AMBER,True)
    metrics=[
        ["ARR (quarter-exit)", f"~{usdk(arr_current)}", "10\u00D7 baseline", st("Tracking",GREY)],
        ["Net new ARR (Q, annualized)", signed_usd(net_new_arr), "(10\u00D7\u22121)", st("Tracking",GREY)],
        ["NRR (quarter)", pct(nrr) if nrr is not None else "needs starting base", "\u2265110%", stat_nrr()],
        ["GRR (quarter)", pct(grr) if grr is not None else "needs starting base", "\u226595%", stat_grr()],
        ["Magic Number", (f"{magic:.2f}" if magic is not None else "n/a (needs prior-Q S&M)"), "\u22651.0",
            (st("\u2713 Pass" if (magic and magic>=1) else "\u26A0 Below",GREEN if (magic and magic>=1) else AMBER,True) if magic is not None else st("Needs S&M",AMBER))],
        ["Rule of 40", ((signed_pct(yoy['pct']) if yoy else "?")+" growth + FCF% (proxy)" ), "\u226540", st("Needs full P&L",AMBER)],
        ["CAC by channel", "not attributable yet", "<18mo payback", st("Needs channel tags",AMBER)],
        ["Tool spend / revenue", pct(tool_pct), "<5%", st("\u2713 Pass",GREEN,True) if tool_pct<0.05 else st("\u26A0 Over",AMBER,True)],
        ["Concentration top-10", pct(conc["top10_share"]), "(watch)",
            st("\u26A0 High",AMBER,True) if conc["top10_share"]>0.70 else st("OK",GREEN,True)],
    ]

    # T3M-equivalent block becomes QoQ table
    qoq_rows=[
        [f"{focal} (this quarter)", usd(fv), ""],
        [f"{quarters[idx-1] if idx>0 else 'Prior Q'} (prior quarter)", usd(prior) if prior else "n/a",
            {"t":signed_pct(qoq["pct"]),"color":RED if (qoq and qoq["pct"]<0) else GREEN} if qoq else ""],
    ]
    if yoy:
        qoq_rows.append({"total":True,"cells":[f"{quarters[idx-4]} (same Q last year)", usd(yoy["prior"]),
            {"t":signed_pct(yoy["pct"]),"color":GREEN if yoy["pct"]>=0 else RED}]})

    computed={
      "meta":meta,"kpis":kpis,"integrity_flags":flags,"bottom_line":narr.get("bottom_line",[]),
      "section_a":{
        "revenue_headline":usd(fv),
        "revenue_sub":f"{fcount} billing events  \u2022  {rev.get('unique_clients','?')} unique paying clients (quarter)",
        "plan_rows":plan_rows,"plan_note":narr.get("plan_note",""),
        "opex_rows":opex_rows,
        "net_eq":f"{usd(fv)} \u2212 {usd(usd_total)} = ","net_val":usd(net),
        "net_sub":f"{pct(contrib)} contribution after tracked tooling/marketing (quarter)",
        "net_caveat":"Not a gross/operating margin \u2014 CDN delivery COGS and payroll are absent from the data room.",
      },
      "section_b":{
        "yoy_pre":(f"{usd(fv)} vs {usd(qoq['prior'])}   \u2192   " if qoq else ""),
        "yoy_val":(f"{signed_usd(qoq['abs'])}  /  {signed_pct(qoq['pct'])} QoQ" if qoq else "n/a"),
        "t3m_rows":qoq_rows,
        "t3m_note":narr.get("qoq_note",""),
        "decomp_rows":decomp_rows,"decomp_net":signed_usd(net_change),
        "decomp_material":dc.get("material_clients",len(rev["top_clients"])),
        "decomp_window":dc.get("window",f"{focal} vs prior quarter"),
        "decomp_note":narr.get("decomp_note","This decomposition feeds the quarterly NRR / GRR in the metrics table."),
      },
      "section_c":{
        "top_rows":top_rows,"top_note":narr.get("top_note",""),
        "conc_pre":f"Top-10 share {pct(conc['top10_share'])}",
        "conc_post":f"  \u2022  Top-1 share {pct(conc['top1_share'])} ({conc.get('top1_name','')})",
        "conc_reco":narr.get("concentration_note","Loss of any single top account is material. Recommend a named-account QBR program and a diversification target (top-10 < 70%)."),
        "hygiene":narr.get("hygiene_note",""),
      },
      "section_d":{
        "arr_rows":[
          ["Quarter-exit run-rate", usd0(exit_mrr), "\u2248"+usdk(arr_current)],
          {"total":True,"cells":["Smoothed (recent complete quarters)", usd0(arr_sm/12), "\u2248"+usdk(arr_sm)]},
        ],
        "arr_reco":narr.get("arr_recommendation",f"Plan against ARR \u2248 {usdk(arr_current)} (quarter-exit run-rate)."),
        "pipeline_note":narr.get("pipeline_note","Pipeline-to-ARR: not computable until a sales pipeline with open stages exists."),
      },
      "section_e":{
        "cac_note":narr.get("cac_note","CAC by channel: not a real CAC until new-logo counts are channel-tagged and fully-loaded S&M cost is captured."),
        "toolspend_note":narr.get("toolspend_note",f"Tool spend vs quarter revenue: \u2248 {pct(tool_pct)} \u2014 {'under' if tool_pct<0.05 else 'over'} the <5% target."),
      },
      "section_f":{
        "runway_title":"Cash runway \u2014 not computable this cycle",
        "runway_lines":narr.get("runway_lines",[
          "Missing inputs: opening cash balance and fully-loaded burn (payroll + CDN infrastructure COGS).",
          f"On tracked OpEx alone the quarter is cash-generative ({signed_usd(net)}), but that excludes people and delivery infra \u2014 NOT a runway statement.",
        ]),
      },
      "metrics_rows":metrics,
      "chart_data":{
        "series":[{"label":q["quarter"],"value":q["revenue"]} for q in qseries],
        "record_month":max(qseries,key=lambda q:q["revenue"])["quarter"],
        "gap_months":zero_q,"focal_month":focal,
        "top_clients":[{"name":c["name"],"value":c["focal"],"new":c.get("new_logo",False)} for c in rev["top_clients"]],
        "concentration":{"top1":conc["top1_share"]*100,"top2_10":(conc["top10_share"]-conc["top1_share"])*100,
                         "rest":(1-conc["top10_share"])*100,"headline":pct(conc["top10_share"])},
        "quarterly":True,
      },
    }
    json.dump(computed,open(out_path,"w",encoding="utf-8"),ensure_ascii=False,indent=2)
    print(f"quarterly computed.json written (NRR={pct(nrr) if nrr is not None else 'n/a'}, GRR={pct(grr) if grr is not None else 'n/a'})")

if __name__=="__main__":
    main(sys.argv[1] if len(sys.argv)>1 else "input.json",
         sys.argv[2] if len(sys.argv)>2 else "computed.json")
