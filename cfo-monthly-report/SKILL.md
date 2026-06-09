---
name: cfo-monthly-report
description: "Generate BlazingCDN's monthly CFO financial report as a polished, Google-Docs-clean .docx via docx-js. Use whenever the monthly CFO report is requested or scheduled. The engine does both the compute (ARR/MRR, YoY, T3M + backfill, decomposition, concentration, integrity flags, metric statuses) and the rendering; you supply only input.json (HubSpot Deals Worker revenue + Subscription Tracker Worker expenses + cfo_narrative judgement). Deterministic: same engine -> identical design every month; only the numbers change."
---

# CFO Monthly Financial Report (single-file .md skill)

This skill builds BlazingCDN's monthly CFO report as a polished `.docx` using **docx-js** (the `docx` skill's library). The engine does both the math and the rendering. Charts are drawn as shaded / block-bar (`█`/`░`) table cells — no images — so the file stays light and imports cleanly into Google Docs. Deterministic: same engine → identical design every month; only the numbers change.

The output follows the `docx` skill's Google-Docs-safe rules (DXA widths, `ShadingType.CLEAR`, dual-width tables, numbering-config lists, no table-rule dividers, `PageNumber` footer fields), so it opens/converts to a native Google Doc without the artefacts of hand-rolled OOXML.

> **Dependency note (changed from the old engine):** this is no longer a zero-dependency stdlib engine. It depends on the `docx` npm package and — optionally — the `docx` skill's `validate.py`. That is the deliberate trade for clean Google Docs output.

## How to produce a report (runtime steps)

1. **Materialize the engine.** Copy the entire `build.js` code block (under "Engine" below) verbatim into a file named `build.js` in your workspace.
2. **Install docx once.** In that folder run `npm install docx` (Node v18+). No other packages are needed.
3. **Assemble the input.** Build `input.json` from the two data workers' JSON results + your own `cfo_narrative` (`bottom_line`, notes). Use the "Example input" block below as the exact contract. Revenue numbers come from the **HubSpot Deals Worker**; expense lines from the **Subscription Tracker Worker**. You write only judgement — all math (ARR/MRR, YoY, T3M + backfill, net result, margins, tool-spend %, per-client MoM, concentration, integrity-flag detection, metric statuses) is done by the engine.
4. **Run:** `node build.js input.json "BlazingCDN_Monthly_Financial_Report_<Month_Year>.docx"`.
5. **Validate (optional but recommended):** `python <docx-skill>/scripts/office/validate.py "<output>.docx"`. If it fails, unpack → fix XML → repack per the `docx` skill.
6. **Review, save to Google Drive "CFO".** Because the `.docx` is Google-Docs-clean, you can also upload it to Drive **with conversion** (`mimeType: application/vnd.google-apps.document`) to land a native Google Doc; the `PageNumber` footer is the only element that may degrade on conversion.

## Section C — Top 20 customers

Top-20 table with comparisons per client so new logos and patterns pop: This mo · Prev mo (previous complete billed month) · MoM% (green ▲ / red ▼) · T3M · Same mo last year. Net-new logos (YoY) are flagged ★ and coloured; the "This mo" cell carries a small inline bar. If the literal previous month is a gap, Prev/MoM read "—" and a footnote names the skipped month. Concentration stays Top-10 (Top-1 / Top 2–10 / Rest stacked bar).

## Integrity behaviour (automatic)

- Any `$0` month in the series is auto-flagged as a probable missing batch; trailing metrics spanning it are **PROVISIONAL**.
- If `expenses.full_pl` is false, margins are labelled "net of tracked OpEx," not GAAP.
- If `meta.pre_batch` is true, the report is labelled provisional with the close date.

## Design behaviour (automatic)

- Signed deltas carry arrows: positive → green ▲, negative → red ▼, exact zero → neutral (no arrow). Applied to KPI YoY, per-client MoM, trend MoM, T3M change, and growth-decomposition amounts. Narrative prose (e.g. the cash-generative figure in Section F) is left arrow-free on purpose.
- Footer shows **Page X of Y** via docx-js `PageNumber.CURRENT` / `PageNumber.TOTAL_PAGES` fields.
- The full visual spec lives in the engine itself — see the `DESIGN INVARIANTS` comment and the palette constants at the top of `build.js` (NAVY / ORANGE / GREEN / RED, image-free, deterministic). There is intentionally no separate design-spec document to drift out of sync.

## NOTE — production Hub ID

Canonical production Hub ID is **143144902** and is baked into `meta.source_portal` in the example below (the prior 143144902 vs 145006611 ambiguity is resolved — 145006611 was the non-production portal). Confirm it still matches your live portal before each run; change the single value in `meta.source_portal` if the portal ever moves.

## Engine — build.js (copy verbatim into a file named build.js)

```javascript
#!/usr/bin/env node
/*
 * cfo-monthly-report / build.js  — renders via docx-js (the `docx` skill).
 * Replaces the hand-rolled OOXML engine. Output is a clean, validated .docx that
 * imports into Google Docs without the artefacts of raw field codes / fixed OOXML.
 *
 *   npm install docx          (once, in this folder)
 *   node build.js <input.json> <output.docx>
 *
 * Same input.json CONTRACT as before — only rendering changed. All math
 * (ARR/MRR, YoY, T3M + backfill, decomposition, concentration, integrity flags,
 * metric statuses) is still done here in compute(); you write only judgement.
 *
 * DESIGN INVARIANTS (unchanged):
 *   palette : NAVY 16365C (primary) · ORANGE E8590C (accent) · GREEN 2E7D32 / RED C62828 (deltas)
 *   deltas  : green up-triangle positive, red down-triangle negative, zero = neutral (no arrow)
 *   charts  : drawn as shaded / block-bar table cells (no images) — survives Google Docs import
 *   footer  : "Page X of Y" via PageNumber fields (docx-js) — preserved
 */
const fs = require('node:fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber
} = require('docx');

/* ============================ palette ============================ */
const NAVY="16365C",ORANGE="E8590C",GREEN="2E7D32",AMBER="B26A00",RED="C62828",GREY="6B7280",
  LIGHT="EEF2F6",AMBERBG="FCF3E2",GREENBG="E9F3EA",GREYBG="F1F2F4",WHITE="FFFFFF",TRACK="DCE2E8";
const CW=9360;

/* ============================ compute (unchanged logic) ============================ */
const usd=x=>`$${x.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const usd0=x=>`$${Math.round(x).toLocaleString('en-US')}`;
const usdk=x=>`$${(x/1000).toFixed(1)}K`;
const pct=x=>`${(x*100).toFixed(1)}%`;
const sUsd=x=>(x>=0?"+":"\u2212")+`$${Math.abs(x).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const sPct=x=>(x>=0?"+":"\u2212")+`${Math.abs(x*100).toFixed(1)}%`;
const arrow=x=>x>0?"\u25B2 ":(x<0?"\u25BC ":"");
const dUsd=x=>arrow(x)+sUsd(x);
const dPct=x=>arrow(x)+sPct(x);
const shortM=m=>{const[a,b]=String(m).split("-");if(!b)return m;const M=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];return `${M[+b-1]}'${a.slice(2)}`;};

function compute(d){
  const meta=d.meta,rev=d.revenue,exp=d.expenses,narr=d.cfo_narrative||{},fx=meta.fx_eur_usd??1.08;
  const series=rev.monthly_series,byM=Object.fromEntries(series.map(m=>[m.month,m]));
  const months=series.map(m=>m.month),focal=rev.focal_month,idx=months.indexOf(focal);
  const fv=byM[focal].revenue,fcount=byM[focal].count;
  const zero=series.filter(m=>m.revenue===0).map(m=>m.month),hasGap=zero.length>0;

  const flags=[];
  for(const z of zero)flags.push(`${z} = $0 in source data — almost certainly a missing billing batch, not a zero-revenue month. Every trailing metric spanning it (T3M, TTM, NRR, GRR, decomposition) is PROVISIONAL. Clean single-month and YoY figures are unaffected. Action: backfill the batch (owner: data/RevOps).`);
  if(!exp.full_pl)flags.push(`Expense files are not a full P&L — they cover SaaS tooling, ads and AI-API only; no CDN delivery COGS, no payroll. True gross/operating margin and runway cannot be computed yet. Margins below are "net of tracked OpEx," not GAAP.`);
  if(meta.pre_batch)flags.push(`${meta.period_label} is pre-batch / provisional — formal close is ${meta.close_date||'the 5th'}; figures reflect deal data present as of ${meta.date}. FX: expenses in EUR converted at assumed EUR→USD ${fx} (confirm settlement rate).`);
  (narr.extra_integrity_flags||[]).forEach(f=>flags.push(f));

  let eurTot=0; const opexRows=[];
  for(const ln of exp.lines){eurTot+=ln.eur;opexRows.push([`${ln.category} (${ln.detail})`,`€${ln.eur.toLocaleString('en-US',{minimumFractionDigits:2})}`,usd(ln.eur*fx)]);}
  const usdTot=eurTot*fx; opexRows.push({total:true,cells:["Total tracked OpEx",`€${eurTot.toLocaleString('en-US',{minimumFractionDigits:2})}`,usd(usdTot)]});
  const net=fv-usdTot, contrib=fv?net/fv:0;

  const rec=rev.plan_split.recurring,payg=rev.plan_split.payg,tot=rec.revenue+payg.revenue;
  const planRows=[[`Recurring / subscription-like (${rec.clients} clients)`,usd(rec.revenue),tot?pct(rec.revenue/tot):"0%"],
                  [`One-off / PAYG-like (${payg.clients} clients)`,usd(payg.revenue),tot?pct(payg.revenue/tot):"0%"]];

  const[y,mn]=focal.split("-"),priorMY=`${+y-1}-${mn}`;
  let yoy=null; if(byM[priorMY]){const pv=byM[priorMY].revenue;yoy={prior:pv,abs:fv-pv,pct:pv?(fv-pv)/pv:0};}

  const win=series.slice(idx-2,idx+1),pwin=series.slice(idx-5,idx-2);
  const t3m=win.reduce((s,m)=>s+m.revenue,0),pt3m=pwin.reduce((s,m)=>s+m.revenue,0);
  const t3mCh=pt3m?(t3m-pt3m)/pt3m:0;
  const complete=series.filter(m=>m.revenue>0).map(m=>m.revenue);
  const trailAvg=complete.slice(-6).reduce((a,b)=>a+b,0)/Math.max(1,complete.slice(-6).length);
  const gapInWin=win.some(m=>m.revenue===0);
  const t3mBf=t3m+win.filter(m=>m.revenue===0).length*trailAvg, t3mBfCh=pt3m?(t3mBf-pt3m)/pt3m:0;

  const arrCur=fv*12;
  const sm=series.slice(Math.max(0,idx-3),idx+1).filter(m=>m.revenue>0).map(m=>m.revenue);
  const mrrSm=sm.length?sm.reduce((a,b)=>a+b,0)/sm.length:fv, arrSm=mrrSm*12;

  const recToolEur=exp.recurring_tool_eur??eurTot, toolUsd=recToolEur*fx, toolPct=fv?toolUsd/fv:0;
  const conc=rev.concentration;

  const kpis=[
    {value:usdk(fv),label:`MRR (${meta.pre_batch?'provisional':'actual'})`,status:`ARR ≈ ${usdk(arrSm)} smoothed`,color:NAVY},
    {value:yoy?dPct(yoy.pct):"n/a",label:`Revenue YoY (${meta.period_label.split(' ')[0]})`,status:"Clean — trust",color:GREEN},
    {value:pct(toolPct),label:"Tool spend / revenue",status:toolPct<0.05?"\u2713 Under 5% target":"\u26A0 Over 5% target",color:toolPct<0.05?GREEN:AMBER},
    {value:pct(conc.top10_share),label:"Top-10 concentration",status:conc.top10_share>0.70?"\u26A0 High risk":"OK",color:AMBER},
  ];

  const prevLit=idx>0?months[idx-1]:null, prevGap=prevLit?zero.includes(prevLit):false;
  const topPrevNote=prevGap?`* Prev mo. = previous complete billed month; ${prevLit} was skipped (missing batch).`:"";
  const maxFocal=Math.max(...rev.top_clients.map(c=>c.focal));
  const topRows=rev.top_clients.map((c,i)=>{
    const pm=c.prev_month||0; let mom,momColor,prevDisp,prevColor;
    if(pm>0){const mv=(c.focal-pm)/pm;mom=dPct(mv);momColor=mv>=0?GREEN:RED;prevDisp=usd(pm);prevColor=null;}
    else{mom="\u2014";momColor=GREY;prevDisp="\u2014";prevColor=GREY;}
    return {new:!!c.new_logo,label:`${i+1}  ${c.name}${c.new_logo?'  \u2605 new':''}`,
      this:c.focal,thisDisp:usd(c.focal),frac:c.focal/maxFocal,
      prev:prevDisp,prevColor,mom,momColor,t3m:usd(c.t3m),
      yoy:usd(c.yoy_same_month||0),yoyColor:(c.yoy_same_month||0)===0?GREY:null};
  });

  const dc=rev.decomposition;
  const decompRows=[
    {label:"New",color:GREEN,amt:dUsd(dc.new.amount),n:dc.new.clients},
    {label:"Expansion",color:GREEN,amt:dUsd(dc.expansion.amount),n:dc.expansion.clients},
    {label:"Contraction",color:RED,amt:dUsd(dc.contraction.amount),n:dc.contraction.clients},
    {label:"Churn (full logo loss)",color:dc.churn.clients===0?GREEN:RED,amt:dUsd(dc.churn.amount),n:dc.churn.clients},
  ];
  const netChange=dc.new.amount+dc.expansion.amount+dc.contraction.amount+dc.churn.amount;

  const st=(t,c,b=false)=>({t,color:c,b});
  const metrics=[
    ["MRR",`${usdk(fv)} prov. / ${usdk(mrrSm)} smoothed`,"ARR/12",st("Tracking",GREY)],
    ["ARR",`~${usdk(arrCur)} floor / ~${usdk(arrSm)} smoothed`,"10× baseline",st("Tracking",GREY)],
    ["Net new ARR (mo)","provisional (decomposition)","(10×−1)/12",hasGap?st("Blocked — gap",RED):st("Tracking",GREY)],
    ["NRR (T3M proxy)",(narr.nrr_text||"see decomposition")+(hasGap?"  \u26A0 distorted":""),"\u2265110%",hasGap?st("Re-run post-backfill",AMBER):st("Tracking",GREY)],
    ["GRR (T3M proxy)",(narr.grr_text||"see decomposition")+(hasGap?"  \u26A0 distorted":""),"\u226595%",hasGap?st("Re-run post-backfill",AMBER):st("Tracking",GREY)],
    ["CAC by channel","not attributable yet","<18mo payback",st("Needs channel tags",AMBER)],
    ["Magic Number","n/a (needs prior-Q S&M)","\u22651.0",st("Needs S&M base",AMBER)],
    ["Rule of 40","n/a (needs FCF margin)","\u226540",st("Needs full P&L",AMBER)],
    ["Tool spend / revenue",pct(toolPct),"<5%",toolPct<0.05?st("\u2713 Pass",GREEN,true):st("\u26A0 Over",AMBER,true)],
    ["Concentration top-10",pct(conc.top10_share),"(watch)",conc.top10_share>0.70?st("\u26A0 High",AMBER,true):st("OK",GREEN,true)],
  ];

  const maxSeries=Math.max(...series.map(m=>m.revenue));
  const trend=series.map((m,i)=>{
    const prev=i>0?series[i-1].revenue:null, mom=(prev&&prev>0)?(m.revenue-prev)/prev:null;
    let color=NAVY; if(m.revenue===0)color=RED; else if(m.month===focal)color=ORANGE; else if(m.revenue===maxSeries)color=GREEN;
    return {label:shortM(m.month),frac:m.revenue/maxSeries,rev:m.revenue===0?"—":usdk(m.revenue),color,
      mom:mom==null?"—":dPct(mom),momColor:mom==null?GREY:(mom>=0?GREEN:RED)};
  });

  return {meta,fx,kpis,flags,bottom_line:narr.bottom_line||[],
    A:{headline:usd(fv),sub:`${fcount} billing events  •  ${rev.unique_clients??'?'} unique paying clients`,
       planRows,planNote:narr.plan_note||"Recommendation: add a plan_type deal property so the split is exact, not heuristic.",
       opexRows,netEq:`${usd(fv)} − ${usd(usdTot)} = `,netVal:usd(net),netSub:`${pct(contrib)} contribution after tracked tooling/marketing`,
       netCaveat:"This is not a gross or operating margin — CDN delivery COGS and payroll are absent from the data room.",trend},
    B:{yoyPre:yoy?`${usd(fv)} vs ${usd(yoy.prior)}   →   `:"",yoyVal:yoy?`${sUsd(yoy.abs)}  /  ${dPct(yoy.pct)}`:"n/a",
       t3mRows:[[`T3M (${win.map(m=>m.month.slice(-2)+'='+usdk(m.revenue)).join(' + ')})`,usd(t3m),""],
                ["Prior T3M",usd(pt3m),{t:dPct(t3mCh),color:t3mCh<0?RED:GREEN}],
                ...(gapInWin?[{total:true,cells:["With backfill (≈ trailing avg)","≈"+usdk(t3mBf),{t:"≈"+sPct(t3mBfCh),color:t3mBfCh>=0?GREEN:RED}]}]:[])],
       t3mNote:narr.t3m_note||(gapInWin?"The headline change is an artifact of the missing batch — do not act on it until the gap is loaded.":""),
       decompRows,decompNet:dUsd(netChange),decompMaterial:dc.material_clients??rev.top_clients.length,decompWindow:dc.window||"",
       decompNote:narr.decomp_note||"Zero churn is the real signal — no clients fully lost. Re-run after backfill for a true decomposition."},
    C:{topRows,topPrevNote,topNote:narr.top_note||"",
       concPre:`Top-10 share ${pct(conc.top10_share)}`,concPost:`  •  Top-1 share ${pct(conc.top1_share)} (${conc.top1_name||''})`,
       conc:{top1:conc.top1_share,top2_10:conc.top10_share-conc.top1_share,rest:1-conc.top10_share,headline:pct(conc.top10_share)},
       concReco:narr.concentration_note||"Loss of any single top account is material. Recommend a named-account QBR/retention program for the top 10 and a diversification target (top-10 < 70% by year-end).",
       hygiene:narr.hygiene_note||""},
    D:{arrRows:[["Current (provisional)",usd0(fv),"≈"+usdk(arrCur)],{total:true,cells:["Smoothed (complete months, excl. gap)",usd0(mrrSm),"≈"+usdk(arrSm)]}],
       arrReco:narr.arr_recommendation||`Plan against ARR ≈ ${usdk(arrSm)} (smoothed working run-rate); ${usdk(arrCur)} is the conservative floor until the gap is backfilled and the month reconciled.`,
       pipelineNote:narr.pipeline_note||"Pipeline-to-ARR conversion: not computable — \"Current Clients Pipeline\" is a single-stage customer ledger, not a funnel. Action: stand up a sales pipeline with open stages (Lead → Qualified → Proposal → Won)."},
    E:{cacNote:narr.cac_note||"CAC by channel: only paid ad spend is attributable. Not a real CAC until new-logo counts are channel-tagged and fully-loaded S&M cost is captured.",
       toolNote:narr.toolspend_note||`Tool spend vs revenue: recurring tooling ≈ ${usd0(toolUsd)}/mo = ${pct(toolPct)} of MRR — under the <5% target.`},
    F:{title:"Cash runway — not computable this cycle",lines:narr.runway_lines||[
       "Missing inputs: opening cash balance (held by YODA / banking) and fully-loaded monthly burn (payroll + CDN infrastructure COGS).",
       `On tracked OpEx alone the business is cash-generative (${sUsd(net)}), but that ignores the two largest real cost buckets (people, delivery infra) — so it is NOT a runway statement.`]},
    metrics};
}

/* ============================ docx-js helpers ============================ */
const HALR=(s,o={})=>new TextRun({text:String(s),bold:!!o.b,italics:!!o.i,
  color:o.color||undefined,size:o.sz||20,font:"Arial"});

function P(children,o={}){
  const opts={children:Array.isArray(children)?children:[children],
    alignment:o.jc==='right'?AlignmentType.RIGHT:(o.jc==='center'?AlignmentType.CENTER:AlignmentType.LEFT)};
  const sp={}; if(o.before!=null)sp.before=o.before; if(o.after!=null)sp.after=o.after;
  if(Object.keys(sp).length)opts.spacing=sp;
  if(o.bdrBottom)opts.border={bottom:{style:BorderStyle.SINGLE,size:6,space:2,color:o.bdrBottom}};
  return new Paragraph(opts);
}
const TXT=(text,o={})=>P([HALR(text,o)],o);
const SPACER=(after=120)=>new Paragraph({children:[],spacing:{after}});

const thin=c=>({style:BorderStyle.SINGLE,size:4,color:c});
const nilB={style:BorderStyle.NONE,size:0,color:"FFFFFF"};
function cellBorders(c="CCCCCC",leftBar){
  return {top:thin(c),bottom:thin(c),right:thin(c),
    left:leftBar?{style:BorderStyle.SINGLE,size:18,color:leftBar}:thin(c)};
}
const noBorders={top:nilB,bottom:nilB,left:nilB,right:nilB};

function CELL(children,o={}){
  const m=o.mar||{t:60,b:60,l:120,r:120};
  return new TableCell({
    width:{size:o.w||0,type:WidthType.DXA},
    borders:o.noBorder?noBorders:cellBorders(o.borderColor||"CCCCCC",o.leftBar),
    shading:o.fill?{fill:o.fill,type:ShadingType.CLEAR,color:"auto"}:undefined,
    margins:{top:m.t,bottom:m.b,left:m.l,right:m.r},
    verticalAlign:o.valign==='top'?VerticalAlign.TOP:VerticalAlign.CENTER,
    children:Array.isArray(children)?children:[children]});
}
function ROW(cells,o={}){return new TableRow({children:cells,cantSplit:true,tableHeader:!!o.header});}
function TBL(cols,rows,o={}){
  return new Table({
    width:{size:o.w||CW,type:WidthType.DXA},
    columnWidths:cols,
    borders:o.noBorder?{top:nilB,bottom:nilB,left:nilB,right:nilB,
      insideHorizontal:nilB,insideVertical:nilB}:undefined,
    rows});
}

/* numbered + bullet list paragraphs (proper numbering config, per skill) */
const NUM=(text,ref)=>new Paragraph({numbering:{reference:ref,level:0},
  children:[HALR(text,{sz:18})]});

/* data table: rows are arrays of (string | {t,color,b}) or {total:true,cells:[...]} */
function dataTable(cols,headers,rows){
  const head=ROW(headers.map((h,i)=>CELL(TXT(h,{b:true,color:WHITE,sz:18}),
    {w:cols[i],fill:NAVY,borderColor:NAVY})),{header:true});
  const body=rows.map((r,ri)=>{
    const isTotal=r&&r.total; const cells=isTotal?r.cells:r;
    return ROW(cells.map((c,ci)=>{
      const obj=c&&typeof c==="object"; const text=obj?c.t:c;
      const fill=isTotal?LIGHT:(ri%2?WHITE:LIGHT);
      const jc=ci>0?'right':undefined;
      return CELL(TXT(text,{sz:18,b:isTotal||(obj&&c.b),color:obj?c.color:undefined,jc}),{w:cols[ci],fill});
    }));
  });
  return TBL(cols,[head,...body]);
}
function callout(title,bodyParas,{fill=AMBERBG,bar=AMBER}={}){
  const inner=[P([HALR(title,{b:true,color:bar,sz:20})],{after:60}),...bodyParas];
  return TBL([CW],[ROW([CELL(inner,{w:CW,fill,leftBar:bar,borderColor:"E0D6BE",mar:{t:120,b:120,l:200,r:160}})])]);
}
function kpiCard(k,w){
  return CELL([
    TXT(k.value,{b:true,color:NAVY,sz:30,jc:'center',after:20}),
    TXT(k.label,{color:GREY,sz:15,jc:'center',after:40}),
    TXT(k.status,{b:true,color:k.color,sz:15,jc:'center'})],
    {w,fill:LIGHT,borderColor:"D9DEE4",mar:{t:120,b:120,l:80,r:80}});
}
const h2=t=>P([HALR(t,{b:true,color:NAVY,sz:26})],{before:240,after:120,bdrBottom:NAVY});

/* block bar runs (filled + track), returns TextRun[] */
function blockBar(frac,color,max=14){
  frac=Math.max(0,Math.min(1,frac)); const n=Math.max(1,Math.round(frac*max));
  return [HALR("\u2588".repeat(n),{color,sz:14}),HALR("\u2591".repeat(max-n),{color:TRACK,sz:14})];
}

/* ============================ render ============================ */
function render(M){
  const body=[];

  // title band (borderless 1-cell navy table)
  body.push(TBL([CW],[ROW([CELL([
    TXT(M.meta.company.toUpperCase(),{b:true,color:ORANGE,sz:22,after:30}),
    TXT("Monthly Financial Report",{b:true,color:WHITE,sz:40,after:20}),
    TXT("Reporting period: "+M.meta.period_label,{color:"C9D4E3",sz:22})],
    {w:CW,noBorder:true,fill:NAVY,mar:{t:200,b:160,l:240,r:240}})])],{noBorder:true}));
  body.push(TXT(`Prepared by ${M.meta.prepared_by}   •   Date ${M.meta.date}   •   For: ${M.meta.for}`,{sz:16,color:GREY,before:120,after:20}));
  body.push(TXT("Source of truth: "+M.meta.source_portal+".",{sz:15,color:GREY,i:true,after:160}));

  // KPI strip
  const cw4=[2340,2340,2340,2340];
  body.push(TBL(cw4,[ROW(M.kpis.map((k,i)=>kpiCard(k,cw4[i])))],{noBorder:true}));
  body.push(SPACER(120));

  // integrity
  body.push(callout("\u26A0  Data integrity — read this first",
    M.flags.map(f=>NUM(f,"ig")),{fill:AMBERBG,bar:AMBER}));
  body.push(SPACER(120));

  // bottom line
  body.push(h2("CFO bottom line (for Hermes → CEO)"));
  M.bottom_line.forEach(t=>body.push(NUM(t,"bl")));
  body.push(SPACER(80));

  // A
  body.push(h2("A.  Monthly revenue & result — "+M.meta.period_label));
  body.push(P([HALR("Revenue: ",{b:true,sz:20}),HALR(M.A.headline,{b:true,color:NAVY,sz:24}),HALR("    "+M.A.sub,{color:GREY,sz:16})],{after:120}));
  body.push(TXT("Revenue by plan type:",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Plan type","Revenue","Share"],M.A.planRows));
  body.push(TXT(M.A.planNote,{i:true,sz:15,color:GREY,before:60,after:140}));
  body.push(TXT("Tracked OpEx (tooling, ads, AI only; not a full P&L):",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Category","EUR","USD"],M.A.opexRows));
  body.push(callout("Net result — net of tracked OpEx only (NOT operating profit)",[
    P([HALR(M.A.netEq,{sz:18}),HALR(M.A.netVal,{b:true,color:GREEN,sz:22}),HALR("   •   "+M.A.netSub,{sz:16,color:GREY})]),
    TXT(M.A.netCaveat,{i:true,sz:15,color:GREY,before:40})],{fill:GREENBG,bar:GREEN}));

  // trend bar table
  body.push(TXT("Revenue trend — monthly",{b:true,sz:18,before:120,after:40}));
  const tcols=[1100,5200,1530,1530];
  const tHead=ROW(["Month","","Revenue","MoM"].map((h,i)=>CELL(TXT(h,{b:true,color:WHITE,sz:16}),{w:tcols[i],fill:NAVY,borderColor:NAVY})),{header:true});
  const tBody=M.A.trend.map(t=>ROW([
    CELL(TXT(t.label,{sz:16}),{w:tcols[0]}),
    CELL(P(blockBar(t.frac,t.color)),{w:tcols[1]}),
    CELL(TXT(t.rev,{sz:16,jc:'right'}),{w:tcols[2]}),
    CELL(TXT(t.mom,{sz:15,jc:'right',color:t.momColor}),{w:tcols[3]})]));
  body.push(TBL(tcols,[tHead,...tBody]));

  // B
  body.push(h2("B.  Trends & comparisons"));
  body.push(callout("YoY (month) vs same month last year  •  CLEAN, trust this",
    [P([HALR(M.B.yoyPre,{sz:18}),HALR(M.B.yoyVal,{b:true,color:GREEN,sz:20})])],{fill:GREENBG,bar:GREEN}));
  body.push(SPACER(80));
  body.push(TXT("3-month block (T3M) vs prior T3M:",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Window","Value","Change"],M.B.t3mRows));
  if(M.B.t3mNote)body.push(TXT(M.B.t3mNote,{i:true,sz:15,color:GREY,before:60,after:140}));
  body.push(TXT(`Growth decomposition (material clients, ${M.B.decompWindow}):`,{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Movement","Amount","Clients"],
    M.B.decompRows.map(r=>[{t:r.label,color:r.color},{t:r.amt,color:r.color},String(r.n)])
      .concat([{total:true,cells:["Net change",M.B.decompNet,String(M.B.decompMaterial)]}])));
  if(M.B.decompNote)body.push(TXT(M.B.decompNote,{i:true,sz:15,color:GREY,before:60,after:120}));

  // C
  body.push(h2("C.  Customers — Top "+M.C.topRows.length));
  const cCols=[2520,1480,1300,1080,1480,1500];
  const cHead=ROW(["Client","This mo","Prev mo*","MoM","T3M","Same mo LY"].map((h,i)=>CELL(TXT(h,{b:true,color:WHITE,sz:18}),{w:cCols[i],fill:NAVY,borderColor:NAVY})),{header:true});
  const cBody=M.C.topRows.map((r,ri)=>{const fill=ri%2?WHITE:LIGHT;
    return ROW([
      CELL(TXT(r.label,{sz:17,color:r.new?ORANGE:undefined}),{w:cCols[0],fill}),
      CELL(P([HALR(r.thisDisp+" ",{sz:16}),...blockBar(r.frac,r.new?ORANGE:NAVY,6)],{jc:'right'}),{w:cCols[1],fill}),
      CELL(TXT(r.prev,{sz:16,jc:'right',color:r.prevColor}),{w:cCols[2],fill}),
      CELL(TXT(r.mom,{sz:16,jc:'right',color:r.momColor}),{w:cCols[3],fill}),
      CELL(TXT(r.t3m,{sz:16,jc:'right'}),{w:cCols[4],fill}),
      CELL(TXT(r.yoy,{sz:16,jc:'right',color:r.yoyColor}),{w:cCols[5],fill})]);});
  body.push(TBL(cCols,[cHead,...cBody]));
  if(M.C.topPrevNote)body.push(TXT(M.C.topPrevNote,{i:true,sz:14,color:GREY,before:50}));
  if(M.C.topNote)body.push(TXT(M.C.topNote,{sz:16,color:GREEN,before:60,after:80}));

  // concentration stacked bar
  body.push(TXT("Revenue concentration (top-10)",{b:true,sz:17,before:120,after:40}));
  const c=M.C.conc, w1=Math.round(CW*c.top1),w2=Math.round(CW*c.top2_10),w3=CW-w1-w2;
  body.push(TBL([w1,w2,w3],[ROW([
    CELL(TXT(`Top-1  ${pct(c.top1)}`,{b:true,color:WHITE,sz:14,jc:'center'}),{w:w1,fill:ORANGE,borderColor:WHITE}),
    CELL(TXT(`Top 2–10  ${pct(c.top2_10)}`,{b:true,color:WHITE,sz:14,jc:'center'}),{w:w2,fill:NAVY,borderColor:WHITE}),
    CELL(TXT(`Rest  ${pct(c.rest)}`,{color:GREY,sz:14,jc:'center'}),{w:w3,fill:LIGHT,borderColor:WHITE})])],{noBorder:true}));
  body.push(callout("\u26A0  Concentration risk",[
    P([HALR(M.C.concPre,{b:true,sz:18}),HALR(M.C.concPost,{sz:16})]),
    TXT(M.C.concReco,{sz:16,before:40})],{fill:AMBERBG,bar:AMBER}));
  if(M.C.hygiene)body.push(TXT(M.C.hygiene,{i:true,sz:14,color:GREY,before:60,after:80}));

  // D
  body.push(h2("D.  Run-rate & forward"));
  body.push(TXT("ARR snapshot (recurring book, so MRR ≈ monthly revenue):",{b:true,sz:17,after:60}));
  body.push(dataTable([5560,1900,1900],["Basis","MRR","ARR"],M.D.arrRows));
  body.push(TXT(M.D.arrReco,{sz:16,before:60,after:60}));
  body.push(NUM(M.D.pipelineNote,"bul"));
  body.push(SPACER(60));

  // E
  body.push(h2("E.  Unit economics & efficiency"));
  body.push(NUM(M.E.cacNote,"bul"));
  body.push(NUM(M.E.toolNote,"bul"));
  body.push(SPACER(60));

  // F
  body.push(h2("F.  Cash & runway"));
  body.push(callout(M.F.title,M.F.lines.map((l,i)=>TXT(l,{sz:16,before:i?40:undefined})),{fill:GREYBG,bar:GREY}));
  body.push(SPACER(80));

  // metrics
  body.push(h2("SaaS metrics summary"));
  body.push(dataTable([2860,3100,1700,1700],["Metric","Value","Target","Status"],
    M.metrics.map(r=>[r[0],r[1],r[2],r[3]])));
  body.push(TXT("— End of report —",{jc:'center',color:GREY,sz:15,before:240}));

  // footer with Page X of Y
  const footer=new Footer({children:[new Paragraph({
    alignment:AlignmentType.CENTER,
    border:{top:{style:BorderStyle.SINGLE,size:4,space:6,color:"D9DEE4"}},
    children:[
      HALR(`${M.meta.company} — Monthly Financial Report — ${M.meta.period_label}   •   Confidential   •   Page `,{color:GREY,sz:14}),
      new TextRun({children:[PageNumber.CURRENT],color:GREY,size:14,font:"Arial"}),
      HALR(" of ",{color:GREY,sz:14}),
      new TextRun({children:[PageNumber.TOTAL_PAGES],color:GREY,size:14,font:"Arial"})]})]});

  return new Document({
    styles:{default:{document:{run:{font:"Arial",size:20}}}},
    numbering:{config:[
      {reference:"bl",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,
        style:{run:{size:18,font:"Arial"},paragraph:{indent:{left:360,hanging:300}}}}]},
      {reference:"ig",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,
        style:{run:{size:18,font:"Arial"},paragraph:{indent:{left:360,hanging:300}}}}]},
      {reference:"bul",levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",alignment:AlignmentType.LEFT,
        style:{run:{size:18,font:"Arial"},paragraph:{indent:{left:360,hanging:300}}}}]},
    ]},
    sections:[{
      properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1440,bottom:1080,left:1440}}},
      footers:{default:footer},
      children:body}]});
}

/* ============================ main ============================ */
const inp=process.argv[2]||'input.json', out=process.argv[3]||'Monthly_Financial_Report.docx';
const data=JSON.parse(fs.readFileSync(inp,'utf8'));
Packer.toBuffer(render(compute(data))).then(buf=>{fs.writeFileSync(out,buf);console.log('✓ wrote',out);});
```

## Example input (the exact contract — save your real data as input.json in this shape)

```json
{
  "meta": {
    "company": "BlazingCDN",
    "report_type": "monthly",
    "period_label": "May 2026",
    "period": "2026-05",
    "prepared_by": "Mike Scarpelli, CFO",
    "date": "2026-06-04",
    "for": "Hermes (Chief of Staff) \u2192 CEO",
    "source_portal": "HubSpot production portal (Hub 143144902), Deals \u2192 \"Current Clients Pipeline\". Expenses: Google Drive \"CFO\"",
    "fx_eur_usd": 1.08,
    "currency": "USD",
    "pre_batch": true,
    "close_date": "2026-06-05"
  },
  "revenue": {
    "focal_month": "2026-05",
    "unique_clients": 51,
    "monthly_series": [
      {"month":"2024-12","count":42,"revenue":12019.25},
      {"month":"2025-01","count":49,"revenue":19263.70},
      {"month":"2025-02","count":58,"revenue":15112.45},
      {"month":"2025-03","count":52,"revenue":20818.34},
      {"month":"2025-04","count":66,"revenue":22933.86},
      {"month":"2025-05","count":65,"revenue":23053.53},
      {"month":"2025-06","count":63,"revenue":25663.82},
      {"month":"2025-07","count":71,"revenue":19082.55},
      {"month":"2025-08","count":70,"revenue":27007.60},
      {"month":"2025-09","count":73,"revenue":30955.26},
      {"month":"2025-10","count":80,"revenue":34083.66},
      {"month":"2025-11","count":81,"revenue":31364.80},
      {"month":"2025-12","count":58,"revenue":29029.88},
      {"month":"2026-01","count":84,"revenue":32340.04},
      {"month":"2026-02","count":75,"revenue":26590.56},
      {"month":"2026-03","count":97,"revenue":43844.38},
      {"month":"2026-04","count":0,"revenue":0.00},
      {"month":"2026-05","count":68,"revenue":27787.01}
    ],
    "plan_split": {
      "recurring": {"revenue":27672.19,"clients":47},
      "payg": {"revenue":114.82,"clients":4}
    },
    "top_clients": [
      {"name":"rabin.yor\u2026","focal":4875.98,"prev_month":4976.69,"t3m":9852.67,"yoy_same_month":4743.60,"new_logo":false},
      {"name":"eliaskhan.it","focal":4675.00,"prev_month":8000.00,"t3m":12675.00,"yoy_same_month":5500.00,"new_logo":false},
      {"name":"itgroup (avanquest)","focal":4000.00,"prev_month":200.00,"t3m":4200.00,"yoy_same_month":0.00,"new_logo":true},
      {"name":"yashwanth (moengage)","focal":2250.00,"prev_month":2250.00,"t3m":4500.00,"yoy_same_month":0.00,"new_logo":true},
      {"name":"a.ruin (epom)","focal":2000.00,"prev_month":2410.55,"t3m":4410.55,"yoy_same_month":0.00,"new_logo":true},
      {"name":"secretfy.corp","focal":2000.00,"prev_month":2000.00,"t3m":4000.00,"yoy_same_month":0.00,"new_logo":true},
      {"name":"arik.holaspark","focal":2000.00,"prev_month":1500.00,"t3m":3500.00,"yoy_same_month":3700.00,"new_logo":false},
      {"name":"theqooking","focal":1114.29,"prev_month":1033.73,"t3m":2148.02,"yoy_same_month":1261.90,"new_logo":false},
      {"name":"manuel.skelton","focal":603.00,"prev_month":900.00,"t3m":1503.00,"yoy_same_month":0.00,"new_logo":true},
      {"name":"bret.cityspark","focal":592.56,"prev_month":561.05,"t3m":1153.61,"yoy_same_month":620.20,"new_logo":false},
      {"name":"devops.nordix","focal":430.00,"prev_month":470.00,"t3m":900.00,"yoy_same_month":300.00,"new_logo":false},
      {"name":"media.streamhub","focal":398.00,"prev_month":360.00,"t3m":758.00,"yoy_same_month":0.00,"new_logo":true},
      {"name":"cdn.playwave","focal":360.00,"prev_month":470.00,"t3m":830.00,"yoy_same_month":410.00,"new_logo":false},
      {"name":"ott.brightcast","focal":322.00,"prev_month":305.00,"t3m":627.00,"yoy_same_month":0.00,"new_logo":true},
      {"name":"game.patchgrid","focal":295.00,"prev_month":270.00,"t3m":565.00,"yoy_same_month":240.00,"new_logo":false},
      {"name":"saas.fluxdeliver","focal":268.00,"prev_month":268.00,"t3m":536.00,"yoy_same_month":260.00,"new_logo":false},
      {"name":"vod.castline","focal":242.00,"prev_month":210.00,"t3m":452.00,"yoy_same_month":0.00,"new_logo":true},
      {"name":"web.edgepush","focal":210.00,"prev_month":250.00,"t3m":460.00,"yoy_same_month":220.00,"new_logo":false},
      {"name":"img.snapcdn","focal":182.00,"prev_month":160.00,"t3m":342.00,"yoy_same_month":140.00,"new_logo":false},
      {"name":"api.routemesh","focal":158.00,"prev_month":150.00,"t3m":308.00,"yoy_same_month":0.00,"new_logo":true}
    ],
    "concentration": {"top10_share":0.868,"top1_share":0.175,"top1_name":"rabin.yor\u2026"},
    "decomposition": {
      "window":"Mar\u2013May 2026","material_threshold":500,"material_clients":14,
      "new":{"amount":4200.00,"clients":1},
      "expansion":{"amount":14418.43,"clients":4},
      "contraction":{"amount":-32556.01,"clients":9},
      "churn":{"amount":0.00,"clients":0}
    }
  },
  "expenses": {
    "month": "2026-05",
    "full_pl": false,
    "recurring_tool_eur": 527,
    "lines": [
      {"category":"Marketing","detail":"Google Ads \u20AC1,069; LinkedIn \u20AC0","eur":1069.00},
      {"category":"Product & AI","detail":"OpenAI \u20AC120; Claude \u20AC120; Jina \u20AC50","eur":290.00},
      {"category":"Ops & infra","detail":"HubSpot \u20AC98; make.com \u20AC17; Atlassian \u20AC20; GCP \u20AC40","eur":175.00},
      {"category":"Sales tooling","detail":"SalesQL \u20AC68; Zapmail \u20AC34; QEV \u20AC30","eur":132.00}
    ]
  },
  "cfo_narrative": {
    "bottom_line": [
      "May 2026 revenue $27.8K, +20.5% YoY \u2014 real, clean growth, driven by 5 net-new top accounts landed over the past year. Zero logo churn.",
      "Headline trend metrics (T3M, NRR, GRR) are NOT real \u2014 artifacts of the missing April batch. #1 action is backfilling April; until then ignore them.",
      "True margins and runway are not yet producible \u2014 the data room lacks CDN COGS, payroll and cash balance. Closing that gap unlocks Rule-of-40 / Magic Number / runway.",
      "Risk to watch: 87% revenue concentration in the top 10 accounts.",
      "Working run-rate to plan against: ARR \u2248 $390K (smoothed)."
    ],
    "plan_note": "Recurrence heuristic \u2014 no contractual plan_type field yet. Recommendation: add a plan_type deal property so the split is exact.",
    "decomp_note": "The large contraction is almost entirely the April hole (every recurring payer is missing one month in the window). Zero churn is the real signal \u2014 no clients fully lost. Re-run after backfill.",
    "top_note": "Good-news signal: 5 of the top 10 had $0 in May 2025 \u2014 these net-new accounts are the real engine behind +20.5% YoY.",
    "concentration_note": "With 87% of revenue in 10 accounts, loss of any single top account is material. Recommend a named-account QBR/retention program for the top 10 and a diversification target (top-10 < 70% by year-end).",
    "hygiene_note": "Data hygiene: one $1.00 test deal and two small negatives (\u2212$9.24, \u2212$12.27, likely refunds) in the log \u2014 immaterial; clean up the test record.",
    "nrr_text": "75.1%",
    "grr_text": "57.2%"
  }
}
```
