---
name: cfo-quarterly-report
description: "Generate the polished BlazingCDN QUARTERLY financial report as a Google-Docs-clean Word (.docx) via docx-js: same house style as the monthly report but quarter-framed, with QoQ + YoY-by-quarter comparisons, a quarterly revenue trend bar-table, a Top-10 customers table (This Q / Prior Q / QoQ% / Same-Q-LY), a Top-10 concentration bar, and the quarter-only metrics promoted to first-class: NRR, GRR, Magic Number, Rule of 40. Single self-contained markdown file: embeds the docx-js render engine plus the input contract. Use at quarter close once the data workers have returned quarterly aggregates. Do NOT use for monthly reports (use cfo-monthly-report)."
---

# CFO Quarterly Financial Report  (single-file .md skill)

This skill builds BlazingCDN's quarterly CFO report as a polished `.docx` using **docx-js** (the `docx` skill's library). The engine does both the math and the rendering. Charts are drawn as shaded / block-bar (`█`/`░`) table cells — no images. House style is **identical** to `cfo-monthly-report`; separate skill so quarterly logic never loads on a monthly run. Deterministic: same engine → identical design; only numbers change.

The output follows the `docx` skill's Google-Docs-safe rules (DXA widths, `ShadingType.CLEAR`, dual-width tables, numbering-config lists, no table-rule dividers, `PageNumber` footer fields), so it opens/converts to a native Google Doc without the artefacts of hand-rolled OOXML.

> **Dependency note (changed from the old engine):** this is no longer a zero-dependency stdlib engine. It depends on the `docx` npm package and — optionally — the `docx` skill's `validate.py`. That is the deliberate trade for clean Google Docs output.

## How to produce a report (runtime steps)
1. **Materialize the engine.** Copy the entire `build.js` code block (under "Engine" below) verbatim into a file named `build.js` in your workspace.
2. **Install docx once.** In that folder run `npm install docx` (Node v18+). No other packages are needed.
3. **Assemble the input.** Build `input.json` from the two data workers' quarterly JSON + your `cfo_narrative`, using the "Example input" block below as the exact contract. Include `decomposition.starting_recurring` so NRR/GRR compute, and optionally `inputs.prior_q_sm_usd` (prior-quarter S&M) so the Magic Number computes.
4. **Run:** `node build.js input.json "BlazingCDN_Quarterly_Financial_Report_<Q_Year>.docx"`.
5. **Validate (optional but recommended):** `python <docx-skill>/scripts/office/validate.py "<output>.docx"`.
6. Review, save to Google Drive "CFO", post the link to Hermes. The `.docx` is Google-Docs-clean, so it can also be uploaded to Drive **with conversion** (`mimeType: application/vnd.google-apps.document`) to land a native Google Doc.

## What differs from monthly
- Trend bar-table and tables are by **quarter**.
- Section B is **QoQ** (this quarter vs prior quarter) plus same-quarter-last-year.
- Section C is the **Top 10** with This Q / Prior Q / QoQ% / Same Q LY.
- Metrics compute **NRR / GRR** (from `starting_recurring`), **Magic Number** (net-new ARR ÷ prior-Q S&M), and frame **Rule of 40** (FCF margin stays an OpEx-only proxy until a full P&L exists). ARR = quarter-exit MRR × 12.

## Design behaviour (automatic)
- **Signed deltas carry arrows:** positive → green ▲, negative → red ▼, exact zero → neutral (no arrow). Applied to KPI QoQ/YoY, per-client QoQ, trend QoQ, the comparison-table Change column, growth-decomposition amounts, and net-new ARR. The Rule-of-40 growth value and Section F narrative prose are left arrow-free on purpose.
- **Footer shows `Page X of Y`** via docx-js `PageNumber.CURRENT` / `PageNumber.TOTAL_PAGES` fields.
- The full visual spec lives in the engine — see the `DESIGN INVARIANTS` comment and the palette constants at the top of `build.js`. House style is identical to `cfo-monthly-report`; there is no separate spec doc to drift.

## NOTE — production Hub ID
Canonical production Hub ID is **143144902**, already baked into `meta.source_portal` in the example below (the prior `143144902` vs `145006611` ambiguity is resolved — `145006611` was the non-production portal). Confirm it still matches your live portal before each run; change the single value in `meta.source_portal` if the portal ever moves.

## Engine — build.js  (copy verbatim into a file named build.js)

```javascript
#!/usr/bin/env node
/*
 * cfo-quarterly-report / build.js  — renders via docx-js (the `docx` skill).
 * Replaces the hand-rolled OOXML engine. Output is a clean, validated .docx that
 * imports into Google Docs without the artefacts of raw field codes / fixed OOXML.
 *
 *   npm install docx          (once, in this folder)
 *   node build.js <input.json> <output.docx>
 *
 * Same quarterly input.json CONTRACT as before — only rendering changed. All math
 * (ARR, QoQ, YoY-by-quarter, decomposition, NRR/GRR, Magic Number, Rule-of-40 frame,
 * concentration, integrity flags, metric statuses) is still done here in compute();
 * you write only judgement.
 *
 * DESIGN INVARIANTS (unchanged, identical house style to cfo-monthly-report):
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

/* ============================ compute (unchanged quarterly logic) ============================ */
const usd=x=>`$${x.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const usd0=x=>`$${Math.round(x).toLocaleString('en-US')}`;
const usdk=x=>`$${(x/1000).toFixed(1)}K`;
const pct=x=>`${(x*100).toFixed(1)}%`;
const sUsd=x=>(x>=0?"+":"\u2212")+`$${Math.abs(x).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const sPct=x=>(x>=0?"+":"\u2212")+`${Math.abs(x*100).toFixed(1)}%`;
const arrow=x=>x>0?"\u25B2 ":(x<0?"\u25BC ":"");
const dUsd=x=>arrow(x)+sUsd(x);
const dPct=x=>arrow(x)+sPct(x);

function compute(d){
  const meta=d.meta,rev=d.revenue,exp=d.expenses,narr=d.cfo_narrative||{},fx=meta.fx_eur_usd??1.08;
  const series=rev.quarterly_series,byQ=Object.fromEntries(series.map(q=>[q.quarter,q]));
  const quarters=series.map(q=>q.quarter),focal=rev.focal_quarter,idx=quarters.indexOf(focal);
  const fv=byQ[focal].revenue,fcount=byQ[focal].count;
  const zero=series.filter(q=>q.revenue===0).map(q=>q.quarter),hasGap=zero.length>0;

  const flags=[];
  for(const z of zero)flags.push(`${z} = $0 in source data — a missing batch, not a zero-revenue quarter; trailing metrics spanning it are PROVISIONAL. Action: backfill.`);
  if(!exp.full_pl)flags.push(`Expense files are not a full P&L — SaaS tooling, ads and AI only; no CDN delivery COGS, no payroll. Rule-of-40 FCF margin is therefore an OpEx-only proxy, not GAAP.`);
  if(meta.pre_batch)flags.push(`${meta.period_label} closes on the batch of ${meta.close_date||'the 5th'}; figures are provisional as of ${meta.date}. FX EUR→USD ${fx}.`);
  (narr.extra_integrity_flags||[]).forEach(f=>flags.push(f));

  let eurTot=0; const opexRows=[];
  for(const ln of exp.lines){eurTot+=ln.eur;opexRows.push([`${ln.category} (${ln.detail})`,`€${ln.eur.toLocaleString('en-US',{minimumFractionDigits:2})}`,usd(ln.eur*fx)]);}
  const usdTot=eurTot*fx; opexRows.push({total:true,cells:["Total tracked OpEx (quarter)",`€${eurTot.toLocaleString('en-US',{minimumFractionDigits:2})}`,usd(usdTot)]});
  const net=fv-usdTot, contrib=fv?net/fv:0;

  const rec=rev.plan_split.recurring,payg=rev.plan_split.payg,tot=rec.revenue+payg.revenue;
  const planRows=[[`Recurring / subscription-like (${rec.clients} clients)`,usd(rec.revenue),tot?pct(rec.revenue/tot):"0%"],
                  [`One-off / PAYG-like (${payg.clients} clients)`,usd(payg.revenue),tot?pct(payg.revenue/tot):"0%"]];

  const priorQ=idx>0?series[idx-1].revenue:null;
  const qoq=priorQ?{prior:priorQ,abs:fv-priorQ,pct:priorQ?(fv-priorQ)/priorQ:0}:null;
  let yoy=null; if(idx-4>=0){const py=series[idx-4].revenue;yoy={prior:py,abs:fv-py,pct:py?(fv-py)/py:0};}

  const exitMrr=fv/3, arrCur=exitMrr*12;
  const complete=series.filter(q=>q.revenue>0).map(q=>q.revenue);
  const arrSm=complete.length?(complete.slice(-2).reduce((a,b)=>a+b,0)/complete.slice(-2).length/3*12):arrCur;

  const recToolEur=exp.recurring_tool_eur??eurTot, toolUsd=recToolEur*fx, toolPct=fv?toolUsd/fv:0;
  const conc=rev.concentration;

  const kpis=[
    {value:usdk(fv),label:`Quarter revenue (${focal})`,status:`ARR ≈ ${usdk(arrCur)} run-rate`,color:NAVY},
    {value:qoq?dPct(qoq.pct):"n/a",label:"QoQ growth",status:"vs prior quarter",color:(qoq&&qoq.pct>=0)?GREEN:RED},
    {value:yoy?dPct(yoy.pct):"n/a",label:"YoY (quarter)",status:"vs same Q last year",color:GREEN},
    {value:pct(conc.top10_share),label:"Top-10 concentration",status:conc.top10_share>0.70?"\u26A0 High risk":"OK",color:AMBER},
  ];

  const maxFocal=Math.max(...rev.top_clients.map(c=>c.focal));
  const topRows=rev.top_clients.map((c,i)=>{
    const pq=c.prior_q||0; let mom,momColor,prevDisp,prevColor;
    if(pq>0){const mv=(c.focal-pq)/pq;mom=dPct(mv);momColor=mv>=0?GREEN:RED;prevDisp=usd(pq);prevColor=null;}
    else{mom="\u2014";momColor=GREY;prevDisp="\u2014";prevColor=GREY;}
    return {new:!!c.new_logo,label:`${i+1}  ${c.name}${c.new_logo?'  \u2605 new':''}`,
      this:c.focal,thisDisp:usd(c.focal),frac:c.focal/maxFocal,
      prev:prevDisp,prevColor,mom,momColor,
      yoy:usd(c.yoy_same_q||0),yoyColor:(c.yoy_same_q||0)===0?GREY:null};
  });

  const dc=rev.decomposition;
  const decompRows=[
    {label:"New",color:GREEN,amt:dUsd(dc.new.amount),n:dc.new.clients},
    {label:"Expansion",color:GREEN,amt:dUsd(dc.expansion.amount),n:dc.expansion.clients},
    {label:"Contraction",color:RED,amt:dUsd(dc.contraction.amount),n:dc.contraction.clients},
    {label:"Churn (full logo loss)",color:dc.churn.clients===0?GREEN:RED,amt:dUsd(dc.churn.amount),n:dc.churn.clients},
  ];
  const netChange=dc.new.amount+dc.expansion.amount+dc.contraction.amount+dc.churn.amount;

  const base=dc.starting_recurring; let nrr=null,grr=null;
  if(base){nrr=(base+dc.expansion.amount+dc.contraction.amount+dc.churn.amount)/base;
           grr=(base+dc.contraction.amount+dc.churn.amount)/base;}
  const smPrior=(d.inputs||{}).prior_q_sm_usd; const netNewArr=netChange*4;
  const magic=smPrior?netNewArr/smPrior:null;

  const st=(t,c,b=false)=>({t,color:c,b});
  const metrics=[
    ["ARR (quarter-exit)",`~${usdk(arrCur)}`,"10× baseline",st("Tracking",GREY)],
    ["Net new ARR (Q, annualized)",dUsd(netNewArr),"(10×−1)",st("Tracking",GREY)],
    ["NRR (quarter)",nrr!=null?pct(nrr):"needs starting base","\u2265110%",nrr==null?st("Needs base",AMBER):st(nrr>=1.10?"\u2713 Pass":"\u26A0 Below",nrr>=1.10?GREEN:AMBER,true)],
    ["GRR (quarter)",grr!=null?pct(grr):"needs starting base","\u226595%",grr==null?st("Needs base",AMBER):st(grr>=0.95?"\u2713 Pass":"\u26A0 Below",grr>=0.95?GREEN:AMBER,true)],
    ["Magic Number",magic!=null?magic.toFixed(2):"n/a (needs prior-Q S&M)","\u22651.0",magic==null?st("Needs S&M",AMBER):st(magic>=1?"\u2713 Pass":"\u26A0 Below",magic>=1?GREEN:AMBER,true)],
    ["Rule of 40",(yoy?sPct(yoy.pct):"?")+" growth + FCF% (proxy)","\u226540",st("Needs full P&L",AMBER)],
    ["CAC by channel","not attributable yet","<18mo payback",st("Needs channel tags",AMBER)],
    ["Tool spend / revenue",pct(toolPct),"<5%",toolPct<0.05?st("\u2713 Pass",GREEN,true):st("\u26A0 Over",AMBER,true)],
    ["Concentration top-10",pct(conc.top10_share),"(watch)",conc.top10_share>0.70?st("\u26A0 High",AMBER,true):st("OK",GREEN,true)],
  ];

  const maxSeries=Math.max(...series.map(q=>q.revenue));
  const trend=series.map((q,i)=>{
    const prev=i>0?series[i-1].revenue:null, mm=(prev&&prev>0)?(q.revenue-prev)/prev:null;
    let color=NAVY; if(q.revenue===0)color=RED; else if(q.quarter===focal)color=ORANGE; else if(q.revenue===maxSeries)color=GREEN;
    return {label:q.quarter,frac:q.revenue/maxSeries,rev:q.revenue===0?"—":usdk(q.revenue),color,
      mom:mm==null?"—":dPct(mm),momColor:mm==null?GREY:(mm>=0?GREEN:RED)};
  });

  return {meta,fx,kpis,flags,bottom_line:narr.bottom_line||[],
    A:{headline:usd(fv),sub:`${fcount} billing events  •  ${rev.unique_clients??'?'} unique paying clients (quarter)`,
       planRows,planNote:narr.plan_note||"",
       opexRows,netEq:`${usd(fv)} − ${usd(usdTot)} = `,netVal:usd(net),netSub:`${pct(contrib)} contribution after tracked tooling/marketing (quarter)`,
       netCaveat:"Not a gross/operating margin — CDN delivery COGS and payroll are absent from the data room.",trend},
    B:{yoyPre:qoq?`${usd(fv)} vs ${usd(qoq.prior)}   →   `:"",yoyVal:qoq?`${sUsd(qoq.abs)}  /  ${dPct(qoq.pct)} QoQ`:"n/a",
       t3mRows:[[`${focal} (this quarter)`,usd(fv),""],
                [`${quarters[idx-1]||'Prior Q'} (prior quarter)`,priorQ!=null?usd(priorQ):"n/a",qoq?{t:dPct(qoq.pct),color:qoq.pct<0?RED:GREEN}:""],
                ...(yoy?[{total:true,cells:[`${quarters[idx-4]} (same Q last year)`,usd(yoy.prior),{t:dPct(yoy.pct),color:yoy.pct>=0?GREEN:RED}]}]:[])],
       t3mNote:narr.qoq_note||"",
       decompRows,decompNet:dUsd(netChange),decompMaterial:dc.material_clients??rev.top_clients.length,decompWindow:dc.window||`${focal} vs prior quarter`,
       decompNote:narr.decomp_note||"This decomposition feeds the quarterly NRR / GRR in the metrics table."},
    C:{topRows,topPrevNote:"",topNote:narr.top_note||"",
       concPre:`Top-10 share ${pct(conc.top10_share)}`,concPost:`  •  Top-1 share ${pct(conc.top1_share)} (${conc.top1_name||''})`,
       conc:{top1:conc.top1_share,top2_10:conc.top10_share-conc.top1_share,rest:1-conc.top10_share,headline:pct(conc.top10_share)},
       concReco:narr.concentration_note||"Loss of any single top account is material. Recommend a named-account QBR program and a diversification target (top-10 < 70%).",
       hygiene:narr.hygiene_note||""},
    D:{arrRows:[["Quarter-exit run-rate",usd0(exitMrr),"≈"+usdk(arrCur)],{total:true,cells:["Smoothed (recent complete quarters)",usd0(arrSm/12),"≈"+usdk(arrSm)]}],
       arrReco:narr.arr_recommendation||`Plan against ARR ≈ ${usdk(arrCur)} (quarter-exit run-rate).`,
       pipelineNote:narr.pipeline_note||"Pipeline-to-ARR: not computable until a sales pipeline with open stages exists."},
    E:{cacNote:narr.cac_note||"CAC by channel: not a real CAC until new-logo counts are channel-tagged and fully-loaded S&M cost is captured.",
       toolNote:narr.toolspend_note||`Tool spend vs quarter revenue: ≈ ${pct(toolPct)} — ${toolPct<0.05?'under':'over'} the <5% target.`},
    F:{title:"Cash runway — not computable this cycle",lines:narr.runway_lines||[
       "Missing inputs: opening cash balance and fully-loaded burn (payroll + CDN infrastructure COGS).",
       `On tracked OpEx alone the quarter is cash-generative (${sUsd(net)}), but that excludes people and delivery infra — NOT a runway statement.`]},
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

  // title band
  body.push(TBL([CW],[ROW([CELL([
    TXT(M.meta.company.toUpperCase(),{b:true,color:ORANGE,sz:22,after:30}),
    TXT("Quarterly Financial Report",{b:true,color:WHITE,sz:40,after:20}),
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
  body.push(h2("A.  Quarterly revenue & result — "+M.meta.period_label));
  body.push(P([HALR("Revenue: ",{b:true,sz:20}),HALR(M.A.headline,{b:true,color:NAVY,sz:24}),HALR("    "+M.A.sub,{color:GREY,sz:16})],{after:120}));
  body.push(TXT("Revenue by plan type:",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Plan type","Revenue","Share"],M.A.planRows));
  if(M.A.planNote)body.push(TXT(M.A.planNote,{i:true,sz:15,color:GREY,before:60,after:140}));
  body.push(TXT("Tracked OpEx (tooling, ads, AI only; not a full P&L):",{b:true,sz:17,before:M.A.planNote?undefined:120,after:60}));
  body.push(dataTable([4680,2340,2340],["Category","EUR","USD"],M.A.opexRows));
  body.push(callout("Net result — net of tracked OpEx only (NOT operating profit)",[
    P([HALR(M.A.netEq,{sz:18}),HALR(M.A.netVal,{b:true,color:GREEN,sz:22}),HALR("   •   "+M.A.netSub,{sz:16,color:GREY})]),
    TXT(M.A.netCaveat,{i:true,sz:15,color:GREY,before:40})],{fill:GREENBG,bar:GREEN}));

  // trend bar table (quarterly)
  body.push(TXT("Revenue trend — quarterly",{b:true,sz:18,before:120,after:40}));
  const tcols=[1500,4800,1530,1530];
  const tHead=ROW(["Quarter","","Revenue","QoQ"].map((h,i)=>CELL(TXT(h,{b:true,color:WHITE,sz:16}),{w:tcols[i],fill:NAVY,borderColor:NAVY})),{header:true});
  const tBody=M.A.trend.map(t=>ROW([
    CELL(TXT(t.label,{sz:16}),{w:tcols[0]}),
    CELL(P(blockBar(t.frac,t.color)),{w:tcols[1]}),
    CELL(TXT(t.rev,{sz:16,jc:'right'}),{w:tcols[2]}),
    CELL(TXT(t.mom,{sz:15,jc:'right',color:t.momColor}),{w:tcols[3]})]));
  body.push(TBL(tcols,[tHead,...tBody]));

  // B
  body.push(h2("B.  Trends & comparisons"));
  body.push(callout("QoQ — this quarter vs prior quarter",
    [P([HALR(M.B.yoyPre,{sz:18}),HALR(M.B.yoyVal,{b:true,color:GREEN,sz:20})])],{fill:GREENBG,bar:GREEN}));
  body.push(SPACER(80));
  body.push(TXT("Quarter vs prior quarter (and same quarter last year):",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Window","Value","Change"],M.B.t3mRows));
  if(M.B.t3mNote)body.push(TXT(M.B.t3mNote,{i:true,sz:15,color:GREY,before:60,after:140}));
  body.push(TXT(`Growth decomposition (material clients, ${M.B.decompWindow}):`,{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Movement","Amount","Clients"],
    M.B.decompRows.map(r=>[{t:r.label,color:r.color},{t:r.amt,color:r.color},String(r.n)])
      .concat([{total:true,cells:["Net change",M.B.decompNet,String(M.B.decompMaterial)]}])));
  if(M.B.decompNote)body.push(TXT(M.B.decompNote,{i:true,sz:15,color:GREY,before:60,after:120}));

  // C (Top 10, 5 columns)
  body.push(h2("C.  Customers — Top "+M.C.topRows.length));
  const cCols=[2740,1620,1500,1100,2400];
  const cHead=ROW(["Client","This Q","Prior Q","QoQ","Same Q LY"].map((h,i)=>CELL(TXT(h,{b:true,color:WHITE,sz:18}),{w:cCols[i],fill:NAVY,borderColor:NAVY})),{header:true});
  const cBody=M.C.topRows.map((r,ri)=>{const fill=ri%2?WHITE:LIGHT;
    return ROW([
      CELL(TXT(r.label,{sz:17,color:r.new?ORANGE:undefined}),{w:cCols[0],fill}),
      CELL(P([HALR(r.thisDisp+" ",{sz:16}),...blockBar(r.frac,r.new?ORANGE:NAVY,6)],{jc:'right'}),{w:cCols[1],fill}),
      CELL(TXT(r.prev,{sz:16,jc:'right',color:r.prevColor}),{w:cCols[2],fill}),
      CELL(TXT(r.mom,{sz:16,jc:'right',color:r.momColor}),{w:cCols[3],fill}),
      CELL(TXT(r.yoy,{sz:16,jc:'right',color:r.yoyColor}),{w:cCols[4],fill})]);});
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
  body.push(TXT("ARR snapshot (quarter-exit MRR × 12):",{b:true,sz:17,after:60}));
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
      HALR(`${M.meta.company} — Quarterly Financial Report — ${M.meta.period_label}   •   Confidential   •   Page `,{color:GREY,sz:14}),
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
const inp=process.argv[2]||'input.json', out=process.argv[3]||'Quarterly_Financial_Report.docx';
const data=JSON.parse(fs.readFileSync(inp,'utf8'));
Packer.toBuffer(render(compute(data))).then(buf=>{fs.writeFileSync(out,buf);console.log('✓ wrote',out);});
```

## Example input  (the exact contract — save your real data as input.json in this shape)

```json
{
  "meta": {
    "company":"BlazingCDN","report_type":"quarterly","period_label":"Q1 2026","period":"2026-Q1",
    "prepared_by":"Mike Scarpelli, CFO","date":"2026-04-05",
    "for":"Hermes (Chief of Staff) \u2192 CEO",
    "source_portal":"HubSpot production portal (Hub 143144902), Deals \u2192 \"Current Clients Pipeline\". Expenses: Google Drive \"CFO\"",
    "fx_eur_usd":1.08,"currency":"USD","pre_batch":false
  },
  "inputs": { "prior_q_sm_usd": 3500 },
  "revenue": {
    "focal_quarter":"Q1 2026",
    "unique_clients":118,
    "quarterly_series":[
      {"quarter":"Q1 2025","count":159,"revenue":55194.49},
      {"quarter":"Q2 2025","count":193,"revenue":71800.91},
      {"quarter":"Q3 2025","count":214,"revenue":77045.41},
      {"quarter":"Q4 2025","count":219,"revenue":94478.34},
      {"quarter":"Q1 2026","count":256,"revenue":102774.98}
    ],
    "plan_split":{"recurring":{"revenue":102200.00,"clients":92},"payg":{"revenue":574.98,"clients":7}},
    "top_clients":[
      {"name":"eliaskhan.it","focal":13675.00,"prior_q":11200.00,"yoy_same_q":9800.00,"new_logo":false},
      {"name":"rabin.yor\u2026","focal":12320.00,"prior_q":11050.00,"yoy_same_q":10500.00,"new_logo":false},
      {"name":"itgroup (avanquest)","focal":8200.00,"prior_q":0.00,"yoy_same_q":0.00,"new_logo":true},
      {"name":"a.ruin (epom)","focal":6410.55,"prior_q":3900.00,"yoy_same_q":0.00,"new_logo":true},
      {"name":"yashwanth (moengage)","focal":6000.00,"prior_q":2250.00,"yoy_same_q":0.00,"new_logo":true},
      {"name":"secretfy.corp","focal":5500.00,"prior_q":4000.00,"yoy_same_q":0.00,"new_logo":true},
      {"name":"arik.holaspark","focal":5200.00,"prior_q":4800.00,"yoy_same_q":4100.00,"new_logo":false},
      {"name":"theqooking","focal":3148.02,"prior_q":2900.00,"yoy_same_q":2600.00,"new_logo":false},
      {"name":"manuel.skelton","focal":2103.00,"prior_q":900.00,"yoy_same_q":0.00,"new_logo":true},
      {"name":"bret.cityspark","focal":1753.61,"prior_q":1600.00,"yoy_same_q":1400.00,"new_logo":false}
    ],
    "concentration":{"top10_share":0.629,"top1_share":0.133,"top1_name":"eliaskhan.it"},
    "decomposition":{
      "window":"Q1 2026 vs Q4 2025","material_threshold":500,"material_clients":22,
      "starting_recurring":89000.00,
      "new":{"amount":8200.00,"clients":1},
      "expansion":{"amount":12450.00,"clients":8},
      "contraction":{"amount":-4180.00,"clients":3},
      "churn":{"amount":-1200.00,"clients":1}
    }
  },
  "expenses":{
    "month":"Q1 2026","full_pl":false,"recurring_tool_eur":1581,
    "lines":[
      {"category":"Marketing","detail":"Google Ads quarter","eur":3207.00},
      {"category":"Product & AI","detail":"OpenAI + Claude + Jina","eur":870.00},
      {"category":"Ops & infra","detail":"HubSpot + make + Atlassian + GCP","eur":525.00},
      {"category":"Sales tooling","detail":"SalesQL + Zapmail + QEV","eur":396.00}
    ]
  },
  "cfo_narrative":{
    "bottom_line":[
      "Q1 2026 revenue $102.8K, +8.8% QoQ and +86% YoY \u2014 strongest quarter on record, no logo churn at the material tier except one small loss.",
      "Quarterly NRR and GRR are now computable (no data gap this quarter) \u2014 see metrics table; expansion outpaced contraction.",
      "Concentration eased to 63% top-10 (from 87% on a single-month view) as the book broadened across net-new logos.",
      "Still no full P&L \u2014 Rule-of-40 FCF margin is an OpEx-only proxy until CDN COGS and payroll land.",
      "Working run-rate: ARR \u2248 $411K (quarter-exit)."
    ],
    "top_note":"Net-new logos landed over the past year now populate half the top 10 \u2014 the diversification engine is working.",
    "concentration_note":"Top-10 down to 63% \u2014 healthier than the single-month view. Keep the QBR program on the top 10 and hold the <70% target.",
    "decomp_note":"Expansion (+$12.5K, 8 clients) outpaced contraction (\u2212$4.2K) and the single small churn (\u2212$1.2K). This feeds the quarterly NRR/GRR below.",
    "arr_recommendation":"Plan against ARR \u2248 $411K (quarter-exit run-rate); the trend is up four straight quarters."
  }
}
```
