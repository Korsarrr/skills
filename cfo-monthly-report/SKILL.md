---
name: cfo-monthly-report
description: "Generate the polished BlazingCDN MONTHLY financial report as a formatted Word (.docx): title band, KPI cards, data-integrity callouts, formatted tables with status colors, a monthly revenue trend bar-table, a Top-20 customers table with This-mo / Prev-mo / MoM% / T3M / Same-month-LY comparisons, and a Top-10 concentration bar. Runs as a single zero-dependency Node ESM file (build.mjs) — no npm packages, no Python, no images. Use on the 5th-of-month cycle once the HubSpot Deals Worker (revenue JSON) and Subscription Tracker Worker (expense JSON) have returned data and the CFO has assembled input.json. Do NOT use for quarterly reports (use cfo-quarterly-report)."
---

# CFO Monthly Financial Report  (zero-dependency .mjs)

Deterministic render: same code → identical design every month; only the numbers in `input.json` change.
`build.mjs` does BOTH the metric math and the Word rendering. It uses **only Node's standard library** —
no `docx` npm package, no matplotlib, no Python, no separate scripts, no network. Charts are drawn as
shaded/▆block-bar table cells (no images), so nothing extra is needed at runtime.

## Files in this skill
- `build.mjs`  — the whole engine (compute + render + a hand-built .docx writer). **Run this.**
- `schema/example_input.json` — the exact input contract (May 2026 worked example).
- (this `SKILL.md` also embeds the full `build.mjs` source at the bottom as a fallback — see "If .mjs won't load").

## When to run
On the **5th-of-month** cycle, AFTER the billing batch loads and both data workers have returned:
HubSpot Deals Worker → revenue data; Subscription Tracker Worker → expense data.

## What the CFO does (the only manual part)
1. Assemble `input.json` from the two workers' JSON + add the qualitative `cfo_narrative`
   (bottom_line, notes). See `schema/example_input.json` for the contract. The CFO writes only judgement;
   all math (ARR/MRR, YoY, T3M + backfill, net result, margins, tool-spend %, MoM per client, concentration,
   integrity-flag auto-detection, metric statuses) is done by build.mjs.
2. Run it (below).
3. Review, save to Google Drive "CFO", post the link to Hermes.

## Run
```bash
node build.mjs <input.json> <output.docx>
# e.g. node build.mjs input.json "BlazingCDN_Monthly_Financial_Report_May_2026.docx"
```
Requires only `node` (v18+). No `npm install`. No Python.

## Section C — Top 20 customers
Top-20 table with four comparisons per client so new logos and patterns pop:
**This mo · Prev mo (previous complete billed month) · MoM% (green ▲ / red ▼) · T3M · Same mo last year**.
Net-new logos (YoY) are flagged ★ and coloured; the "This mo" cell carries a little inline bar.
If the literal previous month is a gap, Prev/MoM read "—" and a footnote names the skipped month.
**Concentration stays Top-10** (Top-1 / Top 2–10 / Rest stacked bar) — that risk metric is unchanged.

## Integrity behaviour (automatic)
- Any $0 month in the series is auto-flagged as a probable missing batch; trailing metrics spanning it are
  marked PROVISIONAL.
- If `expenses.full_pl` is false, margins are labelled "net of tracked OpEx," not GAAP.
- If `meta.pre_batch` is true, the report is labelled provisional with the close date.

## NOTE — production Hub ID
The example uses a `<CONFIRM_PROD_HUB_ID>` placeholder (open ambiguity 143144902 vs 145006611). Put the single
canonical production Hub ID in `meta.source_portal` before running.

## If .mjs won't load in your PaperClip (md + json only)
PaperClip loads this `SKILL.md` (markdown) regardless. The complete `build.mjs` source is reproduced verbatim
below. At runtime, write it to a file named `build.mjs` in the workspace and run `node build.mjs input.json out.docx`.
It is byte-identical to the `build.mjs` file shipped alongside.

<details><summary>build.mjs — full source (copy verbatim to a file, then run with node)</summary>

```javascript
#!/usr/bin/env node
/*
 * cfo-monthly-report / build.mjs  — ZERO external dependencies (Node stdlib only).
 * No npm packages, no Python. Writes a valid .docx via raw OOXML + a hand-built ZIP.
 *   node build.mjs <input.json> <output.docx>
 * Does both compute (metrics) and render. Charts are drawn with shaded/block-bar cells
 * (no images), so nothing extra is required at runtime.
 */
import fs from 'node:fs';

/* ============================ ZIP (STORE) + CRC32 ============================ */
const CRCT=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0;}return t;})();
const crc32=b=>{let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRCT[(c^b[i])&0xFF]^(c>>>8);return (c^0xFFFFFFFF)>>>0;};
function zip(files){
  const out=[],cen=[];let off=0;
  for(const f of files){
    const name=Buffer.from(f.name,'utf8'),data=f.data,crc=crc32(data);
    const lh=Buffer.alloc(30);lh.writeUInt32LE(0x04034b50,0);lh.writeUInt16LE(20,4);
    lh.writeUInt32LE(crc,14);lh.writeUInt32LE(data.length,18);lh.writeUInt32LE(data.length,22);
    lh.writeUInt16LE(name.length,26);out.push(lh,name,data);
    const ch=Buffer.alloc(46);ch.writeUInt32LE(0x02014b50,0);ch.writeUInt16LE(20,4);ch.writeUInt16LE(20,6);
    ch.writeUInt32LE(crc,16);ch.writeUInt32LE(data.length,20);ch.writeUInt32LE(data.length,24);
    ch.writeUInt16LE(name.length,28);ch.writeUInt32LE(off,42);cen.push(ch,name);
    off+=lh.length+name.length+data.length;
  }
  let cdSize=0;for(const c of cen)cdSize+=c.length;
  const end=Buffer.alloc(22);end.writeUInt32LE(0x06054b50,0);end.writeUInt16LE(files.length,8);
  end.writeUInt16LE(files.length,10);end.writeUInt32LE(cdSize,12);end.writeUInt32LE(off,16);
  return Buffer.concat([...out,...cen,end]);
}
const Bx=s=>Buffer.from(s,'utf8');

/* ============================ palette ============================ */
const NAVY="16365C",ORANGE="E8590C",GREEN="2E7D32",AMBER="B26A00",RED="C62828",GREY="6B7280",
  LIGHT="EEF2F6",AMBERBG="FCF3E2",GREENBG="E9F3EA",GREYBG="F1F2F4",WHITE="FFFFFF",TRACK="DCE2E8";
const CW=9360;

/* ============================ OOXML helpers ============================ */
const esc=s=>String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
function run(text,o={}){
  let r="<w:rPr>";
  if(o.b)r+="<w:b/>"; if(o.i)r+="<w:i/>";
  if(o.color)r+=`<w:color w:val="${o.color}"/>`;
  if(o.sz)r+=`<w:sz w:val="${o.sz}"/><w:szCs w:val="${o.sz}"/>`;
  r+="</w:rPr>";
  return `<w:r>${r}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}
function para(inner,o={}){
  let p="<w:pPr>";
  if(o.style)p+=`<w:pStyle w:val="${o.style}"/>`;
  if(o.keepNext)p+="<w:keepNext/>";
  if(o.bdrBottom)p+=`<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="${o.bdrBottom}"/></w:pBdr>`;
  if(o.before!=null||o.after!=null)p+=`<w:spacing${o.before!=null?` w:before="${o.before}"`:''}${o.after!=null?` w:after="${o.after}"`:''}/>`;
  if(o.ind)p+=`<w:ind w:left="${o.ind.left||0}"${o.ind.hanging?` w:hanging="${o.ind.hanging}"`:''}/>`;
  if(o.jc)p+=`<w:jc w:val="${o.jc}"/>`;
  p+="</w:pPr>";
  return `<w:p>${p}${inner}</w:p>`;
}
const txt=(text,o={})=>para(run(text,o),o);
function pageBreak(){return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;}

function tcBorders(c="CCCCCC",leftBar){
  const e=(s,col)=>`<w:${s} w:val="single" w:sz="${s==='left'&&leftBar?18:4}" w:space="0" w:color="${s==='left'&&leftBar?leftBar:col}"/>`;
  return `<w:tcBorders>${e('top',c)}${e('left',c)}${e('bottom',c)}${e('right',c)}</w:tcBorders>`;
}
function cell(inner,o={}){
  let pr=`<w:tcPr><w:tcW w:w="${o.w||0}" w:type="dxa"/>`;
  if(o.noBorder)pr+=`<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>`;
  else pr+=tcBorders(o.borderColor||"CCCCCC",o.leftBar);
  if(o.fill)pr+=`<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>`;
  const m=o.mar||{t:60,b:60,l:120,r:120};
  pr+=`<w:tcMar><w:top w:w="${m.t}" w:type="dxa"/><w:left w:w="${m.l}" w:type="dxa"/><w:bottom w:w="${m.b}" w:type="dxa"/><w:right w:w="${m.r}" w:type="dxa"/></w:tcMar>`;
  pr+=`<w:vAlign w:val="${o.valign||'center'}"/></w:tcPr>`;
  return `<w:tc>${pr}${inner}</w:tc>`;
}
function row(cells,o={}){
  let pr='';
  if(o.cantSplit!==false)pr+='<w:cantSplit/>';
  if(o.header)pr+='<w:tblHeader/>';
  return `<w:tr><w:trPr>${pr}</w:trPr>${cells}</w:tr>`;
}
function table(cols,rows,o={}){
  const grid=cols.map(w=>`<w:gridCol w:w="${w}"/>`).join("");
  const borders=o.noBorder?`<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>`:"";
  return `<w:tbl><w:tblPr><w:tblW w:w="${o.w||CW}" w:type="dxa"/>${borders}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

/* convenience builders */
function dataTable(cols,headers,rows){
  const head=row(headers.map((h,i)=>cell(txt(h,{b:true,color:WHITE,sz:18}),{w:cols[i],fill:NAVY,borderColor:NAVY})).join(""),{header:true});
  const body=rows.map((r,ri)=>{
    const isTotal=r&&r.total; const cells=isTotal?r.cells:r;
    return row(cells.map((c,ci)=>{
      const obj=c&&typeof c==="object"; const text=obj?c.t:c;
      const fill=isTotal?LIGHT:(ri%2?WHITE:LIGHT);
      const jc=ci>0?'right':undefined;
      return cell(txt(text,{sz:18,b:isTotal||(obj&&c.b),color:obj?c.color:undefined,jc}),{w:cols[ci],fill});
    }).join(""));
  }).join("");
  return table(cols,head+body);
}
function callout(title,bodyParas,{fill=AMBERBG,bar=AMBER}={}){
  const inner=para(run(title,{b:true,color:bar,sz:20}),{after:60})+bodyParas;
  return table([CW],row(cell(inner,{w:CW,fill,leftBar:bar,borderColor:"E0D6BE",mar:{t:120,b:120,l:200,r:160}})));
}
function kpiCard(k,w){
  const inner=txt(k.value,{b:true,color:NAVY,sz:30,jc:'center',after:20})+
    txt(k.label,{color:GREY,sz:15,jc:'center',after:40})+
    txt(k.status,{b:true,color:k.color,sz:15,jc:'center'});
  return cell(inner,{w,fill:LIGHT,borderColor:"D9DEE4",mar:{t:120,b:120,l:80,r:80}});
}
const h2=t=>para(run(t,{}),{style:'Heading2',keepNext:true,bdrBottom:NAVY});
function listItem(text,prefix){return txt(prefix+text,{sz:18,ind:{left:360,hanging:300}});}

/* block bar (filled █ + track ░), returns run-xml */
function blockBar(frac,color,max=14){
  frac=Math.max(0,Math.min(1,frac)); const n=Math.max(1,Math.round(frac*max));
  return run("\u2588".repeat(n),{color,sz:14})+run("\u2591".repeat(max-n),{color:TRACK,sz:14});
}

/* ============================ compute ============================ */
const usd=x=>`$${x.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const usd0=x=>`$${Math.round(x).toLocaleString('en-US')}`;
const usdk=x=>`$${(x/1000).toFixed(1)}K`;
const pct=x=>`${(x*100).toFixed(1)}%`;
const sUsd=x=>(x>=0?"+":"\u2212")+`$${Math.abs(x).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const sPct=x=>(x>=0?"+":"\u2212")+`${Math.abs(x*100).toFixed(1)}%`;
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
    {value:yoy?sPct(yoy.pct):"n/a",label:`Revenue YoY (${meta.period_label.split(' ')[0]})`,status:"Clean — trust",color:GREEN},
    {value:pct(toolPct),label:"Tool spend / revenue",status:toolPct<0.05?"\u2713 Under 5% target":"\u26A0 Over 5% target",color:toolPct<0.05?GREEN:AMBER},
    {value:pct(conc.top10_share),label:"Top-10 concentration",status:conc.top10_share>0.70?"\u26A0 High risk":"OK",color:AMBER},
  ];

  // top-N
  const prevLit=idx>0?months[idx-1]:null, prevGap=prevLit?zero.includes(prevLit):false;
  const topPrevNote=prevGap?`* Prev mo. = previous complete billed month; ${prevLit} was skipped (missing batch).`:"";
  const maxFocal=Math.max(...rev.top_clients.map(c=>c.focal));
  const topRows=rev.top_clients.map((c,i)=>{
    const pm=c.prev_month||0; let mom,momColor,prevDisp,prevColor;
    if(pm>0){const mv=(c.focal-pm)/pm;mom=sPct(mv);momColor=mv>=0?GREEN:RED;prevDisp=usd(pm);prevColor=null;}
    else{mom="\u2014";momColor=GREY;prevDisp="\u2014";prevColor=GREY;}
    return {new:!!c.new_logo,label:`${i+1}  ${c.name}${c.new_logo?'  \u2605 new':''}`,
      this:c.focal,thisDisp:usd(c.focal),frac:c.focal/maxFocal,
      prev:prevDisp,prevColor,mom,momColor,t3m:usd(c.t3m),
      yoy:usd(c.yoy_same_month||0),yoyColor:(c.yoy_same_month||0)===0?GREY:null};
  });

  const dc=rev.decomposition;
  const decompRows=[
    {label:"New",color:GREEN,amt:sUsd(dc.new.amount),n:dc.new.clients},
    {label:"Expansion",color:GREEN,amt:sUsd(dc.expansion.amount),n:dc.expansion.clients},
    {label:"Contraction",color:RED,amt:sUsd(dc.contraction.amount),n:dc.contraction.clients},
    {label:"Churn (full logo loss)",color:dc.churn.clients===0?GREEN:RED,amt:sUsd(dc.churn.amount),n:dc.churn.clients},
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
      mom:mom==null?"—":sPct(mom),momColor:mom==null?GREY:(mom>=0?GREEN:RED)};
  });

  return {meta,fx,kpis,flags,bottom_line:narr.bottom_line||[],
    A:{headline:usd(fv),sub:`${fcount} billing events  •  ${rev.unique_clients??'?'} unique paying clients`,
       planRows,planNote:narr.plan_note||"Recommendation: add a plan_type deal property so the split is exact, not heuristic.",
       opexRows,netEq:`${usd(fv)} − ${usd(usdTot)} = `,netVal:usd(net),netSub:`${pct(contrib)} contribution after tracked tooling/marketing`,
       netCaveat:"This is not a gross or operating margin — CDN delivery COGS and payroll are absent from the data room.",trend},
    B:{yoyPre:yoy?`${usd(fv)} vs ${usd(yoy.prior)}   →   `:"",yoyVal:yoy?`${sUsd(yoy.abs)}  /  ${sPct(yoy.pct)}`:"n/a",
       t3mRows:[[`T3M (${win.map(m=>m.month.slice(-2)+'='+usdk(m.revenue)).join(' + ')})`,usd(t3m),""],
                ["Prior T3M",usd(pt3m),{t:sPct(t3mCh),color:t3mCh<0?RED:GREEN}],
                ...(gapInWin?[{total:true,cells:["With backfill (≈ trailing avg)","≈"+usdk(t3mBf),{t:"≈"+sPct(t3mBfCh),color:t3mBfCh>=0?GREEN:RED}]}]:[])],
       t3mNote:narr.t3m_note||(gapInWin?"The headline change is an artifact of the missing batch — do not act on it until the gap is loaded.":""),
       decompRows,decompNet:sUsd(netChange),decompMaterial:dc.material_clients??rev.top_clients.length,decompWindow:dc.window||"",
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

/* ============================ render ============================ */
function render(M){
  const body=[];
  // title band
  body.push(table([CW],row(cell(
    txt(M.meta.company.toUpperCase(),{b:true,color:ORANGE,sz:22,after:30})+
    txt("Monthly Financial Report",{b:true,color:WHITE,sz:40,after:20})+
    txt("Reporting period: "+M.meta.period_label,{color:"C9D4E3",sz:22}),
    {w:CW,fill:NAVY,noBorder:true,mar:{t:200,b:160,l:240,r:240}}),{cantSplit:true}),{noBorder:true}));
  body.push(txt(`Prepared by ${M.meta.prepared_by}   •   Date ${M.meta.date}   •   For: ${M.meta.for}`,{sz:16,color:GREY,before:120,after:20}));
  body.push(txt("Source of truth: "+M.meta.source_portal+".",{sz:15,color:GREY,i:true,after:160}));
  // KPI strip
  const cw4=[2340,2340,2340,2340];
  body.push(table(cw4,row(M.kpis.map((k,i)=>kpiCard(k,cw4[i])).join("")),{noBorder:true}));
  body.push(txt(" ",{after:120}));
  // integrity
  body.push(callout("\u26A0  Data integrity — read this first",M.flags.map((f,i)=>listItem(f,`${i+1}. `)).join(""),{fill:AMBERBG,bar:AMBER}));
  body.push(txt(" ",{after:120}));
  // bottom line
  body.push(h2("CFO bottom line (for Hermes → CEO)"));
  M.bottom_line.forEach((t,i)=>body.push(listItem(t,`${i+1}. `)));
  body.push(txt(" ",{after:80}));
  // A
  body.push(h2("A.  Monthly revenue & result — "+M.meta.period_label));
  body.push(para(run("Revenue: ",{b:true,sz:20})+run(M.A.headline,{b:true,color:NAVY,sz:24})+run("    "+M.A.sub,{color:GREY,sz:16}),{after:120}));
  body.push(txt("Revenue by plan type:",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Plan type","Revenue","Share"],M.A.planRows));
  body.push(txt(M.A.planNote,{i:true,sz:15,color:GREY,before:60,after:140}));
  body.push(txt("Tracked OpEx (tooling, ads, AI only; not a full P&L):",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Category","EUR","USD"],M.A.opexRows));
  body.push(callout("Net result — net of tracked OpEx only (NOT operating profit)",
    para(run(M.A.netEq,{sz:18})+run(M.A.netVal,{b:true,color:GREEN,sz:22})+run("   •   "+M.A.netSub,{sz:16,color:GREY}))+
    txt(M.A.netCaveat,{i:true,sz:15,color:GREY,before:40}),{fill:GREENBG,bar:GREEN}));
  // trend (bar table)
  body.push(txt("Revenue trend — monthly",{b:true,sz:18,before:120,after:40}));
  const trendRows=M.A.trend.map(t=>row(
    cell(txt(t.label,{sz:16}),{w:1100})+
    cell(para(blockBar(t.frac,t.color)),{w:5200})+
    cell(txt(t.rev,{sz:16,jc:'right'}),{w:1530})+
    cell(txt(t.mom,{sz:15,jc:'right',color:t.momColor}),{w:1530})
  )).join("");
  body.push(table([1100,5200,1530,1530],
    row(["Month","","Revenue","MoM"].map((h,i)=>cell(txt(h,{b:true,color:WHITE,sz:16}),{w:[1100,5200,1530,1530][i],fill:NAVY,borderColor:NAVY})).join(""),{header:true})+trendRows));
  // B
  body.push(h2("B.  Trends & comparisons"));
  body.push(callout("YoY (month) vs same month last year  •  CLEAN, trust this",
    para(run(M.B.yoyPre,{sz:18})+run(M.B.yoyVal,{b:true,color:GREEN,sz:20})),{fill:GREENBG,bar:GREEN}));
  body.push(txt(" ",{after:80}));
  body.push(txt("3-month block (T3M) vs prior T3M:",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Window","Value","Change"],M.B.t3mRows));
  if(M.B.t3mNote)body.push(txt(M.B.t3mNote,{i:true,sz:15,color:GREY,before:60,after:140}));
  body.push(txt(`Growth decomposition (material clients, ${M.B.decompWindow}):`,{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Movement","Amount","Clients"],
    M.B.decompRows.map(r=>[{t:r.label,color:r.color},{t:r.amt,color:r.color},String(r.n)])
      .concat([{total:true,cells:["Net change",M.B.decompNet,String(M.B.decompMaterial)]}])));
  if(M.B.decompNote)body.push(txt(M.B.decompNote,{i:true,sz:15,color:GREY,before:60,after:120}));
  // C
  body.push(h2("C.  Customers — Top "+M.C.topRows.length));
  const cCols=[2520,1500,1320,820,1500,1700];
  const cHead=row(["Client","This mo","Prev mo*","MoM","T3M","Same mo LY"].map((h,i)=>cell(txt(h,{b:true,color:WHITE,sz:18}),{w:cCols[i],fill:NAVY,borderColor:NAVY})).join(""),{header:true});
  const cBody=M.C.topRows.map((r,ri)=>{const fill=ri%2?WHITE:LIGHT;
    return row(
      cell(txt(r.label,{sz:17,color:r.new?ORANGE:undefined}),{w:cCols[0],fill})+
      cell(para(run(r.thisDisp+" ",{sz:16,jc:'right'})+blockBar(r.frac,r.new?ORANGE:NAVY,6),{jc:'right'}),{w:cCols[1],fill})+
      cell(txt(r.prev,{sz:16,jc:'right',color:r.prevColor}),{w:cCols[2],fill})+
      cell(txt(r.mom,{sz:16,jc:'right',color:r.momColor}),{w:cCols[3],fill})+
      cell(txt(r.t3m,{sz:16,jc:'right'}),{w:cCols[4],fill})+
      cell(txt(r.yoy,{sz:16,jc:'right',color:r.yoyColor}),{w:cCols[5],fill})
    );}).join("");
  body.push(table(cCols,cHead+cBody));
  if(M.C.topPrevNote)body.push(txt(M.C.topPrevNote,{i:true,sz:14,color:GREY,before:50}));
  if(M.C.topNote)body.push(txt(M.C.topNote,{sz:16,color:GREEN,before:60,after:80}));
  // concentration stacked bar (top-10 kept)
  body.push(txt("Revenue concentration (top-10)",{b:true,sz:17,before:120,after:40}));
  const c=M.C.conc, w1=Math.round(CW*c.top1),w2=Math.round(CW*c.top2_10),w3=CW-w1-w2;
  body.push(table([w1,w2,w3],row(
    cell(txt(`Top-1  ${pct(c.top1)}`,{b:true,color:WHITE,sz:14,jc:'center'}),{w:w1,fill:ORANGE,borderColor:WHITE})+
    cell(txt(`Top 2–10  ${pct(c.top2_10)}`,{b:true,color:WHITE,sz:14,jc:'center'}),{w:w2,fill:NAVY,borderColor:WHITE})+
    cell(txt(`Rest  ${pct(c.rest)}`,{color:GREY,sz:14,jc:'center'}),{w:w3,fill:LIGHT,borderColor:WHITE})),{noBorder:true}));
  body.push(callout("\u26A0  Concentration risk",
    para(run(M.C.concPre,{b:true,sz:18})+run(M.C.concPost,{sz:16}))+txt(M.C.concReco,{sz:16,before:40}),{fill:AMBERBG,bar:AMBER}));
  if(M.C.hygiene)body.push(txt(M.C.hygiene,{i:true,sz:14,color:GREY,before:60,after:80}));
  // D
  body.push(h2("D.  Run-rate & forward"));
  body.push(txt("ARR snapshot (recurring book, so MRR ≈ monthly revenue):",{b:true,sz:17,after:60}));
  body.push(dataTable([5560,1900,1900],["Basis","MRR","ARR"],M.D.arrRows));
  body.push(txt(M.D.arrReco,{sz:16,before:60,after:60}));
  body.push(listItem(M.D.pipelineNote,"•  "));
  body.push(txt(" ",{after:60}));
  // E
  body.push(h2("E.  Unit economics & efficiency"));
  body.push(listItem(M.E.cacNote,"•  ")); body.push(listItem(M.E.toolNote,"•  "));
  body.push(txt(" ",{after:60}));
  // F
  body.push(h2("F.  Cash & runway"));
  body.push(callout(M.F.title,M.F.lines.map((l,i)=>txt(l,{sz:16,before:i?40:undefined})).join(""),{fill:GREYBG,bar:GREY}));
  body.push(txt(" ",{after:80}));
  // metrics
  body.push(h2("SaaS metrics summary"));
  body.push(dataTable([2860,3100,1700,1700],["Metric","Value","Target","Status"],M.metrics.map(r=>[r[0],r[1],r[2],r[3]])));
  body.push(txt("— End of report —",{jc:'center',color:GREY,sz:15,before:240}));

  const sectPr=`<w:sectPr><w:footerReference w:type="default" r:id="rId2"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1440" w:bottom="1080" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
  const document=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body.join("")}${sectPr}</w:body></w:document>`;

  const footer=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="6" w:color="D9DEE4"/></w:pBdr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="6B7280"/><w:sz w:val="14"/></w:rPr><w:t xml:space="preserve">${esc(M.meta.company)} — Monthly Financial Report — ${esc(M.meta.period_label)}   •   Confidential   •   Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:color w:val="6B7280"/><w:sz w:val="14"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`;

  const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="16365C"/><w:sz w:val="26"/></w:rPr></w:style></w:styles>`;

  const ct=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`;
  const rootRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
  const docRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`;

  return zip([
    {name:'[Content_Types].xml',data:Bx(ct)},
    {name:'_rels/.rels',data:Bx(rootRels)},
    {name:'word/document.xml',data:Bx(document)},
    {name:'word/styles.xml',data:Bx(styles)},
    {name:'word/footer1.xml',data:Bx(footer)},
    {name:'word/_rels/document.xml.rels',data:Bx(docRels)},
  ]);
}

/* ============================ main ============================ */
const inp=process.argv[2]||'input.json', out=process.argv[3]||'Monthly_Financial_Report.docx';
const data=JSON.parse(fs.readFileSync(inp,'utf8'));
fs.writeFileSync(out, render(compute(data)));
console.log('✓ wrote', out);
```

</details>
