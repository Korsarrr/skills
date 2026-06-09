---
name: cfo-quarterly-report
description: "Generate the polished BlazingCDN QUARTERLY financial report as a formatted Word (.docx): same house style as the monthly report but quarter-framed, with QoQ + YoY-by-quarter comparisons, a quarterly revenue trend bar-table, a Top-10 customers table (This Q / Prior Q / QoQ% / Same-Q-LY), a Top-10 concentration bar, and the quarter-only metrics promoted to first-class: NRR, GRR, Magic Number, Rule of 40. This skill is a single self-contained markdown file: it embeds a zero-dependency Node ESM engine (no npm, no Python, no images) plus the input contract. Use at quarter close once the data workers have returned quarterly aggregates. Do NOT use for monthly reports (use cfo-monthly-report)."
---

# CFO Quarterly Financial Report  (single-file .md skill)

Everything needed is inside THIS markdown file: the same zero-dependency render engine as the monthly skill
(Node standard library only — no `docx` npm, no Python, no images) and the input contract. Separate skill so
quarterly logic never loads on a monthly run. Deterministic: same engine → identical design; only numbers change.

## How to produce a report (runtime steps)
1. **Materialize the engine.** Copy the entire `build.mjs` code block (under "Engine" below) verbatim into a
   file named `build.mjs` in your workspace.
2. **Assemble the input.** Build `input.json` from the two data workers' quarterly JSON + your `cfo_narrative`,
   using the "Example input" block below as the exact contract. Include `decomposition.starting_recurring` so
   NRR/GRR compute, and optionally `inputs.prior_q_sm_usd` (prior-quarter S&M) so the Magic Number computes.
3. **Run:** `node build.mjs input.json "BlazingCDN_Quarterly_Financial_Report_<Q_Year>.docx"` (needs only `node`
   v18+; no `npm install`).
4. Review, save to Google Drive "CFO", post the link to Hermes.

## What differs from monthly
- Trend bar-table and tables are by **quarter**.
- Section B is **QoQ** (this quarter vs prior quarter) plus same-quarter-last-year.
- Section C is the **Top 10** with This Q / Prior Q / QoQ% / Same Q LY.
- Metrics compute **NRR / GRR** (from `starting_recurring`), **Magic Number** (net-new ARR ÷ prior-Q S&M), and
  frame **Rule of 40** (FCF margin stays an OpEx-only proxy until a full P&L exists). ARR = quarter-exit MRR × 12.

## Design behaviour (automatic)
- **Signed deltas carry arrows:** positive → green ▲, negative → red ▼, exact zero → neutral (no arrow). Applied
  to KPI QoQ/YoY, per-client QoQ, trend QoQ, the comparison-table Change column, growth-decomposition amounts,
  and net-new ARR. The Rule-of-40 growth value and Section F narrative prose are left arrow-free on purpose.
- **Footer shows `Page X of Y`** (PAGE / NUMPAGES fields).
- The full visual spec lives in the engine — see the DESIGN INVARIANTS comment at the top of `build.mjs` (palette,
  image-free, deterministic). House style is identical to `cfo-monthly-report`; there is no separate spec doc to drift.

## NOTE — production Hub ID
Canonical production Hub ID is **143144902**, already baked into `meta.source_portal` in the example below (the prior
`143144902` vs `145006611` ambiguity is resolved — `145006611` was the non-production portal). Confirm it still matches
your live portal before each run; change the single value in `meta.source_portal` if the portal ever moves.

## Engine — build.mjs  (copy verbatim into a file named build.mjs)

```javascript
#!/usr/bin/env node
/*
 * cfo-quarterly-report / build.mjs  — ZERO external dependencies (Node stdlib only).
 * No npm packages, no Python. Writes a valid .docx via raw OOXML + a hand-built ZIP.
 *   node build.mjs <input.json> <output.docx>
 * Does both compute (metrics) and render. Charts are drawn with shaded/block-bar cells
 * (no images), so nothing extra is required at runtime.
 *
 * DESIGN INVARIANTS (the look lives here, in code — not in a separate spec doc):
 *   palette : NAVY 16365C (primary) · ORANGE E8590C (accent) · GREEN 2E7D32 / RED C62828 (deltas)
 *   rules   : image-free · deterministic (same engine -> identical design each quarter) · stdlib-only
 *   deltas  : green up-triangle for positive, red down-triangle for negative, zero = neutral (no arrow)
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
const arrow=x=>x>0?"\u25B2 ":(x<0?"\u25BC ":"");           // ▲ pos / ▼ neg / "" zero
const dUsd=x=>arrow(x)+sUsd(x);                            // delta $ with arrow
const dPct=x=>arrow(x)+sPct(x);                            // delta % with arrow
const shortM=m=>{const[a,b]=String(m).split("-");if(!b)return m;const M=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];return `${M[+b-1]}'${a.slice(2)}`;};

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

/* ============================ render ============================ */
function render(M){
  const body=[];
  // title band
  body.push(table([CW],row(cell(
    txt(M.meta.company.toUpperCase(),{b:true,color:ORANGE,sz:22,after:30})+
    txt("Quarterly Financial Report",{b:true,color:WHITE,sz:40,after:20})+
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
  body.push(h2("A.  Quarterly revenue & result — "+M.meta.period_label));
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
  body.push(txt("Revenue trend — quarterly",{b:true,sz:18,before:120,after:40}));
  const trendRows=M.A.trend.map(t=>row(
    cell(txt(t.label,{sz:16}),{w:1500})+
    cell(para(blockBar(t.frac,t.color)),{w:4800})+
    cell(txt(t.rev,{sz:16,jc:'right'}),{w:1530})+
    cell(txt(t.mom,{sz:15,jc:'right',color:t.momColor}),{w:1530})
  )).join("");
  body.push(table([1500,4800,1530,1530],
    row(["Quarter","","Revenue","QoQ"].map((h,i)=>cell(txt(h,{b:true,color:WHITE,sz:16}),{w:[1500,4800,1530,1530][i],fill:NAVY,borderColor:NAVY})).join(""),{header:true})+trendRows));
  // B
  body.push(h2("B.  Trends & comparisons"));
  body.push(callout("QoQ — this quarter vs prior quarter",
    para(run(M.B.yoyPre,{sz:18})+run(M.B.yoyVal,{b:true,color:GREEN,sz:20})),{fill:GREENBG,bar:GREEN}));
  body.push(txt(" ",{after:80}));
  body.push(txt("Quarter vs prior quarter (and same quarter last year):",{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Window","Value","Change"],M.B.t3mRows));
  if(M.B.t3mNote)body.push(txt(M.B.t3mNote,{i:true,sz:15,color:GREY,before:60,after:140}));
  body.push(txt(`Growth decomposition (material clients, ${M.B.decompWindow}):`,{b:true,sz:17,after:60}));
  body.push(dataTable([4680,2340,2340],["Movement","Amount","Clients"],
    M.B.decompRows.map(r=>[{t:r.label,color:r.color},{t:r.amt,color:r.color},String(r.n)])
      .concat([{total:true,cells:["Net change",M.B.decompNet,String(M.B.decompMaterial)]}])));
  if(M.B.decompNote)body.push(txt(M.B.decompNote,{i:true,sz:15,color:GREY,before:60,after:120}));
  // C
  body.push(h2("C.  Customers — Top "+M.C.topRows.length));
  const cCols=[2740,1620,1500,1100,2400];
  const cHead=row(["Client","This Q","Prior Q","QoQ","Same Q LY"].map((h,i)=>cell(txt(h,{b:true,color:WHITE,sz:18}),{w:cCols[i],fill:NAVY,borderColor:NAVY})).join(""),{header:true});
  const cBody=M.C.topRows.map((r,ri)=>{const fill=ri%2?WHITE:LIGHT;
    return row(
      cell(txt(r.label,{sz:17,color:r.new?ORANGE:undefined}),{w:cCols[0],fill})+
      cell(para(run(r.thisDisp+" ",{sz:16,jc:'right'})+blockBar(r.frac,r.new?ORANGE:NAVY,6),{jc:'right'}),{w:cCols[1],fill})+
      cell(txt(r.prev,{sz:16,jc:'right',color:r.prevColor}),{w:cCols[2],fill})+
      cell(txt(r.mom,{sz:16,jc:'right',color:r.momColor}),{w:cCols[3],fill})+
      cell(txt(r.yoy,{sz:16,jc:'right',color:r.yoyColor}),{w:cCols[4],fill})
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
  body.push(txt("ARR snapshot (quarter-exit MRR × 12):",{b:true,sz:17,after:60}));
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
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="6" w:color="D9DEE4"/></w:pBdr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="6B7280"/><w:sz w:val="14"/></w:rPr><w:t xml:space="preserve">${esc(M.meta.company)} — Quarterly Financial Report — ${esc(M.meta.period_label)}   •   Confidential   •   Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:color w:val="6B7280"/><w:sz w:val="14"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple><w:r><w:rPr><w:color w:val="6B7280"/><w:sz w:val="14"/></w:rPr><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:rPr><w:color w:val="6B7280"/><w:sz w:val="14"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`;

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
const inp=process.argv[2]||'input.json', out=process.argv[3]||'Quarterly_Financial_Report.docx';
const data=JSON.parse(fs.readFileSync(inp,'utf8'));
fs.writeFileSync(out, render(compute(data)));
console.log('✓ wrote', out);
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
