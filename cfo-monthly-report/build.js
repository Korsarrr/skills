// cfo-monthly-report / build.js
// Usage: node build.js <computed.json> <chartsDir> <output.docx>
const fs=require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak } = require("docx");

const C = JSON.parse(fs.readFileSync(process.argv[2]||"computed.json","utf-8"));
const DIR = process.argv[3]||"assets";
const OUT = process.argv[4]||"report.docx";

const NAVY="16365C",ORANGE="E8590C",GREEN="2E7D32",AMBER="B26A00",RED="C62828",GREY="6B7280",
  LIGHT="EEF2F6",AMBERBG="FCF3E2",GREENBG="E9F3EA",GREYBG="F1F2F4",WHITE="FFFFFF";
const CW=9360;
const b=(c="CCCCCC")=>({style:BorderStyle.SINGLE,size:1,color:c});
const allB=(c="CCCCCC")=>({top:b(c),bottom:b(c),left:b(c),right:b(c)});
const noB={top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}};
const CMAR={top:60,bottom:60,left:120,right:120};
const run=(t,o={})=>new TextRun({text:t,bold:o.bold?true:undefined,italics:o.italics?true:undefined,color:o.color,size:o.size,font:o.font});
function txt(text,o={}){const k=Array.isArray(text)?text:[run(text,o)];
  return new Paragraph({children:k,alignment:o.align,spacing:o.spacing,border:o.border,indent:o.indent,keepNext:o.keepNext});}
function cell(ch,o={}){return new TableCell({borders:o.borders??allB(o.borderColor),
  width:{size:o.w,type:WidthType.DXA},shading:o.fill?{fill:o.fill,type:ShadingType.CLEAR}:undefined,
  margins:o.margins??CMAR,verticalAlign:o.valign??VerticalAlign.CENTER,
  children:Array.isArray(ch)?ch:[ch]});}
function dataTable(colW,header,rows){
  const head=new TableRow({cantSplit:true,children:header.map((h,i)=>cell([txt(h,{bold:true,color:WHITE,size:18})],{w:colW[i],fill:NAVY,borderColor:NAVY}))});
  const body=rows.map((r,ri)=>{
    const isTotal=r&&r.total; const cells=isTotal?r.cells:r;
    return new TableRow({cantSplit:true,children:cells.map((c,ci)=>{
      const obj=c&&typeof c==="object"&&!Array.isArray(c); const text=obj?c.t:c;
      const fill=isTotal?LIGHT:(ri%2?WHITE:LIGHT);
      return cell([txt(String(text),{size:18,bold:isTotal||(obj&&c.bold),color:obj?c.color:undefined,
        align:ci>0?AlignmentType.RIGHT:undefined})],{w:colW[ci],fill,
        align:ci>0?AlignmentType.RIGHT:undefined});
    })});
  });
  return new Table({width:{size:CW,type:WidthType.DXA},columnWidths:colW,rows:[head,...body]});
}
function callout(title,bodyParas,{fill=AMBERBG,bar=AMBER}={}){
  return new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[CW],rows:[new TableRow({cantSplit:true,children:[
    new TableCell({width:{size:CW,type:WidthType.DXA},shading:{fill,type:ShadingType.CLEAR},
      borders:{left:{style:BorderStyle.SINGLE,size:18,color:bar},top:b("E0D6BE"),bottom:b("E0D6BE"),right:b("E0D6BE")},
      margins:{top:120,bottom:120,left:200,right:160},
      children:[txt(title,{bold:true,color:bar,size:20,spacing:{after:60}}),...bodyParas]})]})]});
}
function kpiCard(k,w){return cell([
  txt(k.value,{bold:true,color:NAVY,size:30,align:AlignmentType.CENTER,spacing:{after:20}}),
  txt(k.label,{color:GREY,size:15,align:AlignmentType.CENTER,spacing:{after:40}}),
  txt(k.status,{bold:true,color:k.color,size:15,align:AlignmentType.CENTER}),
],{w,fill:LIGHT,borderColor:"D9DEE4",margins:{top:120,bottom:120,left:80,right:80}});}
function img(path,w,h){return new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:80,after:80},
  children:[new ImageRun({type:"png",data:fs.readFileSync(path),transformation:{width:w,height:h}})]});}
function pngDims(p){const buf=fs.readFileSync(p);return {w:buf.readUInt32BE(16),h:buf.readUInt32BE(20)};}
function imgFit(path,maxW){const d=pngDims(path);const w=Math.min(maxW,d.w);return img(path,w,Math.round(w*d.h/d.w));}
function h2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,keepNext:true,children:[new TextRun({text:t})],
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:NAVY,space:2}}});}
const numbered=(t,ref="n")=>new Paragraph({numbering:{reference:ref,level:0},children:[run(t,{size:18})]});
const bullet=(t)=>new Paragraph({numbering:{reference:"bl",level:0},children:[run(t,{size:18})]});

const A=C.section_a,B=C.section_b,Cc=C.section_c,D=C.section_d,E=C.section_e,F=C.section_f,M=C.meta;
const ch=[];

// title band
ch.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:[CW],rows:[new TableRow({children:[
  new TableCell({width:{size:CW,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},borders:noB,
    margins:{top:200,bottom:160,left:240,right:240},children:[
      txt(M.company.toUpperCase(),{bold:true,color:ORANGE,size:22,spacing:{after:30}}),
      txt("Monthly Financial Report",{bold:true,color:WHITE,size:40,spacing:{after:20}}),
      txt("Reporting period: "+M.period_label,{color:"C9D4E3",size:22}),
    ]})]})]}));
ch.push(txt(`Prepared by ${M.prepared_by}   \u2022   Date ${M.date}   \u2022   For: ${M.for}`,{size:16,color:GREY,spacing:{before:120,after:20}}));
ch.push(txt("Source of truth: "+M.source_portal+".",{size:15,color:GREY,italics:true,spacing:{after:160}}));

// KPI strip
const cw4=[2340,2340,2340,2340];
ch.push(new Table({width:{size:CW,type:WidthType.DXA},columnWidths:cw4,rows:[new TableRow({children:C.kpis.map((k,i)=>kpiCard(k,cw4[i]))})]}));
ch.push(txt("",{spacing:{after:120}}));

// integrity callout
ch.push(callout("\u26A0  Data integrity \u2014 read this first",C.integrity_flags.map(f=>numbered(f,"f")),{fill:AMBERBG,bar:AMBER}));
ch.push(txt("",{spacing:{after:120}}));

// bottom line
ch.push(h2("CFO bottom line (for Hermes \u2192 CEO)"));
C.bottom_line.forEach(t=>ch.push(numbered(t,"n")));
ch.push(txt("",{spacing:{after:80}}));

// A
ch.push(h2("A.  Monthly revenue & result \u2014 "+M.period_label));
ch.push(txt([run("Revenue: ",{bold:true,size:20}),run(A.revenue_headline,{bold:true,color:NAVY,size:24}),
  run("    "+A.revenue_sub,{color:GREY,size:16})],{spacing:{after:120}}));
ch.push(txt("Revenue by plan type:",{bold:true,size:17,spacing:{after:60}}));
ch.push(dataTable([4680,2340,2340],["Plan type","Revenue","Share"],A.plan_rows));
if(A.plan_note) ch.push(txt(A.plan_note,{italics:true,size:15,color:GREY,spacing:{before:60,after:140}}));
ch.push(txt("Tracked OpEx (tooling, ads, AI only; not a full P&L):",{bold:true,size:17,spacing:{after:60}}));
ch.push(dataTable([4680,2340,2340],["Category","EUR","USD"],A.opex_rows));
ch.push(callout("Net result \u2014 net of tracked OpEx only (NOT operating profit)",[
  txt([run(A.net_eq,{size:18}),run(A.net_val,{bold:true,color:GREEN,size:22}),run("   \u2022   "+A.net_sub,{size:16,color:GREY})]),
  txt(A.net_caveat,{italics:true,size:15,color:GREY,spacing:{before:40}}),
],{fill:GREENBG,bar:GREEN}));
ch.push(txt("Revenue trend \u2014 monthly",{bold:true,size:18,spacing:{before:120,after:20}}));
ch.push(imgFit(`${DIR}/revenue_trend.png`,620));

// B
ch.push(h2("B.  Trends & comparisons"));
ch.push(callout("YoY (month) vs same month last year  \u2022  CLEAN, trust this",[
  txt([run(B.yoy_pre,{size:18}),run(B.yoy_val,{bold:true,color:GREEN,size:20})]),
],{fill:GREENBG,bar:GREEN}));
ch.push(txt("",{spacing:{after:80}}));
ch.push(txt("3-month block (T3M) vs prior T3M:",{bold:true,size:17,spacing:{after:60}}));
ch.push(dataTable([4680,2340,2340],["Window","Value","Change"],B.t3m_rows));
if(B.t3m_note) ch.push(txt(B.t3m_note,{italics:true,size:15,color:GREY,spacing:{before:60,after:140}}));
ch.push(txt(`Growth decomposition (material clients, ${B.decomp_window}):`,{bold:true,size:17,spacing:{after:60}}));
ch.push(dataTable([4680,2340,2340],["Movement","Amount","Clients"],
  B.decomp_rows.map(r=>[{t:r.label,color:r.color},{t:r.amt,color:r.color},String(r.n)])
   .concat([{total:true,cells:["Net change",B.decomp_net,String(B.decomp_material)]}])));
if(B.decomp_note) ch.push(txt(B.decomp_note,{italics:true,size:15,color:GREY,spacing:{before:60,after:120}}));

// C
ch.push(new Paragraph({children:[new PageBreak()]}));
ch.push(h2("C.  Customers \u2014 Top "+Cc.top_rows.length));
ch.push(dataTable([2700,1380,1380,900,1500,1500],
  ["Client","This mo","Prev mo*","MoM","T3M","Same mo LY"],
  Cc.top_rows.map(r=>[
    {t:r.label,color:r.new?ORANGE:undefined}, r.this,
    {t:r.prev,color:r.prev_color}, {t:r.mom,color:r.mom_color},
    r.t3m, {t:r.yoy,color:r.yoy_color}])));
if(Cc.top_prev_note) ch.push(txt(Cc.top_prev_note,{italics:true,size:14,color:GREY,spacing:{before:50,after:0}}));
if(Cc.top_note) ch.push(txt(Cc.top_note,{size:16,color:GREEN,spacing:{before:60,after:80}}));
ch.push(imgFit(`${DIR}/top10.png`,600));
ch.push(txt("Revenue concentration (top-10)",{bold:true,size:17,align:AlignmentType.CENTER,spacing:{before:40,after:0}}));
ch.push(imgFit(`${DIR}/concentration.png`,300));
ch.push(callout("\u26A0  Concentration risk",[
  txt([run(Cc.conc_pre,{bold:true,size:18}),run(Cc.conc_post,{size:16})]),
  txt(Cc.conc_reco,{size:16,spacing:{before:40}}),
],{fill:AMBERBG,bar:AMBER}));
if(Cc.hygiene) ch.push(txt(Cc.hygiene,{italics:true,size:14,color:GREY,spacing:{before:60,after:80}}));

// D
ch.push(h2("D.  Run-rate & forward"));
ch.push(txt("ARR snapshot (recurring book, so MRR \u2248 monthly revenue):",{bold:true,size:17,spacing:{after:60}}));
ch.push(dataTable([5560,1900,1900],["Basis","MRR","ARR"],D.arr_rows));
ch.push(txt(D.arr_reco,{size:16,spacing:{before:60,after:60}}));
ch.push(bullet(D.pipeline_note));
ch.push(txt("",{spacing:{after:60}}));

// E
ch.push(h2("E.  Unit economics & efficiency"));
ch.push(bullet(E.cac_note));
ch.push(bullet(E.toolspend_note));
ch.push(txt("",{spacing:{after:60}}));

// F
ch.push(h2("F.  Cash & runway"));
ch.push(callout(F.runway_title,F.runway_lines.map((l,i)=>txt(l,{size:16,spacing:i?{before:40}:undefined})),{fill:GREYBG,bar:GREY}));
ch.push(txt("",{spacing:{after:80}}));

// metrics summary
ch.push(h2("SaaS metrics summary"));
ch.push(dataTable([2860,3100,1700,1700],["Metric","Value","Target","Status"],
  C.metrics_rows.map(r=>[r[0],r[1],r[2],r[3]])));
ch.push(txt("\u2014 End of report \u2014",{align:AlignmentType.CENTER,color:GREY,size:15,spacing:{before:240}}));

const doc=new Document({
  styles:{default:{document:{run:{font:"Arial",size:20}}},
    paragraphStyles:[{id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,
      run:{size:26,bold:true,font:"Arial",color:NAVY},paragraph:{spacing:{before:240,after:120},outlineLevel:1}}]},
  numbering:{config:[
    {reference:"n",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:560,hanging:300}}}}]},
    {reference:"f",levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:520,hanging:300}}}}]},
    {reference:"bl",levels:[{level:0,format:LevelFormat.BULLET,text:"\u2022",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:520,hanging:280}}}}]},
  ]},
  sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1080,right:1440,bottom:1080,left:1440}}},
    footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
      border:{top:{style:BorderStyle.SINGLE,size:4,color:"D9DEE4",space:6}},
      children:[new TextRun({text:`${M.company} \u2014 Monthly Financial Report \u2014 ${M.period_label}   \u2022   Confidential   \u2022   Page `,size:14,color:GREY}),
        new TextRun({children:[PageNumber.CURRENT],size:14,color:GREY})]})]})},
    children:ch}]
});
Packer.toBuffer(doc).then(buf=>{fs.writeFileSync(OUT,buf);console.log("written",OUT);});
