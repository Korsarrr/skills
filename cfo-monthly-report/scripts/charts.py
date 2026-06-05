#!/usr/bin/env python3
"""cfo-monthly-report / charts.py  —  reads computed.json, writes 3 chart PNGs to <outdir>."""
import json, sys, os
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.patches import Patch
import numpy as np

NAVY="#16365C"; ORANGE="#E8590C"; GREEN="#2E7D32"; RED="#D32F2F"; LIGHT="#E8EDF2"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":11,"axes.edgecolor":"#C7CDD4",
    "axes.linewidth":0.8,"axes.grid":True,"grid.color":"#E6E9ED","grid.linewidth":0.8,"figure.dpi":150})
usd=lambda x,_:f"${x/1000:.0f}K"

def short_month(m):  # "2026-05" -> "May'26"
    import datetime
    try: return datetime.datetime.strptime(m,"%Y-%m").strftime("%b'%y")
    except Exception: return m

def main(computed_path, outdir):
    os.makedirs(outdir, exist_ok=True)
    cd=json.load(open(computed_path,encoding="utf-8"))["chart_data"]
    series=cd["series"]; labels=[short_month(s["label"]) for s in series]; vals=[s["value"] for s in series]
    gap=set(cd["gap_months"]); rec=cd["record_month"]; foc=cd["focal_month"]
    raw=[s["label"] for s in series]

    # ----- revenue trend -----
    colors=[]
    for r in raw:
        if r in gap: colors.append(RED)
        elif r==rec: colors.append(GREEN)
        elif r==foc: colors.append(ORANGE)
        else: colors.append(NAVY)
    fig,ax=plt.subplots(figsize=(9.2,3.5))
    ax.bar(range(len(vals)),vals,color=colors,width=0.72,zorder=3)
    ax.set_xticks(range(len(vals))); ax.set_xticklabels(labels,rotation=45,ha="right",fontsize=8.5)
    ax.yaxis.set_major_formatter(FuncFormatter(usd)); ax.set_ylabel("Revenue (USD)",fontsize=9.5)
    ax.set_axisbelow(True); ax.spines[["top","right"]].set_visible(False)
    ymax=max(vals)*1.18 if vals else 1
    for r in raw:
        if r in gap:
            i=raw.index(r)
            ax.annotate("DATA GAP\n(missing batch)",xy=(i,ymax*0.03),xytext=(i,ymax*0.42),ha="center",
                fontsize=8,color=RED,fontweight="bold",arrowprops=dict(arrowstyle="->",color=RED,lw=1.2))
    ri=raw.index(rec); ax.annotate("record",xy=(ri,vals[ri]),xytext=(ri,vals[ri]+ymax*0.06),
        ha="center",fontsize=8,color=GREEN,fontweight="bold")
    ax.set_ylim(0,ymax)
    legend=[Patch(color=NAVY,label="Actual"),Patch(color=GREEN,label="Record month"),
            Patch(color=ORANGE,label="Focal (provisional)")]
    if gap: legend.append(Patch(color=RED,label="Data gap"))
    ax.legend(handles=legend,loc="upper left",fontsize=8,frameon=False,ncol=2)
    plt.tight_layout(); plt.savefig(f"{outdir}/revenue_trend.png",bbox_inches="tight"); plt.close()

    # ----- top clients (height scales with N) -----
    tc=cd["top_clients"]
    n=len(tc); fs_lbl=8 if n>12 else 9; fs_val=7.5 if n>12 else 8; h=max(3.6, 0.34*n+1.1)
    names=[(c["name"][:18]+"\u2026" if len(c["name"])>19 else c["name"])+("*" if c["new"] else "") for c in tc]
    cv=[c["value"] for c in tc]; ccols=[ORANGE if c["new"] else NAVY for c in tc]
    fig,ax=plt.subplots(figsize=(9.2,h)); y=np.arange(len(tc))[::-1]
    ax.barh(y,cv,color=ccols,height=0.66,zorder=3); ax.set_yticks(y); ax.set_yticklabels(names,fontsize=fs_lbl)
    ax.xaxis.set_major_formatter(FuncFormatter(usd)); ax.set_axisbelow(True); ax.spines[["top","right"]].set_visible(False)
    for yi,v in zip(y,cv): ax.text(v+max(cv)*0.012,yi,f"${v:,.0f}",va="center",fontsize=fs_val,color="#333")
    ax.set_xlim(0,max(cv)*1.15 if cv else 1)
    ax.legend(handles=[Patch(color=NAVY,label="Existing"),Patch(color=ORANGE,label="Net-new logo (YoY) *")],
        loc="lower right",fontsize=8.5,frameon=False)
    ax.set_title(f"Top {len(tc)} clients by {short_month(foc)} revenue",fontsize=10.5,color=NAVY,fontweight="bold",loc="left",pad=8)
    plt.tight_layout(); plt.savefig(f"{outdir}/top10.png",bbox_inches="tight"); plt.close()

    # ----- concentration donut -----
    cc=cd["concentration"]
    fig,ax=plt.subplots(figsize=(4.4,3.6))
    sizes=[cc["top1"],cc["top2_10"],cc["rest"]]
    dlabels=[f"Top-1\n{cc['top1']:.1f}%",f"Top 2\u201310\n{cc['top2_10']:.1f}%",f"Rest of book\n{cc['rest']:.1f}%"]
    w,_=ax.pie(sizes,colors=[ORANGE,NAVY,LIGHT],startangle=90,wedgeprops=dict(width=0.42,edgecolor="white",linewidth=2))
    ax.text(0,0.08,cc["headline"],ha="center",va="center",fontsize=20,fontweight="bold",color=NAVY)
    ax.text(0,-0.22,"in top 10",ha="center",va="center",fontsize=9,color="#555")
    ax.legend(w,dlabels,loc="center",bbox_to_anchor=(0.5,-0.18),ncol=3,fontsize=7.6,frameon=False,handlelength=1.1)
    ax.set_aspect("equal"); plt.tight_layout(); plt.savefig(f"{outdir}/concentration.png",bbox_inches="tight"); plt.close()
    print("charts written ->",outdir)

if __name__=="__main__":
    main(sys.argv[1] if len(sys.argv)>1 else "computed.json",
         sys.argv[2] if len(sys.argv)>2 else "assets")
