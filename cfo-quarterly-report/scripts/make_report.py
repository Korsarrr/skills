#!/usr/bin/env python3
"""
cfo-monthly-report / make_report.py  —  one-shot orchestrator.
  python3 make_report.py <input.json> <output.docx>
Pipeline: compute.py -> charts.py -> build.js -> docx validate.
Requires: python(matplotlib), node(docx) — both present in the agent env.
"""
import subprocess, sys, os, tempfile, shutil

HERE=os.path.dirname(os.path.abspath(__file__))
DOCX_VALIDATE="/mnt/skills/public/docx/scripts/office/validate.py"

def run(cmd, **kw):
    print("›", " ".join(cmd))
    subprocess.run(cmd, check=True, **kw)

def main():
    inp = sys.argv[1] if len(sys.argv)>1 else "input.json"
    out = sys.argv[2] if len(sys.argv)>2 else "Monthly_Financial_Report.docx"
    work = tempfile.mkdtemp(prefix="cfo_m_")
    charts = os.path.join(work,"assets"); os.makedirs(charts, exist_ok=True)
    computed = os.path.join(work,"computed.json")
    env = dict(os.environ, NODE_PATH=subprocess.check_output(["npm","root","-g"]).decode().strip())
    run(["python3", f"{HERE}/compute.py", inp, computed])
    run(["python3", f"{HERE}/charts.py", computed, charts])
    run(["node", f"{HERE}/build.js", computed, charts, out], env=env)
    if os.path.exists(DOCX_VALIDATE):
        run(["python3", DOCX_VALIDATE, out])
    shutil.rmtree(work, ignore_errors=True)
    print("\n✓ Report ready:", out)

if __name__=="__main__":
    main()
