"""Daily ADO Test Plan bug report — generic template.

Reads ADO_TOKEN from env. Produces an Excel workbook, HTML dashboard,
HTML email body, and last_run_summary.json.

Edit the constants block below for your project, then run via run_daily.ps1.
"""
import os, sys, json, urllib.parse, urllib.request, urllib.error, html, time
from pathlib import Path
from datetime import datetime, timezone

# ============ EDIT THESE ============
OUT = Path(r"__OUT__")
ORG = "__ORG__"
PROJ = "__PROJ__"
PLAN = int("__PLAN__")
ITER = r"__ITER__"
ITER_START = "__ITER_START__"   # "" to disable date floor
# ====================================

OUT.mkdir(parents=True, exist_ok=True)

TOKEN = os.environ.get("ADO_TOKEN", "").strip()
if not TOKEN:
    print("ERROR: ADO_TOKEN env var is empty", file=sys.stderr)
    sys.exit(2)


def req(url, params=None, method="GET", body=None, max_tries=4):
    params = params or {}
    params.setdefault("api-version", "7.1")
    sep = "&" if "?" in url else "?"
    full = url + sep + urllib.parse.urlencode(params)
    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    data = None
    if body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    last = None
    for attempt in range(1, max_tries + 1):
        try:
            r = urllib.request.Request(full, data=data, headers=headers, method=method)
            with urllib.request.urlopen(r, timeout=90) as resp:
                return json.loads(resp.read()), dict(resp.getheaders())
        except urllib.error.HTTPError as e:
            last = e
            if e.code in (429, 500, 502, 503, 504) and attempt < max_tries:
                time.sleep(min(2 ** attempt, 15)); continue
            raise
        except (urllib.error.URLError, TimeoutError) as e:
            last = e
            if attempt < max_tries:
                time.sleep(min(2 ** attempt, 15)); continue
            raise
    raise last


def fetch_plan():
    data, _ = req(f"{ORG}/{PROJ}/_apis/testplan/plans/{PLAN}")
    return data


def fetch_iteration_bugs():
    where = [f"[System.TeamProject] = '{PROJ}'", "[System.WorkItemType] = 'Bug'",
             f"[System.IterationPath] UNDER '{ITER}'"]
    if ITER_START:
        where.append(f"[System.CreatedDate] >= '{ITER_START}'")
    wiql = {"query": "Select [System.Id] From WorkItems Where " + " AND ".join(where)}
    data, _ = req(f"{ORG}/{PROJ}/_apis/wit/wiql", method="POST", body=wiql)
    ids = [w["id"] for w in data.get("workItems", [])]
    if not ids:
        return []
    out = []
    for i in range(0, len(ids), 200):
        chunk = ids[i:i+200]
        d, _ = req(f"{ORG}/{PROJ}/_apis/wit/workitems",
                   params={"ids": ",".join(map(str, chunk)),
                           "fields": "System.Id,System.Title,System.State,Microsoft.VSTS.Common.Severity,System.AssignedTo,System.CreatedDate"})
        out.extend(d.get("value", []))
    return out


def main():
    print(f"[plan] fetching {PLAN}...")
    plan = fetch_plan()
    (OUT / "raw_plan.json").write_text(json.dumps(plan, indent=2))

    print("[bugs] fetching via WIQL...")
    bugs = fetch_iteration_bugs()
    (OUT / "raw_iter_bugs.json").write_text(json.dumps(bugs, indent=2))
    print(f"  {len(bugs)} bugs")

    # ---- Excel ----
    try:
        from openpyxl import Workbook
    except ImportError:
        print("ERROR: openpyxl missing. Run: pip install openpyxl", file=sys.stderr)
        sys.exit(3)
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Plan", plan.get("name", "")])
    ws.append(["Iteration", ITER])
    ws.append(["Bugs", len(bugs)])
    ws.append(["Generated", datetime.now(timezone.utc).isoformat()])

    bs = wb.create_sheet("Bugs")
    bs.append(["ID", "Title", "State", "Severity", "Assigned To", "Created"])
    for b in bugs:
        f = b.get("fields", {})
        bs.append([
            b.get("id"),
            f.get("System.Title", ""),
            f.get("System.State", ""),
            f.get("Microsoft.VSTS.Common.Severity", ""),
            (f.get("System.AssignedTo") or {}).get("displayName", ""),
            f.get("System.CreatedDate", ""),
        ])
    xlsx = OUT / f"Test_Plan_{PLAN}_BugReport.xlsx"
    wb.save(xlsx)

    # ---- Dashboard HTML ----
    dash = OUT / f"Test_Plan_{PLAN}_Dashboard.html"
    rows = "".join(
        f"<tr><td>{b.get('id')}</td><td>{html.escape(b.get('fields', {}).get('System.Title', ''))}</td>"
        f"<td>{html.escape(b.get('fields', {}).get('System.State', ''))}</td></tr>"
        for b in bugs
    )
    dash.write_text(
        f"<html><body><h1>{html.escape(plan.get('name', ''))}</h1>"
        f"<p>{len(bugs)} bugs</p><table border=1>{rows}</table></body></html>",
        encoding="utf-8",
    )

    # ---- Email body ----
    (OUT / "email_body.html").write_text(
        f"<p>Daily ADO bug report for <b>{html.escape(plan.get('name', ''))}</b>.</p>"
        f"<p>Total bugs in scope: <b>{len(bugs)}</b>.</p>"
        f"<p>See attached Excel and Dashboard.</p>",
        encoding="utf-8",
    )

    # ---- Summary JSON ----
    states = {}
    for b in bugs:
        s = b.get("fields", {}).get("System.State", "")
        states[s] = states.get(s, 0) + 1
    (OUT / "last_run_summary.json").write_text(json.dumps({
        "xlsx": str(xlsx),
        "dashboard": str(dash),
        "kpi": {"bugs": len(bugs), "states": states},
        "scenarios": [],
    }, indent=2))

    print("[done]")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback; traceback.print_exc()
        sys.exit(1)
