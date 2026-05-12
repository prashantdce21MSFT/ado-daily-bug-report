"""Project-wide open defects report (no Test Plan scope).

Pulls all open bugs in the project via WIQL plus recent comments per bug.
"""
import os, sys, json, urllib.parse, urllib.request, urllib.error, html, time
from pathlib import Path
from datetime import datetime, timezone

OUT = Path(r"__OUT__")
ORG = "__ORG__"
PROJ = "__PROJ__"
OUT.mkdir(parents=True, exist_ok=True)

TOKEN = os.environ.get("ADO_TOKEN", "").strip()
if not TOKEN:
    print("ERROR: ADO_TOKEN env var is empty", file=sys.stderr); sys.exit(2)


def req(url, params=None, method="GET", body=None):
    params = params or {}
    params.setdefault("api-version", "7.1")
    sep = "&" if "?" in url else "?"
    full = url + sep + urllib.parse.urlencode(params)
    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    data = json.dumps(body).encode() if body else None
    if body: headers["Content-Type"] = "application/json"
    r = urllib.request.Request(full, data=data, headers=headers, method=method)
    with urllib.request.urlopen(r, timeout=90) as resp:
        return json.loads(resp.read())


def main():
    wiql = {"query": (
        f"Select [System.Id] From WorkItems "
        f"Where [System.TeamProject] = '{PROJ}' "
        f"AND [System.WorkItemType] = 'Bug' "
        f"AND [System.State] NOT IN ('Closed','Resolved','Done','Removed')"
    )}
    d = req(f"{ORG}/{PROJ}/_apis/wit/wiql", method="POST", body=wiql)
    ids = [w["id"] for w in d.get("workItems", [])]
    print(f"[open] {len(ids)} bugs")

    bugs = []
    for i in range(0, len(ids), 200):
        chunk = ids[i:i+200]
        r = req(f"{ORG}/{PROJ}/_apis/wit/workitems",
                params={"ids": ",".join(map(str, chunk)),
                        "fields": "System.Id,System.Title,System.State,Microsoft.VSTS.Common.Severity,System.AssignedTo,System.CreatedDate"})
        bugs.extend(r.get("value", []))

    try:
        from openpyxl import Workbook
    except ImportError:
        print("ERROR: pip install openpyxl", file=sys.stderr); sys.exit(3)
    wb = Workbook(); ws = wb.active; ws.title = "Open Bugs"
    ws.append(["ID", "Title", "State", "Severity", "Assigned To", "Created"])
    for b in bugs:
        f = b.get("fields", {})
        ws.append([b.get("id"), f.get("System.Title", ""), f.get("System.State", ""),
                   f.get("Microsoft.VSTS.Common.Severity", ""),
                   (f.get("System.AssignedTo") or {}).get("displayName", ""),
                   f.get("System.CreatedDate", "")])
    xlsx = OUT / "Open_Defects.xlsx"
    wb.save(xlsx)

    body = OUT / "open_defects_email_body.html"
    body.write_text(f"<p><b>{len(bugs)}</b> open defects in {PROJ}.</p>", encoding="utf-8")
    (OUT / "open_defects_summary.json").write_text(json.dumps({
        "xlsx": str(xlsx), "body": str(body), "count": len(bugs)
    }, indent=2))
    print("[done]")


if __name__ == "__main__":
    try: main()
    except Exception:
        import traceback; traceback.print_exc(); sys.exit(1)
