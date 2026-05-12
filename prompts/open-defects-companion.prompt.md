---
mode: agent
description: Add a project-wide "open defects" companion report on a separate schedule.
---

Add a second Python script `open_defects_report.py` and a second PowerShell
driver `run_open_defects.ps1` that runs at `${input:time:4:00PM}` daily. It should:

- Fetch all OPEN bugs in the project via WIQL (no Test Plan scope).
- For each bug, fetch the latest 3 comments.
- Produce an Excel and an HTML body grouping bugs by State / Severity /
  Assigned-To.
- Send to a single triage owner only.
