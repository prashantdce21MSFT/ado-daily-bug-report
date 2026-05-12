---
mode: ask
description: Common operational debugging prompts for this daily-report pipeline.
---

Help me diagnose this issue with my ADO daily report pipeline:

`${input:symptom:e.g. "Scheduled task shows LastTaskResult 267011" / "HTTP 203 + HTML sign-in page in log" / "WIQL returned exactly 1000 rows" / "Outlook COM throws 0x80080005" / "Excel says 0 bugs but ADO has bugs"}`

Walk through likely causes for THIS specific symptom and the minimal fix
(commands or code edits). Reference the well-known root causes:
- Battery / power policy on Task Scheduler.
- Wrong Azure CLI tenant active.
- WIQL 1000-row cap (paginate by ID range or date filter).
- New Outlook vs classic Outlook COM availability.
- PowerShell 7 vs 5.1 quirks.
- Iteration-path mismatch (UNDER vs leaf).
