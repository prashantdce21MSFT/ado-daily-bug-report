---
mode: agent
description: Scaffold the PowerShell driver for the daily ADO report.
---

Create `${input:path:./run_daily.ps1}` that:

1. Starts a transcript at `logs\run_<yyyy-MM-dd>.log`.
2. Mints an ADO access token via
   `az account get-access-token --resource 499b84ac-1321-427f-aa17-267ca6975798`
   and exports `$env:ADO_TOKEN`. If empty/failed, email a single
   "auth expired" alert to the maintainer and `throw`.
3. Runs `python daily_report.py`. If exit code is non-zero, email the
   maintainer with the log attached and `throw`.
4. Reads `last_run_summary.json` and `email_body.html`.
5. Sends one email via Outlook COM
   (`New-Object -ComObject Outlook.Application`, `CreateItem(0)`) to the
   recipient list, subject line including KPIs, attaching the `.xlsx` and
   dashboard `.html`.
6. Stops the transcript in a `finally` block.

Use `$ErrorActionPreference = 'Stop'`. Parameterize `$Recipient` and
`$AlertRecipient` at the top.
