---
mode: agent
description: Generate the Register-ScheduledTask PowerShell for a daily run.
---

Give me the PowerShell to register a daily scheduled task that runs
`${input:script:run_daily.ps1}` at `${input:time:7:00AM}` local time, hidden
window, ExecutionPolicy Bypass, with
`-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries` and a 1-hour
`ExecutionTimeLimit`. Show me how to add a second daily trigger and how to
verify `NextRunTime` with `Get-ScheduledTaskInfo`.
