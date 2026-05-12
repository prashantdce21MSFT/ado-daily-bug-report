# ADO Daily Bug Report — Local Setup

This folder was scaffolded by the **ADO Daily Bug Report** VS Code extension.

## Quick start

1. One-time `az login`:
   ```powershell
   az login --allow-no-subscriptions
   az account show
   ```
2. Manual test:
   ```powershell
   $env:ADO_TOKEN = & az account get-access-token --resource '499b84ac-1321-427f-aa17-267ca6975798' --query accessToken -o tsv
   python "__WORK__\daily_report.py"
   ```
3. Test the email pipeline:
   ```powershell
   & "__WORK__\run_daily.ps1"
   ```
4. Register the daily task (Command Palette → *ADO Daily Report: Show Register-ScheduledTask Command*) or:
   ```powershell
   $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "__WORK__\run_daily.ps1"'
   $trigger = New-ScheduledTaskTrigger -Daily -At '7:00AM'
   $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)
   Register-ScheduledTask -TaskName 'ADO Daily Bug Report' -Action $action -Trigger $trigger -Settings $settings
   ```

## Constants currently configured

- **ORG**: `__ORG__`
- **PROJ**: `__PROJ__`
- **PLAN**: `__PLAN__`
- **ITER**: `__ITER__`
- **ITER_START**: `__ITER_START__`
- **Recipients**: `__RECIPIENTS__`
- **Alert recipient**: `__ALERT__`

## Iterating with Copilot

Run *ADO Daily Report: Install Copilot Chat Prompts* in your workspace to add
slash-command prompts under `.github/prompts/`.
