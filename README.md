# ADO Daily Bug Report

Scaffold a Windows-scheduled Azure DevOps daily bug report (Python + PowerShell + Outlook COM)
and install Copilot Chat prompt templates so you can iterate with GitHub Copilot.

## Commands (Ctrl+Shift+P)

- **ADO Daily Report: Scaffold New Report Project** — pick a folder, answer prompts for ORG / PROJECT / PLAN / ITERATION / recipients; generates `daily_report.py`, `run_daily.ps1`, optional open-defects pair, `logs/`, and a `SETUP.md`.
- **ADO Daily Report: Install Copilot Chat Prompts** — drops `.prompt.md` files into `.github/prompts/` of the current workspace. Open Copilot Chat and type `/` to invoke them.
- **ADO Daily Report: Show Register-ScheduledTask Command** — builds the exact PowerShell command for a daily trigger and copies it to your clipboard.
- **ADO Daily Report: Run Now (manual test)** — opens a terminal, mints the ADO token via `az`, and runs `python daily_report.py`.

## Requirements

- Windows 10/11, Python 3.10+, `pip install openpyxl`, Azure CLI, classic Outlook desktop, PowerShell 5.1.
- One-time: `az login --allow-no-subscriptions`, pick the account with ADO access.

## Prompts installed

- `scaffold-fetcher.prompt.md`
- `scaffold-driver.prompt.md`
- `register-task.prompt.md`
- `adapt-new-project.prompt.md`
- `open-defects-companion.prompt.md`
- `debug-common.prompt.md`

See generated `SETUP.md` in your scaffolded folder for the full operations guide.
