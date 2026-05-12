---
mode: agent
description: Scaffold a single-file Python ADO Test Plan daily fetcher/builder.
---

Build a single-file Python 3 script at `${input:path:./daily_report.py}` that pulls a full
Azure DevOps Test Plan and produces an Excel workbook and an HTML dashboard.

Requirements:
- Read ADO_TOKEN from the environment. Exit with code 2 if missing.
- No third-party deps except `openpyxl`. Use only `urllib`, `json`, `html` from stdlib.
- Constants block at the top: `OUT`, `ORG`, `PROJ`, `PLAN`, `ITER`, optional `ITER_START`.
- A `req()` helper with retry/backoff on 429/5xx and default `api-version=7.1`.
- Fetch in this order: plan, all suites (paginated via `continuationToken`),
  all test points per suite, linked bugs per test case, iteration bugs via WIQL,
  test runs in the iteration, and test results per run.
- Cache every raw payload to `raw_*.json`.
- Build an Excel workbook with sheets: Summary, KPIs, Per-Suite, Per-Scenario,
  Bugs, Missing-Bug Scenarios.
- Build a standalone HTML dashboard (no external CSS/JS) with KPI tiles + tables.
- Write `last_run_summary.json` and `email_body.html`.
- Exit 0 on success, non-zero with stderr message on failure.
