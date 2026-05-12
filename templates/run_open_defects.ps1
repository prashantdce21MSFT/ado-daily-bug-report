# Project-wide open defects driver. Independent scheduled task.
$ErrorActionPreference = 'Stop'
$Work = '__WORK__'
$LogDir = Join-Path $Work 'logs'
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$Stamp = Get-Date -Format 'yyyy-MM-dd'
$Log = Join-Path $LogDir "open_defects_$Stamp.log"
Start-Transcript -Path $Log -Append | Out-Null

$Recipient = '__ALERT__'   # default: maintainer only

try {
    $token = & az account get-access-token --resource '499b84ac-1321-427f-aa17-267ca6975798' --query accessToken -o tsv 2>&1
    if ($LASTEXITCODE -ne 0 -or -not $token) { throw "az login cache expired" }
    $env:ADO_TOKEN = $token

    & python (Join-Path $Work 'open_defects_report.py')
    if ($LASTEXITCODE -ne 0) { throw "Python failed: $LASTEXITCODE" }

    $s = Get-Content (Join-Path $Work 'open_defects_summary.json') -Raw | ConvertFrom-Json
    $body = Get-Content $s.body -Raw -Encoding UTF8
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.Subject = "__SUBJECT_TAG__ Open Defects - $($s.count) - $Stamp"
    $mail.HTMLBody = $body
    $mail.To = $Recipient
    if (Test-Path $s.xlsx) { $null = $mail.Attachments.Add($s.xlsx) }
    $mail.Send()
}
catch { Write-Error $_; throw }
finally { Stop-Transcript | Out-Null }
