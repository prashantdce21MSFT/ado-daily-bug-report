# Daily ADO bug report driver. Runs via Windows Task Scheduler.
$ErrorActionPreference = 'Stop'
$Work = '__WORK__'
$LogDir = Join-Path $Work 'logs'
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$Stamp = Get-Date -Format 'yyyy-MM-dd'
$Log = Join-Path $LogDir "run_$Stamp.log"
Start-Transcript -Path $Log -Append | Out-Null

$Recipient = '__RECIPIENTS__'
$AlertRecipient = '__ALERT__'

function Send-OutlookMail {
    param([Parameter(Mandatory)][string]$Subject,
          [Parameter(Mandatory)][string]$HtmlBody,
          [Parameter(Mandatory)][string]$To,
          [string[]]$Attachments = @())
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.Subject = $Subject
    $mail.HTMLBody = $HtmlBody
    $mail.To = $To
    foreach ($a in $Attachments) { if (Test-Path $a) { $null = $mail.Attachments.Add($a) } }
    $mail.Send()
}

try {
    Write-Host "=== Daily ADO Report Run: $(Get-Date) ==="
    Write-Host 'Refreshing ADO token from cached az login...'
    $token = & az account get-access-token --resource '499b84ac-1321-427f-aa17-267ca6975798' --query accessToken -o tsv 2>&1
    if ($LASTEXITCODE -ne 0 -or -not $token) {
        $errMsg = "az login cache invalid or expired. $token"
        Send-OutlookMail -To $AlertRecipient `
            -Subject "[Action required] ADO daily report FAILED - auth expired ($Stamp)" `
            -HtmlBody "<p>Run <code>az login --allow-no-subscriptions</code> then the next run will succeed.</p><pre>$errMsg</pre>"
        throw $errMsg
    }
    $env:ADO_TOKEN = $token

    Write-Host 'Running daily_report.py...'
    & python (Join-Path $Work 'daily_report.py')
    if ($LASTEXITCODE -ne 0) {
        $errMsg = "Python script exited with code $LASTEXITCODE."
        Send-OutlookMail -To $AlertRecipient `
            -Subject "[Action required] ADO daily report FAILED ($Stamp)" `
            -HtmlBody "<pre>$errMsg</pre>" -Attachments @($Log)
        throw $errMsg
    }

    $summaryPath = Join-Path $Work 'last_run_summary.json'
    $bodyPath = Join-Path $Work 'email_body.html'
    if (-not (Test-Path $summaryPath)) { throw "Summary file missing: $summaryPath" }
    if (-not (Test-Path $bodyPath)) { throw "Email body missing: $bodyPath" }
    $s = Get-Content $summaryPath -Raw | ConvertFrom-Json
    $body = Get-Content $bodyPath -Raw -Encoding UTF8

    $subject = "__SUBJECT_TAG__ $($s.kpi.bugs) bugs - $Stamp"
    Send-OutlookMail -To $Recipient -Subject $subject -HtmlBody $body -Attachments @($s.xlsx, $s.dashboard)
    Write-Host "Email sent."
    Write-Host '=== Done ==='
}
catch { Write-Error $_; throw }
finally { Stop-Transcript | Out-Null }
