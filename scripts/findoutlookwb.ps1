# PowerShell Script: Find WebView2 processes hosted by Outlook that include localhost
$keyword = "localhost"
$outlookPID = (Get-Process -Name "OUTLOOK").Id

Write-Host "`n🔍 Searching for WebView2 child processes of OUTLOOK.exe (PID: $outlookPID)..."

Get-CimInstance Win32_Process `
| Where-Object {
    $_.Name -eq "msedgewebview2.exe" -and $_.ParentProcessId -eq $outlookPID
} `
| ForEach-Object {
    $cmd = $_.CommandLine
    if ($cmd -match $keyword) {
        Write-Host "`n✅ Match found:"
        Write-Host "PID     : $($_.ProcessId)"
        Write-Host "Command : $cmd"
    }
}
