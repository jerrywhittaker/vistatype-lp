# Install-Watch42.ps1 - installs (or with -Remove, removes) the issue #42 watcher on VistaBuild.
# Run over SSH from the repo:  scp tools/windows/watch42.ps1 tools/windows/Install-Watch42.ps1 vistabuild:VistaWatch/
#   then: ssh vistabuild "powershell -NoProfile -ExecutionPolicy Bypass -File VistaWatch\Install-Watch42.ps1"
# The task runs as jerry, only when he is logged on (his session, so window titles are readable),
# hidden (conhost --headless: no console window flashes), at low priority, and starts again at each logon.
param([switch]$Remove)
$dir  = "$env:USERPROFILE\VistaWatch"
$name = 'VistaType Watch42'
if ($Remove) {
    Stop-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction SilentlyContinue
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" |
        Where-Object { $_.CommandLine -cmatch '\\watch42\.ps1' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
    "Removed the task. Output left in $dir - delete the folder by hand if it is no longer wanted."
    return
}
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$user = "$env:COMPUTERNAME\$env:USERNAME"
$act  = New-ScheduledTaskAction -Execute "$env:WINDIR\System32\conhost.exe" `
        -Argument "--headless powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$dir\watch42.ps1`""
$trig = New-ScheduledTaskTrigger -AtLogOn -User $user
$set  = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -Priority 7 -MultipleInstances IgnoreNew `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$prin = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName $name -Action $act -Trigger $trig -Settings $set -Principal $prin -Force | Out-Null
# Start now if Jerry's desktop session is up (explorer.exe in a session other than 0).
if (Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" | Where-Object { $_.SessionId -ne 0 }) {
    Start-ScheduledTask -TaskName $name; Start-Sleep 5
}
Get-ScheduledTask -TaskName $name | Select-Object TaskName, State | Format-Table -AutoSize
Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.CommandLine -cmatch '\\watch42\.ps1' } |
    ForEach-Object { "watcher running: pid $($_.ProcessId), session $($_.SessionId)" }
if (Test-Path "$dir\watch42.log") { Get-Content "$dir\watch42.log" -Tail 5 }
