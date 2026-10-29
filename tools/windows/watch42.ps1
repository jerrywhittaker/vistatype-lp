# watch42.ps1 - read-only watcher for issue #42 (Word crash/hang in MathType's Format Equations).
# Runs in Jerry's own session on VistaBuild. Every 30 s it logs one line per WINWORD.EXE and
# MathType*.exe process. When Word crashes or hangs, or Word vanishes while VistaType LP's timing
# log says MathType is formatting, it copies the evidence Windows and Office clear away soon after.
# READ-ONLY toward Word, MathType and the add-in: it never stops a process and never touches
# STARTUP or any document. It writes only under $OutDir.
#   -Once    one pass, then stop      -DryRun  write nothing; print what would be logged/copied
param([string]$OutDir = "$env:USERPROFILE\VistaWatch", [int]$Seconds = 30,
      [long]$CapBytes = 5MB, [int]$LookBackMin = 30, [switch]$Once, [switch]$DryRun)

$ErrorActionPreference = 'SilentlyContinue'
try { (Get-Process -Id $PID).PriorityClass = 'Idle' } catch {}
$log      = Join-Path $OutDir 'watch42.log'
$timing   = "$env:APPDATA\VistaType LP Settings\VistaType-Timing.log"
$diagDir  = "$env:LOCALAPPDATA\Temp\Diagnostics\WINWORD"
$werRoots = 'C:\ProgramData\Microsoft\Windows\WER\ReportArchive','C:\ProgramData\Microsoft\Windows\WER\ReportQueue',
            "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive","$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue"
if (-not $DryRun) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

function Write-Line([string]$s) {
    if ($DryRun) { Write-Output $s; return }
    if ((Test-Path $log) -and (Get-Item $log).Length -gt $CapBytes) { Move-Item $log "$log.1" -Force }
    Add-Content -Path $log -Value $s -Encoding UTF8
}
function Last-TimingLine {
    if (-not (Test-Path $timing)) { return '' }
    $t = Get-Content $timing -Tail 5 | Where-Object { $_.Trim() -ne '' }
    if ($t) { return @($t)[-1] } else { return '' }
}
function Capture([string]$why) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $dest  = Join-Path $OutDir "capture-$stamp"
    # Office does not update these files' modified time while Word writes to them, so take the
    # newest by creation time: each Word start makes two to four, each capped at 20 MB.
    $diag  = Get-ChildItem $diagDir -Recurse -File | Sort-Object CreationTime | Select-Object -Last 8
    $wer   = foreach ($r in $werRoots) { Get-ChildItem $r -Directory | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-$LookBackMin) -and $_.Name -match 'WINWORD|MT6|MathType' } }
    $mtu   = Get-ChildItem $env:TEMP -Filter 'mtu*.tmp' | Sort-Object CreationTime |
             ForEach-Object { '{0:yyyy-MM-dd HH:mm:ss} {1} {2}b' -f $_.CreationTime, $_.Name, $_.Length }
    $ev    = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddMinutes(-$LookBackMin); Id=1000,1001,1002} |
             Where-Object { $_.Message -match 'WINWORD|MT6|MathType' } |
             ForEach-Object { "{0:yyyy-MM-dd HH:mm:ss} id{1}`r`n{2}`r`n" -f $_.TimeCreated, $_.Id, $_.Message }
    if ($DryRun) {
        Write-Output "DRYRUN capture -> $dest  ($why)"
        $diag | ForEach-Object { Write-Output ("  diag  {0}  {1:N0}b" -f $_.FullName, $_.Length) }
        $wer  | ForEach-Object { Write-Output "  wer   $($_.FullName)" }
        Write-Output "  mtu   $(@($mtu).Count) file(s); events $(@($ev).Count); timing tail 400 lines"
        return
    }
    New-Item -ItemType Directory -Force -Path $dest, "$dest\diagnostics", "$dest\wer" | Out-Null
    Set-Content "$dest\why.txt" "$(Get-Date -Format o)  $env:COMPUTERNAME  $why"
    $diag | ForEach-Object { Copy-Item $_.FullName "$dest\diagnostics\" -Force }
    $wer  | ForEach-Object { Copy-Item $_.FullName "$dest\wer\" -Recurse -Force }
    if (Test-Path $timing) { Get-Content $timing -Tail 400 | Set-Content "$dest\timing-tail.txt" }
    $mtu | Set-Content "$dest\mtu-listing.txt"
    $ev  | Set-Content "$dest\events.txt"
    Write-Line ("{0:yyyy-MM-dd HH:mm:ss} CAPTURE {1} -> {2}" -f (Get-Date), $why, $dest)
}

$lastEvent = Get-Date
$prevWord  = @{}
Write-Line ("{0:yyyy-MM-dd HH:mm:ss} START {1} pid {2} session {3}" -f (Get-Date), $env:COMPUTERNAME, $PID, (Get-Process -Id $PID).SessionId)
while ($true) {
    $now  = Get-Date
    $cim  = Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE' OR Name LIKE 'MathType%'"
    $word = @{}
    foreach ($c in $cim) {
        $p = Get-Process -Id $c.ProcessId
        $cpu = if ($p) { [math]::Round($p.TotalProcessorTime.TotalSeconds, 1) } else { '' }
        $title = if ($p) { $p.MainWindowTitle } else { '' }
        Write-Line ("{0:yyyy-MM-dd HH:mm:ss} {1} pid {2} sess {3} start {4:yyyy-MM-dd HH:mm:ss} cpu {5} title ""{6}""" -f `
            $now, $c.Name, $c.ProcessId, $c.SessionId, $c.CreationDate, $cpu, $title)
        if ($c.Name -eq 'WINWORD.EXE') { $word[[int]$c.ProcessId] = $c.CreationDate }
    }
    # A crash or hang event for Word since the last pass.
    $new = Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=$lastEvent; Id=1000,1002} |
           Where-Object { $_.Message -match 'WINWORD\.EXE' }
    $lastEvent = $now
    if ($new) { Capture ("event id{0} at {1:HH:mm:ss}" -f @($new)[0].Id, @($new)[0].TimeCreated) }
    # Word gone while the timing log's last line says MathType is formatting.
    else {
        $gone = @($prevWord.Keys | Where-Object { -not $word.ContainsKey($_) })
        if ($gone.Count -gt 0 -and (Last-TimingLine) -match 'MathType is formatting the equations') {
            Capture ("Word pid {0} gone during Format Equations" -f ($gone -join ','))
        }
    }
    $prevWord = $word
    if ($Once) { break }
    Start-Sleep -Seconds $Seconds
}
