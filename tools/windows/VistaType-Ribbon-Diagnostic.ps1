<#
VistaType-Ribbon-Diagnostic.ps1  --  run ON the machine with the problem. Reads only.

Reports the four things that decide which ribbon a transcriber actually sees, and which of
them is stale:

  * the add-in in the STARTUP folder, and whether the ribbon EMBEDDED in it carries the
    buttons in question - if it does not, the .dotm itself did not get replaced
  * the user's own Word.officeUI, Roaming AND Local, since Word reads whichever its profile
    dictates and the other silently goes stale
  * UserRibbonTabs, which is what makes the add-in's own tabs go dark
  * the tail of qat.log, which records exactly what the last install did

Written 8/6/2026 for the fault Jerry hit: upgrading over an older version left both ribbons
without the delete-prodnote buttons, and the braille tab without change-prodnote-to-TN,
while a clean install on a machine with no previous version was perfect.

Pass -Label to say WHEN this snapshot was taken. Every run APPENDS to
vistatype-ribbon-log.txt IN THIS SCRIPT'S OWN FOLDER, so a whole sequence can be sent in one
piece and the file is where you just were. The full path is printed at the end. Three
snapshots settle it: before the upgrade, after it with Word still closed, and after Word has
been opened and closed once. Right in the second and wrong in the third means Word is
rewriting the file, not the installer.

Usage:
    powershell -ExecutionPolicy Bypass -File VistaType-Ribbon-Diagnostic.ps1 -Label "before upgrade"
#>
param([string]$Label = "")

$ErrorActionPreference = 'SilentlyContinue'

$script:out = @()
function W($s) { $script:out += $s; Write-Host $s }

$S = [Environment]::GetFolderPath('ApplicationData')      + "\Microsoft\Word\STARTUP\LPandBRL.dotm"
$R = [Environment]::GetFolderPath('ApplicationData')      + "\Microsoft\Office\Word.officeUI"
$L = [Environment]::GetFolderPath('LocalApplicationData') + "\Microsoft\Office\Word.officeUI"

$stamp = "== VistaType ribbon diagnostic =="
if ($Label) { $stamp += "  [$Label]" }
W ($stamp + "  " + (Get-Date))

if (Test-Path $S) {
    $i = Get-Item $S
    W ("  add-in : {0:N0} bytes, {1}" -f $i.Length, $i.LastWriteTime)

    # Read the ribbon straight out of the .dotm: if the buttons are missing HERE, the add-in
    # itself was not replaced and nothing about Word.officeUI matters.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tmp = Join-Path $env:TEMP 'vtdiag.dotm'
    Copy-Item $S $tmp -Force
    $x = ''
    $z = [IO.Compression.ZipFile]::OpenRead($tmp)
    $e = $z.Entries | Where-Object { $_.FullName -like '*customUI14.xml' }
    if ($e) {
        $sr = New-Object IO.StreamReader($e.Open())
        $x = $sr.ReadToEnd()
        $sr.Close()
    }
    $z.Dispose()
    Remove-Item $tmp -Force
    $d1 = ([regex]::Matches($x, 'Sh_Delete_Prodnote_Paragraphs')).Count
    $t1 = ([regex]::Matches($x, 'Dx_Change_Prodnotes_To_Transcriber_Notes')).Count
    W ("  embedded ribbon: delete-prodnote={0}  change-to-TN={1}" -f $d1, $t1)
} else {
    W "  add-in : NOT FOUND in the STARTUP folder"
}

foreach ($pair in @(@('Roaming', $R), @('Local', $L))) {
    $name = $pair[0]
    $path = $pair[1]
    if (Test-Path $path) {
        $y  = Get-Content $path -Raw
        $tb = ([regex]::Matches($y, 'vt_tab_')).Count
        $d2 = ([regex]::Matches($y, 'Sh_Delete_Prodnote_Paragraphs')).Count
        $t2 = ([regex]::Matches($y, 'Dx_Change_Prodnotes_To_Transcriber_Notes')).Count
        $wt = (Get-Item $path).LastWriteTime
        W ("  {0,-8}: tabs={1} delete-prodnote={2} change-to-TN={3}  {4}" -f $name, $tb, $d2, $t2, $wt)
    } else {
        W ("  {0,-8}: no file" -f $name)
    }
}

$reg = (Get-ItemProperty 'HKCU:\Software\VistaType LP').UserRibbonTabs
W ("  UserRibbonTabs = '{0}'" -f $reg)

$log = [Environment]::GetFolderPath('ApplicationData') + "\VistaType LP\qat.log"
if (Test-Path $log) {
    W "  -- last 8 lines of qat.log --"
    Get-Content $log -Tail 8 | ForEach-Object { W ("     " + $_) }
} else {
    W "  qat.log: none"
}

# ---- WHICH COPY OF THE ADD-IN IS WORD ACTUALLY LOADING? -------------------------------
# Added 8/7/2026. Jerry installed 3.0.95 over itself and the About box still showed an older
# version, so the add-in Word loads is not the one the installer writes. Word's startup folder
# is configurable, and a global template can be registered at any path, so there can easily be
# a second, older LPandBRL.dotm that wins.
W "  -- copies of LPandBRL.dotm found --"
$seen = @{}
$places = @(
    (Join-Path $env:APPDATA      'Microsoft\Word\STARTUP'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\Word\STARTUP'),
    (Join-Path $env:APPDATA      'Microsoft\Templates'),
    [Environment]::GetFolderPath('MyDocuments'),
    [Environment]::GetFolderPath('Desktop'),
    'C:\Program Files\Microsoft Office\root\Office16\STARTUP',
    'C:\Program Files (x86)\Microsoft Office\root\Office16\STARTUP'
)
# Word's own configured startup folder, which overrides the default one.
foreach ($v in @('16.0','15.0','14.0')) {
    $k = "HKCU:\Software\Microsoft\Office\$v\Word\Options"
    $sp = (Get-ItemProperty -Path $k -Name 'STARTUP-PATH' -ErrorAction SilentlyContinue).'STARTUP-PATH'
    if ($sp) { W ("  Word $v startup folder is set to: " + $sp); $places += $sp }
}
foreach ($dir in $places) {
    if (-not $dir -or -not (Test-Path $dir)) { continue }
    foreach ($f in (Get-ChildItem -Path $dir -Filter 'LPandBRL.dotm' -ErrorAction SilentlyContinue)) {
        if ($seen.ContainsKey($f.FullName)) { continue }
        $seen[$f.FullName] = $true
        W ("     {0:N0} bytes  {1}  {2}" -f $f.Length, $f.LastWriteTime, $f.FullName)
    }
}
if ($seen.Count -eq 0) { W "     none found in the usual places" }
if ($seen.Count -gt 1) { W "     MORE THAN ONE - Word may well be loading the wrong one" }

# Global templates registered by path, which load regardless of the startup folder.
foreach ($v in @('16.0','15.0','14.0')) {
    $k = "HKCU:\Software\Microsoft\Office\$v\Word\Options"
    $p2 = Get-ItemProperty -Path $k -ErrorAction SilentlyContinue
    if ($p2) {
        foreach ($n in ($p2.PSObject.Properties.Name | Where-Object { $_ -like 'GLOBALDOTNAME*' })) {
            W ("  Word $v $n = " + $p2.$n)
        }
    }
}

# Write the log NEXT TO THIS SCRIPT. It used to go to the Desktop, which is the wrong answer
# on any machine with OneDrive folder redirection: GetFolderPath('Desktop') returns the
# OneDrive path and the file lands somewhere the user is not looking. Beside the script is
# where they just were. Falls back to the Desktop, then TEMP.
# Errors are NOT suppressed here: the whole point is knowing where the file went, and the
# SilentlyContinue at the top of this script previously hid a failed write completely.
$ErrorActionPreference = 'Continue'
$logFile = $null
$candidates = @()
if ($PSScriptRoot) { $candidates += (Join-Path $PSScriptRoot 'vistatype-ribbon-log.txt') }
$candidates += (Join-Path ([Environment]::GetFolderPath('Desktop')) 'vistatype-ribbon-log.txt')
$candidates += (Join-Path $env:TEMP 'vistatype-ribbon-log.txt')
foreach ($c in $candidates) {
    try {
        Add-Content -Path $c -Value ($script:out -join [Environment]::NewLine) -ErrorAction Stop
        Add-Content -Path $c -Value "" -ErrorAction Stop
        $logFile = $c
        break
    } catch { }
}
Write-Host ""
if ($logFile) {
    Write-Host "----------------------------------------------------------------"
    Write-Host ("LOG FILE:  " + $logFile)
    Write-Host "----------------------------------------------------------------"
} else {
    Write-Host "Could not write a log file anywhere. Copy the text above instead."
}
