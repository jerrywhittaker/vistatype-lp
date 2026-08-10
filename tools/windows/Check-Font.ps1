<#
Check-Font.ps1 -- why is "VistaTypeLP Legible" grayed out on the attach dialog?

Run it ON the machine that has the problem, signed in as the transcriber who sees it:

    powershell -ExecutionPolicy Bypass -File Check-Font.ps1

It answers the four questions in order, and the first "no" is the cause:

  1. Is Windows new enough?  Per-user fonts need Windows 10 1803 (build 17134). Below that the
     installer skips the font deliberately and the dialog grays it out -- working as intended.
  2. Are the four .ttf files in the per-user Fonts folder?  If not, the install did not get
     that far.
  3. Has Windows been TOLD about them (HKCU ... CurrentVersion\Fonts)?  Files without registry
     entries are files Word will never see.
  4. Does Word list the family?  This is what Sh_Is_Font_Installed asks, and what grays the
     button out.

If 1-3 are yes and 4 is no, the answer is almost always SIGN OUT AND BACK IN: Windows loads
per-user fonts into the session at sign-in, and a font added mid-session does not always reach
programs started later.

NOTE: run this interactively, in a normal Command Prompt or PowerShell window on the machine
itself. Run over a remote/service connection it can report "Word lists ...: False" on a machine
where the font is perfectly fine, because that session never loaded the per-user fonts.
#>

# --- VistaType LP font diagnostic -------------------------------------------------
$family = "VistaTypeLP Legible"
Write-Host "=== Windows ==="
$cv = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Write-Host ("  {0}  build {1}.{2}" -f $cv.ProductName, $cv.CurrentBuild, $cv.UBR)
$need = 17134
if ([int]$cv.CurrentBuild -lt $need) {
  Write-Host ("  *** build {0} is older than {1} (Windows 10 1803). Per-user fonts are not supported;" -f $cv.CurrentBuild, $need)
  Write-Host  "      the installer SKIPS the font on purpose and the dialog grays it out. This is the cause."
}

Write-Host "=== the font files ==="
$dir = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"
foreach ($f in @("VistaTypeLPLegible-Regular.ttf","VistaTypeLPLegible-Bold.ttf","VistaTypeLPLegible-Italic.ttf","VistaTypeLPLegible-BoldItalic.ttf")) {
  $p = Join-Path $dir $f
  if (Test-Path $p) { Write-Host ("  {0,-34} {1} bytes" -f $f, (Get-Item $p).Length) }
  else { Write-Host ("  {0,-34} NOT THERE" -f $f) }
}
$machine = Join-Path $env:WINDIR "Fonts\VistaTypeLPLegible-Regular.ttf"
Write-Host ("  also in C:\Windows\Fonts?  {0}" -f (Test-Path $machine))

Write-Host "=== is Windows told about it? (this is what makes Word see it) ==="
$key = "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
$hit = $false
if (Test-Path $key) {
  (Get-ItemProperty $key).PSObject.Properties |
    Where-Object { $_.Name -like "*VistaTypeLP*" } |
    ForEach-Object { $hit = $true; Write-Host ("  HKCU  {0}  =  {1}" -f $_.Name, $_.Value) }
}
if (-not $hit) { Write-Host "  *** NOTHING registered under HKCU. The files are there but Windows has not been told." }

Write-Host "=== does Word see it? ==="
try {
  $w = New-Object -ComObject Word.Application
  $w.Visible = $false
  $found = @($w.FontNames) -contains $family
  Write-Host ("  Word lists '{0}': {1}" -f $family, $found)
  $w.Quit()
} catch { Write-Host ("  could not ask Word: " + $_.Exception.Message) }
