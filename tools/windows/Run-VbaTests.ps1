<#
Run-VbaTests.ps1  --  Run ON the Windows build box (has Word installed).

Runs the VBA unit tests in tests/vba against a throwaway document, invisibly, and writes the
results to a file the Linux side reads back. Nothing is installed and nothing is left behind.

The tests do NOT reference the built add-in. tools/lib/build_vba_test_bundle.py has already
lifted the procedures under test out of src/vba into VtFunctionsUnderTest.bas, so what runs is
the source in front of you rather than the last build -- and `Private` helpers, which a
referencing project could not call at all, are reachable.

PREREQUISITE: the same "Trust access to the VBA project object model" the build needs.

Usage (invoked over SSH by `make vba-test`):
    powershell -ExecutionPolicy Bypass -File Run-VbaTests.ps1 -TestRoot "C:/build/vistatype/vbatests"
#>
param(
    [Parameter(Mandatory=$true)][string]$TestRoot
)
$ErrorActionPreference = "Stop"

$TestRoot = [IO.Path]::GetFullPath($TestRoot.Replace('/', '\'))
if (-not (Test-Path $TestRoot)) { throw "TestRoot not found: $TestRoot" }

$ResultFile = Join-Path $TestRoot "results.txt"
Remove-Item -LiteralPath $ResultFile -Force -ErrorAction SilentlyContinue

# --- Word must not already be running -------------------------------------------------
#
# READ THIS OUTPUT BEFORE KILLING ANYTHING. A WINWORD.EXE in SessionId 0 with
# "/Automation -Embedding" is one of these runs, left behind by a timeout. Anything in another
# session is somebody's own Word with unsaved work in it. This script never kills either; it
# stops and says which it found. (8/21/2026: a kill-first "pre-install check" was run while
# Jerry had Word open.)
$running = @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
    Write-Host "Word is already running - not starting another, and not killing this one:"
    $running | ForEach-Object {
        Write-Host ("  pid {0}  session {1}  {2}" -f $_.ProcessId, $_.SessionId, $_.CommandLine)
    }
    Write-Host ""
    Write-Host "A session 0 process with /Automation -Embedding is a leftover test run and is"
    Write-Host "safe to end. Any other session is a real Word with real work in it."
    exit 1
}

# A Word that was killed leaves a hidden ~$ owner file behind, and the next run then meets the
# modal "File In Use" dialog, which nobody can answer headlessly. See Import-Vba.ps1.
Get-ChildItem -Path $TestRoot -Force -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name.StartsWith('~$') } |
    ForEach-Object {
        Write-Host "clearing stale Word lock file $($_.Name)"
        Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
    }

# --- the modules to import, in order ---------------------------------------------------
$libDir = Join-Path $TestRoot "lib"
$modules = @(
    (Join-Path $libDir "TestSuite.cls"),
    (Join-Path $libDir "TestCase.cls"),
    (Join-Path $libDir "VtFileReporter.cls"),
    (Join-Path $TestRoot "VtFunctionsUnderTest.bas")
)
$modules += @(Get-ChildItem -Path $TestRoot -Filter "Test*.bas" -File | Sort-Object Name |
              ForEach-Object { $_.FullName })

foreach ($m in $modules) {
    if (-not (Test-Path $m)) {
        throw "missing module: $m  (run it with make vba-test, which builds the bundle and copies tests/vba)"
    }
}

# VtRunTests.bas carries a placeholder for the result path. Environ$("APPDATA") inside an
# automation Word points at the SYSTEM profile, not the user's, so every path is built here and
# substituted into the VBA text. Measured 9/1/2026.
$runnerSrc = Join-Path $TestRoot "VtRunTests.bas"
if (-not (Test-Path $runnerSrc)) { throw "missing module: $runnerSrc" }
$runnerTmp = Join-Path $env:TEMP ("VtRunTests-{0}.bas" -f [Guid]::NewGuid().ToString("N"))
$Ansi = [System.Text.Encoding]::GetEncoding(0)
$runnerText = [IO.File]::ReadAllText($runnerSrc, $Ansi).Replace("__VT_RESULT_PATH__", $ResultFile)
[IO.File]::WriteAllText($runnerTmp, $runnerText, $Ansi)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.WordBasic.DisableAutoMacros(1)

$doc = $null
try {
    # Visible:=False. A visible document becomes the active one, and $word.Run then looks for
    # the macro THERE rather than in the document carrying the modules (9/1/2026).
    $doc = $word.Documents.Add([Type]::Missing, [Type]::Missing, [Type]::Missing, $false)
    $proj = $doc.VBProject
    if ($proj -eq $null) {
        throw "No VBProject. Enable 'Trust access to the VBA project object model'."
    }

    # No reference is added. vba-test's TestCase wanted an early-bound Scripting.Dictionary;
    # adding that reference here SUCCEEDS and the type still does not resolve, which is a
    # compile error, which is a modal dialog behind an invisible Word, which is a hang. The
    # vendored copy late-binds instead. See tests/vba/README.md.

    foreach ($m in $modules) { $proj.VBComponents.Import($m) | Out-Null }
    $proj.VBComponents.Import($runnerTmp) | Out-Null

    $doc.Activate()
    $word.Run("VtRunAllTests") | Out-Null

    $doc.Close($false)
    $doc = $null
} finally {
    if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Remove-Item -LiteralPath $runnerTmp -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path $ResultFile)) {
    Write-Host "NO RESULTS. The run produced no file at $ResultFile."
    Write-Host "That usually means a module did not compile, so VtRunAllTests was never entered."
    exit 1
}

$lines = Get-Content -LiteralPath $ResultFile
$lines | ForEach-Object { Write-Host $_ }

$verdict = $lines | Where-Object { $_ -like "RESULT *" } | Select-Object -Last 1
if ($verdict -eq "RESULT OK") { exit 0 }
exit 1
