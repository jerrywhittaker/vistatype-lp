<#
Remove-Qat.ps1  --  run by the uninstaller, per-user.

Restores the user's Quick Access Toolbar to what it was before VistaType was installed.
Merge-Qat backed up the pristine Word.officeUI once (as "<file>.vtqatbak"); here we restore it
and delete the backup. If there is no backup (we created the file from scratch), we instead
strip only VistaType's own items -- our ribbon-control references (idQ prefix maps to a
namespace = the LPandBRL.dotm path) plus legacy "x1:VT_"/"msox:VT_" markers -- and leave the
user's ribbon and their own QAT items intact. Cleans BOTH the Roaming and Local copies.
#>
$ErrorActionPreference = "Stop"

$MSO = "http://schemas.microsoft.com/office/2009/07/customui"
$Targets = @(
    (Join-Path $env:APPDATA      "Microsoft\Office\Word.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Word.officeUI")
)

foreach ($Target in $Targets) {
    $bak = "$Target.vtqatbak"
    if (Test-Path -LiteralPath $bak) {
        if ((Get-Item -LiteralPath $bak).Length -eq 0) {
            # Empty sentinel: there was no original -- we created the file, so remove it.
            if (Test-Path -LiteralPath $Target) { Remove-Item -LiteralPath $Target -Force }
            Write-Host "Removed VistaType-created QAT file $Target"
        } else {
            # Restore the pristine pre-install QAT.
            Copy-Item -LiteralPath $bak -Destination $Target -Force
            Write-Host "Restored pre-install QAT for $Target"
        }
        Remove-Item -LiteralPath $bak -Force
        continue
    }
    if (-not (Test-Path -LiteralPath $Target)) { continue }

    # No backup: strip only VistaType items.
    [xml]$doc = Get-Content -LiteralPath $Target -Raw
    $root = $doc.DocumentElement
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("mso", $MSO)
    $shared = $doc.SelectSingleNode("//mso:qat/mso:sharedControls", $ns)
    if (-not $shared) { continue }

    $removed = 0
    foreach ($c in @($shared.ChildNodes)) {
        $idQ = $c.GetAttribute("idQ"); if (-not $idQ) { continue }
        $isOurs = ($idQ -like "x1:VT_*" -or $idQ -like "msox:VT_*")
        if (-not $isOurs -and $idQ.Contains(":")) {
            $prefix = $idQ.Substring(0, $idQ.IndexOf(":"))
            if ($prefix -ne "mso" -and ($root.GetAttribute("xmlns:$prefix") -like "*LPandBRL.dotm")) {
                $isOurs = $true
            }
        }
        if ($isOurs) { [void]$shared.RemoveChild($c); $removed++ }
    }
    if ($removed -gt 0) {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $writer = [System.Xml.XmlWriter]::Create($Target, $settings)
        try { $doc.Save($writer) } finally { $writer.Close() }
    }
    Write-Host "Removed $removed VistaType QAT item(s) from $Target"
}
