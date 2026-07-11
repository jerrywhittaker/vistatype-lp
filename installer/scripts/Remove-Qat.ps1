<#
Remove-Qat.ps1  --  run by the uninstaller, per-user.

Removes ONLY VistaType's quick-access-toolbar items -- our buttons and our group
separator, all marked with the idQ prefix "x1:VT_" -- from the user's Word.officeUI,
leaving their ribbon and their own QAT items intact.
#>
$ErrorActionPreference = "Stop"

$MSO = "http://schemas.microsoft.com/office/2009/07/customui"
$Target = Join-Path $env:APPDATA "Microsoft\Office\Word.officeUI"
if (-not (Test-Path -LiteralPath $Target)) { exit 0 }

[xml]$doc = Get-Content -LiteralPath $Target -Raw
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("mso", $MSO)

$shared = $doc.SelectSingleNode("//mso:qat/mso:sharedControls", $ns)
if ($shared) {
    $removed = 0
    foreach ($c in @($shared.ChildNodes)) {
        if ($c.GetAttribute("idQ") -like "x1:VT_*") { [void]$shared.RemoveChild($c); $removed++ }
    }
    if ($removed -gt 0) {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $writer = [System.Xml.XmlWriter]::Create($Target, $settings)
        try { $doc.Save($writer) } finally { $writer.Close() }
    }
    Write-Host "Removed $removed VistaType QAT buttons from $Target"
}
