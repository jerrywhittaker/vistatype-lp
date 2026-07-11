<#
Merge-Qat.ps1  --  run by the installer, per-user, on the target machine.

Non-destructively adds VistaType's quick-access-toolbar icons to the user's own
Word.officeUI. Their ribbon and their existing QAT items are left untouched; only
our buttons (marked with an "x1:VT_" idQ prefix) are added. Idempotent: re-running
replaces our buttons rather than duplicating them.

The ribbon TABS come from the embedded customUI in LPandBRL.dotm; this script only
touches the QAT, which a template add-in cannot populate on its own.

Usage:  powershell -ExecutionPolicy Bypass -File Merge-Qat.ps1 -Fragment qat-controls.xml
#>
param(
    [Parameter(Mandatory=$true)][string]$Fragment
)
$ErrorActionPreference = "Stop"

$MSO = "http://schemas.microsoft.com/office/2009/07/customui"
$X1  = "http://schemas.microsoft.com/office/2009/07/customui/macro"
$Target = Join-Path $env:APPDATA "Microsoft\Office\Word.officeUI"

[xml]$frag = Get-Content -LiteralPath $Fragment -Raw
$buttons = @($frag.qatControls.button)

if (Test-Path -LiteralPath $Target) {
    [xml]$doc = Get-Content -LiteralPath $Target -Raw
} else {
    New-Item -ItemType Directory -Force -Path (Split-Path $Target) | Out-Null
    [xml]$doc = "<mso:customUI xmlns:x1=`"$X1`" xmlns:mso=`"$MSO`">" +
                "<mso:ribbon><mso:qat><mso:sharedControls></mso:sharedControls>" +
                "</mso:qat></mso:ribbon></mso:customUI>"
}
$root = $doc.DocumentElement

# idQ="x1:VT_..." needs the x1 macro namespace declared on the root.
if ([string]::IsNullOrEmpty($root.GetAttribute("xmlns:x1"))) {
    $root.SetAttribute("xmlns:x1", $X1)
}

$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("mso", $MSO)

function Get-OrCreate($parent, $name) {
    $node = $parent.SelectSingleNode("mso:$name", $ns)
    if (-not $node) {
        $node = $doc.CreateElement("mso", $name, $MSO)
        [void]$parent.AppendChild($node)
    }
    return $node
}
$ribbon = Get-OrCreate $root   "ribbon"
$qat    = Get-OrCreate $ribbon "qat"
$shared = Get-OrCreate $qat    "sharedControls"

# Remove any previously-injected VistaType buttons (idempotent reinstall/upgrade).
foreach ($c in @($shared.ChildNodes)) {
    if ($c.GetAttribute("idQ") -like "x1:VT_*") { [void]$shared.RemoveChild($c) }
}

# Append our buttons.
foreach ($b in $buttons) {
    $btn = $doc.CreateElement("mso", "button", $MSO)
    $btn.SetAttribute("idQ", "x1:VT_" + $b.macro)
    $btn.SetAttribute("label", $b.label)
    if ($b.imageMso) { $btn.SetAttribute("imageMso", $b.imageMso) }
    $btn.SetAttribute("onAction", $b.macro)
    $btn.SetAttribute("visible", "true")
    [void]$shared.AppendChild($btn)
}

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)  # UTF-8, no BOM
$writer = [System.Xml.XmlWriter]::Create($Target, $settings)
try { $doc.Save($writer) } finally { $writer.Close() }
Write-Host "Merged $($buttons.Count) VistaType QAT buttons into $Target"
