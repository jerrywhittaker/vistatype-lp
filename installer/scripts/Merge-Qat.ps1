<#
Merge-Qat.ps1  --  run by the installer, per-user, on the target machine.

Installs VistaType's standard Quick Access Toolbar (from qat-template.officeUI) while being
non-destructive to the user:
  * Only the QAT sharedControls are replaced; the user's ribbon customizations (tabs, etc.)
    in Word.officeUI are left untouched.
  * Any QAT icons the user added themselves are preserved and re-appended to the RIGHT of the
    VistaType block (deduped against it).
  * The user's original Word.officeUI is backed up (once) so uninstall can restore it.

Two things that make it actually work (learned the hard way):
  1. LOCATION. Word reads Word.officeUI from %APPDATA% (Roaming) on most machines but from
     %LOCALAPPDATA% (Local) when the profile roams/redirects or Office can't roam. We don't
     know which a machine uses, so we write BOTH.
  2. FORMAT. VistaType QAT buttons are REFERENCES to the add-in's own ribbon controls
     (<mso:control idQ="x1:btn_<macro>"> with x1 = the installed .dotm's path) -- the exact
     shape Word writes when a user adds one of our ribbon buttons to the QAT by hand.

Usage:  powershell -ExecutionPolicy Bypass -File Merge-Qat.ps1 -Template qat-template.officeUI
#>
param(
    [Parameter(Mandatory=$true)][string]$Template
)
$ErrorActionPreference = "Stop"

$MSO      = "http://schemas.microsoft.com/office/2009/07/customui"
$MSOX     = "http://schemas.microsoft.com/office/2006/01/customui/special"
$XMLNS    = "http://www.w3.org/2000/xmlns/"
$DotmPath = Join-Path $env:APPDATA "Microsoft\Word\STARTUP\LPandBRL.dotm"

$Targets = @(
    (Join-Path $env:APPDATA      "Microsoft\Office\Word.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Word.officeUI")
)

function Split-IdQ($idQ) {
    $i = $idQ.IndexOf(":")
    if ($i -lt 0) { return @($null, $idQ) }
    return @($idQ.Substring(0,$i), $idQ.Substring($i+1))
}
# The URI an idQ prefix maps to (mso/msox are well-known; others come from xmlns decls).
function Resolve-Uri($doc, $prefix) {
    if ($prefix -eq "mso")  { return $MSO }
    if ($prefix -eq "msox") { return $MSOX }
    if (-not $prefix) { return $null }
    $n = $doc.DocumentElement.GetAttributeNode($prefix, $XMLNS)
    if ($n) { return $n.Value }
    return $null
}
# Reuse an existing prefix mapping to $uri; else declare $preferred if free; else vt1, vt2 ...
function Ensure-Prefix($rootEl, $doc, $uri, $preferred) {
    foreach ($a in $rootEl.Attributes) {
        if ($a.Prefix -eq "xmlns" -and $a.Value -eq $uri) { return $a.LocalName }
    }
    $try = $preferred; $n = 1
    while ($rootEl.GetAttributeNode($try, $XMLNS)) { $try = "vt$n"; $n++ }
    $decl = $doc.CreateAttribute("xmlns", $try, $XMLNS); $decl.Value = $uri
    [void]$rootEl.Attributes.Append($decl)
    return $try
}
function Get-OrCreate($doc, $ns, $parent, $name) {
    $n = $parent.SelectSingleNode("mso:$name", $ns)
    if (-not $n) { $n = $doc.CreateElement("mso", $name, $MSO); [void]$parent.AppendChild($n) }
    return $n
}

# --- Load the template and pin its x1 namespace to this machine's add-in path. ---
$tplText = (Get-Content -LiteralPath $Template -Raw).Replace("__VT_DOTM_PATH__", $DotmPath)
[xml]$tpl = $tplText
$tplShared = $tpl.SelectSingleNode("//*[local-name()='sharedControls']")
$tplItems  = @($tplShared.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })

# (uri|local) set the template provides, so user items already covered aren't re-added.
$tplKeys = @{}
foreach ($it in $tplItems) {
    $idQ = $it.GetAttribute("idQ"); if (-not $idQ) { continue }
    $p, $l = Split-IdQ $idQ
    $tplKeys["$(Resolve-Uri $tpl $p)|$l"] = $true
}

function Merge-One([string]$Target) {
    # Back up the pristine original ONCE, for uninstall. If there is no original (we're
    # creating the file), leave an EMPTY sentinel so a re-run never captures our own output
    # and uninstall knows to delete the file rather than "restore" a VistaType QAT.
    $bak = "$Target.vtqatbak"
    $exists = Test-Path -LiteralPath $Target
    if (-not $exists) { New-Item -ItemType Directory -Force -Path (Split-Path $Target) | Out-Null }
    if (-not (Test-Path -LiteralPath $bak)) {
        if ($exists) { Copy-Item -LiteralPath $Target -Destination $bak -Force }
        else         { New-Item -ItemType File -Path $bak -Force | Out-Null }
    }
    if ($exists) {
        [xml]$doc = Get-Content -LiteralPath $Target -Raw
    } else {
        [xml]$doc = "<mso:customUI xmlns:mso=`"$MSO`"><mso:ribbon><mso:qat>" +
                    "<mso:sharedControls></mso:sharedControls></mso:qat></mso:ribbon></mso:customUI>"
    }
    $root = $doc.DocumentElement
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("mso", $MSO)

    $ribbon = Get-OrCreate $doc $ns $root   "ribbon"
    $qat    = Get-OrCreate $doc $ns $ribbon "qat"
    $shared = Get-OrCreate $doc $ns $qat    "sharedControls"

    # Prefixes for the template's custom namespaces (add-in path + separators ns).
    $vtPrefix  = Ensure-Prefix $root $doc $DotmPath "x1"
    $sepPrefix = Ensure-Prefix $root $doc $MSOX     "msox"

    # Snapshot the user's existing QAT items, then clear the block.
    $userItems = @($shared.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })
    foreach ($c in $userItems) { [void]$shared.RemoveChild($c) }

    # 1) Lay down the VistaType standard toolbar (template), rewriting its custom-ns prefixes.
    foreach ($it in $tplItems) {
        $el = $doc.CreateElement("mso", $it.LocalName, $MSO)
        foreach ($a in $it.Attributes) {
            $val = $a.Value
            if ($a.Name -eq "idQ") {
                $p, $l = Split-IdQ $val
                if     ($p -eq "x1")   { $val = "${vtPrefix}:$l" }
                elseif ($p -eq "msox") { $val = "${sepPrefix}:$l" }
            }
            $el.SetAttribute($a.Name, $val)
        }
        [void]$shared.AppendChild($el)
    }

    # 2) Re-append the user's own QAT icons to the RIGHT: skip separators and anything the
    #    template already provides (deduped by resolved namespace + local name).
    foreach ($u in $userItems) {
        if ($u.LocalName -eq "separator") { continue }
        $idQ = $u.GetAttribute("idQ"); if (-not $idQ) { continue }
        $p, $l = Split-IdQ $idQ
        if ($tplKeys.ContainsKey("$(Resolve-Uri $doc $p)|$l")) { continue }
        [void]$shared.AppendChild($u)     # keeps its own prefix, whose decl is still on root
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)  # UTF-8, no BOM
    $writer = [System.Xml.XmlWriter]::Create($Target, $settings)
    try { $doc.Save($writer) } finally { $writer.Close() }
    Write-Host "Installed VistaType QAT into $Target"
}

foreach ($t in $Targets) { Merge-One $t }
