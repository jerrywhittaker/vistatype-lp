<#
Merge-Qat.ps1  --  run by the installer, per-user, on the target machine.

Sets up the user's Quick Access Toolbar according to what they chose in the installer.

  -Mode Vista    Install VistaType's standard toolbar WHOLE. Their own toolbar is saved
                 first and can be had back at any time (Restore, or the Switch Toolbar
                 button in Word).
  -Mode Mine     Leave their toolbar as it is and append VistaType's own icons on the end.
                 Nothing of theirs is hidden, moved or reordered.
  -Mode Restore  Put back the toolbar they had before VistaType was ever installed, then
                 append VistaType's icons. This is the rescue for anyone whose toolbar an
                 older VistaType install rewrote.
  -Mode None     Do not touch the toolbar at all.

Replaces the old merge behavior deliberately. Merging VistaType's curated toolbar INTO the
user's was the source of the trouble: it pushed their icons right, discarded their
separators, and hid buttons they had chosen to keep (Undo among them) because suppressing
Word's defaults is the only way to make a curated toolbar look clean. Whole-replace plus a
way back is honest; merging was not.

Two things that make this work at all (learned the hard way):
  1. LOCATION. Word reads Word.officeUI from %APPDATA% (Roaming) on most machines but from
     %LOCALAPPDATA% (Local) when the profile roams/redirects or Office can't roam. We don't
     know which a machine uses, so we write BOTH.
  2. FORMAT. VistaType toolbar buttons are REFERENCES to the add-in's own ribbon controls
     (<mso:control idQ="x1:btn_<macro>"> with x1 = the installed .dotm's path) -- the exact
     shape Word writes when a user adds one of our ribbon buttons to the toolbar by hand.

Word must not be running: it reads this file at startup and would not see the change.
(Verified 7/28/2026: Word does NOT rewrite the file on exit, and does not lock it.)

Usage:
  powershell -ExecutionPolicy Bypass -File Merge-Qat.ps1 -Mode Mine `
      -FullTemplate qat-template.officeUI -IconsTemplate qat-icons-only.officeUI
#>
param(
    [Parameter(Mandatory=$true)][ValidateSet("Vista","Mine","Restore","None")][string]$Mode,
    [string]$FullTemplate,
    [string]$IconsTemplate,
    # Ribbon tabs, chosen separately from the toolbar. "-Mode None -Tabs Install" is a
    # perfectly ordinary combination: leave my toolbar alone, but give me the tabs.
    [ValidateSet("Install","Skip","Remove")][string]$Tabs = "Skip",
    [string]$TabsTemplate,
    # Testing only. The installer never passes this; leaving it unset uses the two real
    # locations below. It exists so the modes can be exercised against sample toolbar files
    # without touching the machine's own Word setup.
    [string[]]$TargetPaths
)
$ErrorActionPreference = "Stop"

$MSO      = "http://schemas.microsoft.com/office/2009/07/customui"
$MSOX     = "http://schemas.microsoft.com/office/2006/01/customui/special"
$XMLNS    = "http://www.w3.org/2000/xmlns/"
$DotmPath = Join-Path $env:APPDATA "Microsoft\Word\STARTUP\LPandBRL.dotm"
$LogDir   = Join-Path $env:APPDATA "VistaType LP"
$LogFile  = Join-Path $LogDir "qat.log"

$Targets = if ($TargetPaths) { $TargetPaths } else { @(
    (Join-Path $env:APPDATA      "Microsoft\Office\Word.officeUI"),
    (Join-Path $env:LOCALAPPDATA "Microsoft\Office\Word.officeUI")
) }

function Write-Log($msg) {
    try {
        if (-not (Test-Path -LiteralPath $LogDir)) {
            New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
        }
        Add-Content -LiteralPath $LogFile -Value ("{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg)
    } catch { }
    Write-Host $msg
}

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
function Get-Elements($node) {
    # Always filter to elements: a comment or whitespace node has no GetAttribute, and under
    # ErrorActionPreference=Stop that aborts the whole run.
    @($node.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })
}
# Copy across the xmlns declaration an element's idQ prefix depends on, or the reference
# dangles once the element is imported into a different document.
function Import-NamespaceFor($destDoc, $srcDoc, $el) {
    $idQ = $el.GetAttribute("idQ")
    if (-not $idQ) { return }
    $p, $l = Split-IdQ $idQ
    if (-not $p -or $p -eq "mso") { return }
    $uri = Resolve-Uri $srcDoc $p
    if ($uri -and -not $destDoc.DocumentElement.GetAttributeNode($p, $XMLNS)) {
        $decl = $destDoc.CreateAttribute("xmlns", $p, $XMLNS)
        $decl.Value = $uri
        [void]$destDoc.DocumentElement.Attributes.Append($decl)
    }
}

# --- load a template and pin its x1 namespace to this machine's add-in path ---
# $node says which element holds the items: "sharedControls" for a toolbar, "tabs" for the
# ribbon tabs.
function Load-Template([string]$path, [string]$node = "sharedControls") {
    if (-not $path)                        { throw "no template path supplied" }
    if (-not (Test-Path -LiteralPath $path)) { throw "template not found: $path" }
    $text = (Get-Content -LiteralPath $path -Raw).Replace("__VT_DOTM_PATH__", $DotmPath)
    [xml]$doc = $text
    $holder = $doc.SelectSingleNode("//*[local-name()='$node']")
    if (-not $holder) { throw "no $node in $path" }
    return @{ Doc = $doc; Items = (Get-Elements $holder) }
}

# --- map every namespace a template declares onto a prefix that is free in the target ---
# Resolving by URI rather than by prefix NAME matters: the preferred prefix may already mean
# something else on this machine. Jerry's own build box has x1 bound to SWIFT, so VistaType's
# entries have to be written under a different prefix there or they would point at SWIFT.
function Build-PrefixMap($srcDoc, $destRoot, $destDoc) {
    $map = @{}
    foreach ($a in $srcDoc.DocumentElement.Attributes) {
        if ($a.Prefix -ne "xmlns") { continue }
        if ($a.Value -eq $MSO)     { continue }     # mso is always mso
        $map[$a.Value] = Ensure-Prefix $destRoot $destDoc $a.Value $a.LocalName
    }
    return $map
}

# --- copy a template element (and everything under it) into the target document ---
# Recursive, because a tab is three levels deep: tab > group > control. The toolbar's entries
# are flat, so this is a no-op for them beyond the attribute copy.
function Copy-Tree($doc, $srcDoc, $item, $prefixMap) {
    $el = $doc.CreateElement("mso", $item.LocalName, $MSO)
    foreach ($a in $item.Attributes) {
        $val = $a.Value
        if ($a.Name -eq "idQ") {
            $p, $l = Split-IdQ $val
            $uri = Resolve-Uri $srcDoc $p
            if ($uri -and $prefixMap.ContainsKey($uri)) { $val = "$($prefixMap[$uri]):$l" }
        }
        $el.SetAttribute($a.Name, $val)
    }
    foreach ($child in (Get-Elements $item)) {
        [void]$el.AppendChild((Copy-Tree $doc $srcDoc $child $prefixMap))
    }
    return $el
}

function Save-Xml($doc, $path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)   # UTF-8, no BOM: Word requires it
    $writer = [System.Xml.XmlWriter]::Create($path, $settings)
    try { $doc.Save($writer) } finally { $writer.Close() }
}

# --- record exactly what we wrote, so uninstall can remove precisely that and nothing else ---
function Write-Manifest([string]$target, $elements, $doc, $tabIds) {
    $man = "$target.vtqatmanifest"
    $lines = @("# VistaType manifest -- what this install wrote. Do not edit.",
               "mode=$Mode",
               "tabs=$Tabs",
               "written=$(Get-Date -Format o)")
    foreach ($el in $elements) {
        $idQ = $el.GetAttribute("idQ")
        if ($idQ) {
            $p, $l = Split-IdQ $idQ
            $uri = Resolve-Uri $doc $p
            $lines += "item=$($el.LocalName)|$uri|$l"
        }
    }
    # Tabs carry `id`, not `idQ`, so they get their own line type rather than being squeezed
    # into `item=` (which records a namespace URI and is read back as a qualified key).
    foreach ($t in $tabIds) { $lines += "tab=$t" }
    Set-Content -LiteralPath $man -Value $lines -Encoding UTF8
}

# --- take VistaType's tabs back out of the user's ribbon ---
# A tab is ours if it still carries any control pointing into the add-in. That fingerprint
# survives the user reordering, renaming or unticking the tab -- all of which Word may
# rewrite the id for -- because Word cannot change those references without the user editing
# the group's contents.
function Remove-OurTabs($doc, $tabsNode) {
    $removed = 0
    foreach ($tab in (Get-Elements $tabsNode)) {
        if ($tab.LocalName -ne "tab") { continue }

        $ours = ($tab.GetAttribute("id") -like "vt_tab_*")
        if (-not $ours) {
            foreach ($n in $tab.SelectNodes(".//*[@idQ]")) {
                $p, $l = Split-IdQ $n.GetAttribute("idQ")
                $uri = Resolve-Uri $doc $p
                if ($uri -and $uri -like "*LPandBRL.dotm") { $ours = $true; break }
            }
        }
        if (-not $ours) { continue }

        foreach ($g in (Get-Elements $tab)) {
            if ($g.GetAttribute("id") -like "vt_grp_*") { [void]$tab.RemoveChild($g) }
        }
        if ((Get-Elements $tab).Count -eq 0) {
            [void]$tabsNode.RemoveChild($tab)
            $removed++
        } else {
            # They added a group of their own to our tab; keep the tab so their work survives.
            Write-Log "  kept $($tab.GetAttribute('id')): it still holds a group the user added"
        }
    }
    return $removed
}

# --- the entries a previous VistaType install laid down, for Mine/Restore to take back out ---
function Remove-OurEntries($doc, $shared, $fullItems) {
    $removed = 0
    $existing = Get-Elements $shared

    # 1) Anything whose namespace is the add-in itself is unambiguously ours.
    foreach ($c in $existing) {
        $idQ = $c.GetAttribute("idQ"); if (-not $idQ) { continue }
        $p, $l = Split-IdQ $idQ
        $uri = Resolve-Uri $doc $p
        if ($uri -and $uri -like "*LPandBRL.dotm") { [void]$shared.RemoveChild($c); $removed++ }
    }

    # 2) The built-in Word entries the curated toolbar imposes are indistinguishable from a
    #    user's own by inspection, so only strip them when the curated block is present in
    #    full, in template order -- a fingerprint strong enough to be sure we wrote it.
    if ($fullItems) {
        $want = @()
        foreach ($it in $fullItems) {
            $idQ = $it.GetAttribute("idQ")
            if ($idQ) { $p, $l = Split-IdQ $idQ; if ($p -eq "mso") { $want += $l } }
        }
        $have = @()
        foreach ($c in (Get-Elements $shared)) {
            $idQ = $c.GetAttribute("idQ")
            if ($idQ) { $p, $l = Split-IdQ $idQ; if ($p -eq "mso") { $have += $l } }
        }
        if ($want.Count -gt 0 -and $have.Count -ge $want.Count) {
            $prefixMatches = $true
            for ($i = 0; $i -lt $want.Count; $i++) {
                if ($have[$i] -ne $want[$i]) { $prefixMatches = $false; break }
            }
            if ($prefixMatches) {
                $n = 0
                foreach ($c in (Get-Elements $shared)) {
                    if ($n -ge $want.Count) { break }
                    $idQ = $c.GetAttribute("idQ")
                    if ($idQ) {
                        $p, $l = Split-IdQ $idQ
                        if ($p -eq "mso") { [void]$shared.RemoveChild($c); $removed++; $n++ }
                    }
                }
                Write-Log "  recognised VistaType's standard toolbar and took it back out"
            }
        }
    }

    # 3) Our separators (the special namespace) only ever come from us.
    foreach ($c in (Get-Elements $shared)) {
        if ($c.LocalName -ne "separator") { continue }
        $idQ = $c.GetAttribute("idQ"); if (-not $idQ) { continue }
        if ($idQ -match "^[^:]+:(vtsep|sep)\d+$") { [void]$shared.RemoveChild($c); $removed++ }
    }
    return $removed
}

function Apply-One([string]$Target) {
    $bak  = "$Target.vtqatbak"
    $prev = "$Target.vtqatprev"
    $exists = Test-Path -LiteralPath $Target
    if (-not $exists) { New-Item -ItemType Directory -Force -Path (Split-Path $Target) | Out-Null }

    # The pristine pre-VistaType toolbar, captured ONCE and never overwritten. On a machine
    # that has had VistaType for years this is still their true original -- it is what makes
    # "put back the toolbar I had" possible at all. Never delete it.
    if (-not (Test-Path -LiteralPath $bak)) {
        if ($exists) { Copy-Item -LiteralPath $Target -Destination $bak -Force }
        else         { New-Item -ItemType File -Path $bak -Force | Out-Null }   # empty = we created the file
    }
    # A snapshot of whatever was there a moment ago, so any single install is undoable by hand.
    if ($exists) { Copy-Item -LiteralPath $Target -Destination $prev -Force }

    # Restore is deliberately NOT done by copying the backup over the file. The backup
    # predates every ribbon change the user has made since installing - their own tabs and
    # groups, and from 3.0.34 VistaType's tabs too - so overwriting would silently discard
    # all of it. Instead the backup's toolbar entries are grafted into the LIVE document
    # further down, once it has been parsed. (3.0.33 overwrote; that was wrong.)
    $restoreFromBackup = ($Mode -eq "Restore") -and
                         (Test-Path -LiteralPath $bak) -and
                         ((Get-Item -LiteralPath $bak).Length -gt 0)
    if ($Mode -eq "Restore" -and -not $restoreFromBackup) {
        Write-Log "  no pre-VistaType toolbar was saved for $Target; starting from what is there"
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

    # Reserve prefixes for the namespaces VistaType writes, whatever the user already has.
    $vtPrefix  = Ensure-Prefix $root $doc $DotmPath "x1"
    $sepPrefix = Ensure-Prefix $root $doc $MSOX     "msox"
    $prefixMap = @{ $DotmPath = $vtPrefix; $MSOX = $sepPrefix }

    # Restore: swap the live toolbar entries for the saved ones, in place. Everything else in
    # the document -- the user's ribbon tabs and groups, and VistaType's -- is left alone.
    if ($restoreFromBackup) {
        [xml]$bakDoc = Get-Content -LiteralPath $bak -Raw
        $bakShared = $bakDoc.SelectSingleNode("//*[local-name()='sharedControls']")
        if ($bakShared) {
            foreach ($c in (Get-Elements $shared)) { [void]$shared.RemoveChild($c) }
            $n = 0
            foreach ($c in (Get-Elements $bakShared)) {
                Import-NamespaceFor $doc $bakDoc $c
                [void]$shared.AppendChild($doc.ImportNode($c, $true))
                $n++
            }
            Write-Log "  put back $n saved toolbar entry(ies) for $Target, leaving every ribbon change alone"
        } else {
            Write-Log "  the saved copy for $Target has no toolbar section; nothing to put back"
        }
    }

    $written = @()
    $writtenTabs = @()

    if ($Mode -eq "None") {
        Write-Log "  leaving the Quick Access Toolbar untouched, as chosen"
    }
    elseif ($Mode -eq "Vista") {
        # Whole replacement. No merging: their toolbar is saved, not blended.
        $full = Load-Template $FullTemplate
        foreach ($c in (Get-Elements $shared)) { [void]$shared.RemoveChild($c) }
        foreach ($it in $full.Items) {
            $el = Copy-Tree $doc $it.OwnerDocument $it $prefixMap
            [void]$shared.AppendChild($el)
            $written += $el
        }
        Write-Log "  installed VistaType's standard toolbar ($($full.Items.Count) entries) into $Target"

    } else {
        # Mine / Restore: keep what is there, take out anything a previous VistaType install
        # put in, then append our icons on the end.
        $fullItems = $null
        if ($FullTemplate -and (Test-Path -LiteralPath $FullTemplate)) {
            $fullItems = (Load-Template $FullTemplate).Items
        }
        $stripped = Remove-OurEntries $doc $shared $fullItems
        if ($stripped -gt 0) { Write-Log "  removed $stripped entry(ies) left by a previous VistaType install" }

        $icons = Load-Template $IconsTemplate
        $have = @{}
        foreach ($c in (Get-Elements $shared)) {
            $idQ = $c.GetAttribute("idQ"); if (-not $idQ) { continue }
            $p, $l = Split-IdQ $idQ
            $have["$(Resolve-Uri $doc $p)|$l"] = $true
        }
        $added = 0
        foreach ($it in $icons.Items) {
            $idQ = $it.GetAttribute("idQ")
            if ($idQ) {
                $p, $l = Split-IdQ $idQ
                $uri = if ($p -eq "x1") { $DotmPath } elseif ($p -eq "msox") { $MSOX } else { Resolve-Uri $icons.Doc $p }
                if ($have.ContainsKey("$uri|$l")) { continue }
            }
            $el = Copy-Tree $doc $it.OwnerDocument $it $prefixMap
            [void]$shared.AppendChild($el)
            $written += $el
            $added++
        }
        Write-Log "  kept the existing toolbar and added $added VistaType icon(s) to $Target"
    }

    # ---- ribbon tabs -------------------------------------------------------------------
    # Word does not list add-in tabs in Customize the Ribbon, so tabs defined inside the
    # add-in cannot be hidden, reordered or renamed. Written here instead, they are ordinary
    # custom tabs and all three become possible.
    #
    # The policy, and it is worth stating plainly: THE CONTENTS OF OUR TABS ARE OURS; THE
    # TAB'S PLACE, NAME AND VISIBILITY ARE THE USER'S. So an upgrade refreshes the buttons
    # inside an existing VistaType tab and touches nothing else about it -- a tab the user
    # moved, renamed or unticked stays moved, renamed and unticked.
    if ($Tabs -ne "Skip") {
        $tabsNode = Get-OrCreate $doc $ns $ribbon "tabs"

        if ($Tabs -eq "Install") {
            $tpl = Load-Template $TabsTemplate "tabs"
            $installed = 0; $refreshed = 0
            foreach ($t in $tpl.Items) {
                $wantId = $t.GetAttribute("id")
                $existing = $null
                foreach ($e in (Get-Elements $tabsNode)) {
                    if ($e.LocalName -eq "tab" -and $e.GetAttribute("id") -eq $wantId) { $existing = $e; break }
                }
                if ($existing) {
                    # Refresh only the groups we own; leave the tab element itself, and any
                    # group the user added to it, exactly as they are.
                    foreach ($g in (Get-Elements $existing)) {
                        if ($g.GetAttribute("id") -like "vt_grp_*") { [void]$existing.RemoveChild($g) }
                    }
                    foreach ($g in (Get-Elements $t)) {
                        [void]$existing.AppendChild((Copy-Tree $doc $t.OwnerDocument $g $prefixMap))
                    }
                    $writtenTabs += $wantId
                    $refreshed++
                } else {
                    [void]$tabsNode.AppendChild((Copy-Tree $doc $t.OwnerDocument $t $prefixMap))
                    $writtenTabs += $wantId
                    $installed++
                }
            }
            Write-Log "  ribbon tabs: $installed added, $refreshed refreshed in $Target"
        }
        elseif ($Tabs -eq "Remove") {
            $gone = Remove-OurTabs $doc $tabsNode
            Write-Log "  ribbon tabs: removed $gone from $Target"
        }
    }

    Save-Xml $doc $Target
    Write-Manifest $Target $written $doc $writtenTabs
}

# ---------------------------------------------------------------------------------------
Write-Log "Merge-Qat -Mode $Mode -Tabs $Tabs"

if ($Mode -eq "None" -and $Tabs -eq "Skip") {
    Write-Log "  nothing to do: toolbar and ribbon both left alone, as chosen"
    return
}

foreach ($t in $Targets) {
    try {
        Apply-One $t
    } catch {
        # One bad or unreadable file must not abort the install, and must not stop the other
        # location being written -- Word only reads one of the two.
        Write-Log "  !! $t : $($_.Exception.Message)"
    }
}

# Tell the add-in whether the user now has their own copies of the tabs. VtTabVisible in
# RibbonCallbacks.bas reads this: when it is 1 the tabs built into the add-in go dark, so the
# two never appear at once. Absent (a hand-copied install, or the option declined) means the
# built-in tabs show exactly as they always have -- nobody ends up with no tabs at all.
if ($Tabs -ne "Skip" -and -not $TargetPaths) {
    try {
        $key = "HKCU:\Software\VistaType LP"
        if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
        $val = if ($Tabs -eq "Install") { "1" } else { "0" }
        Set-ItemProperty -Path $key -Name "UserRibbonTabs" -Value $val -Type String
        Write-Log "  UserRibbonTabs = $val"
    } catch {
        Write-Log "  !! could not record UserRibbonTabs: $($_.Exception.Message)"
    }
}
