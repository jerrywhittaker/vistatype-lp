<#
Remove-Qat.ps1  --  run by the uninstaller, per-user.

Takes VistaType's icons back off the Quick Access Toolbar and leaves everything else alone.

What it will NOT do, and why: the previous version copied the pre-install backup over the
live file wholesale. That reverted the toolbar to its state on the day VistaType was
installed, silently discarding every icon -- and every ribbon customization -- the user had
added in the months since. An uninstaller has no business undoing work that was never ours.

So instead:
  * Merge-Qat wrote a manifest of exactly what it put there. Remove exactly that.
  * No manifest (installed by 3.0.32 or earlier)? Fall back to identifying our entries by
    namespace: their idQ prefix resolves to the path of LPandBRL.dotm.
  * If removing our entries would leave the toolbar EMPTY, VistaType had replaced it whole
    (the "VistaType toolbar" choice), so put the saved original back -- and re-append
    anything the user added on top of ours, which the saved original predates.
  * Keep <file>.vtqatbak. It is the only copy of the toolbar they had before VistaType, and
    it costs nothing to leave behind.

Cleans BOTH the Roaming and Local copies; Word reads one or the other depending on the
machine. Word must not be running.
#>
param(
    # Testing only; the uninstaller never passes this. See the same note in Merge-Qat.ps1.
    [string[]]$TargetPaths
)
$ErrorActionPreference = "Stop"

$MSO   = "http://schemas.microsoft.com/office/2009/07/customui"
$XMLNS = "http://www.w3.org/2000/xmlns/"
$LogDir  = Join-Path $env:APPDATA "VistaType LP"
$LogFile = Join-Path $LogDir "qat.log"

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
function Resolve-Uri($doc, $prefix) {
    if ($prefix -eq "mso")  { return $MSO }
    if (-not $prefix) { return $null }
    $n = $doc.DocumentElement.GetAttributeNode($prefix, $XMLNS)
    if ($n) { return $n.Value }
    return $null
}
function Get-Elements($node) {
    # Filter to elements. A comment or whitespace node has no GetAttribute, and with
    # ErrorActionPreference=Stop that would throw and abort the uninstall.
    @($node.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })
}
function Key($doc, $el) {
    $idQ = $el.GetAttribute("idQ")
    if (-not $idQ) { return $null }
    $p, $l = Split-IdQ $idQ
    return "$($el.LocalName)|$(Resolve-Uri $doc $p)|$l"
}
function Save-Xml($doc, $path) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)   # UTF-8, no BOM
    $writer = [System.Xml.XmlWriter]::Create($path, $settings)
    try { $doc.Save($writer) } finally { $writer.Close() }
}
function Shared-Of($doc) {
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("mso", $MSO)
    return $doc.SelectSingleNode("//mso:qat/mso:sharedControls", $ns)
}

# --- take VistaType's tabs off the user's ribbon ---
# A tab is ours if it still carries a control pointing into the add-in. That fingerprint
# survives the user reordering, renaming or unticking the tab, any of which may make Word
# rewrite the tab's id -- it cannot rewrite those references without editing the contents.
# A tab id of our own shape counts too, for a tab whose groups they emptied.
function Remove-OurTabs($doc) {
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("mso", $MSO)
    $tabsNode = $doc.SelectSingleNode("//mso:ribbon/mso:tabs", $ns)
    if (-not $tabsNode) { return 0 }

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
            Write-Log "  kept $($tab.GetAttribute('id')): it still holds a group the user added"
        }
    }
    return $removed
}

# --- our own bookkeeping goes; their saved original stays ---
function Cleanup-Bookkeeping($man, $prev, $bak, $Target) {
    foreach ($f in @($man, $prev)) {
        if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force }
    }
    if (Test-Path -LiteralPath $bak) {
        Write-Log "  $Target : leaving $([System.IO.Path]::GetFileName($bak)) in place (their pre-VistaType toolbar)"
    }
}

function Clean-One([string]$Target) {
    $bak  = "$Target.vtqatbak"
    $man  = "$Target.vtqatmanifest"
    $prev = "$Target.vtqatprev"

    if (-not (Test-Path -LiteralPath $Target)) {
        Write-Log "  $Target : nothing there"
        return
    }

    # What did we write? Manifest if we have one, otherwise identify by namespace.
    $ourKeys = @{}
    $mode = "unknown"
    if (Test-Path -LiteralPath $man) {
        foreach ($line in (Get-Content -LiteralPath $man)) {
            if ($line -like "mode=*")  { $mode = $line.Substring(5) }
            if ($line -like "item=*")  { $ourKeys[$line.Substring(5)] = $true }
        }
        Write-Log "  $Target : manifest found (installed as '$mode', $($ourKeys.Count) entries)"
    } else {
        Write-Log "  $Target : no manifest; identifying our entries by add-in path"
    }

    [xml]$doc = Get-Content -LiteralPath $Target -Raw

    # Ribbon tabs first, and INDEPENDENTLY of the toolbar. Someone who chose "leave my
    # toolbar alone" still has our tabs, and returning early because there is no toolbar
    # section would strand them on the ribbon forever.
    $tabsGone = Remove-OurTabs $doc

    $shared = Shared-Of $doc
    if (-not $shared) {
        if ($tabsGone -gt 0) {
            Save-Xml $doc $Target
            Write-Log "  $Target : removed $tabsGone VistaType ribbon tab(s); no toolbar section to clean"
        } else {
            Write-Log "  $Target : nothing of ours here"
        }
        Cleanup-Bookkeeping $man $prev $bak $Target
        return
    }

    $removed = 0
    foreach ($c in (Get-Elements $shared)) {
        $isOurs = $false
        if ($ourKeys.Count -gt 0) {
            $k = Key $doc $c
            if ($k -and $ourKeys.ContainsKey($k)) { $isOurs = $true }
        } else {
            # Fallback: our controls are the ones whose namespace IS the add-in, plus the
            # legacy VT_ markers from an older scheme.
            $idQ = $c.GetAttribute("idQ")
            if ($idQ) {
                $isOurs = ($idQ -like "x1:VT_*" -or $idQ -like "msox:VT_*")
                if (-not $isOurs) {
                    $p, $l = Split-IdQ $idQ
                    $uri = Resolve-Uri $doc $p
                    if ($uri -and $uri -like "*LPandBRL.dotm") { $isOurs = $true }
                }
            }
        }
        if ($isOurs) { [void]$shared.RemoveChild($c); $removed++ }
    }

    $left = @(Get-Elements $shared)
    $haveBackup = (Test-Path -LiteralPath $bak) -and ((Get-Item -LiteralPath $bak).Length -gt 0)

    # VistaType had replaced their toolbar whole if the install was "Vista", or if taking our
    # entries out leaves nothing at all. Either way the saved original is what they want back.
    # They may also have added icons on top of ours since; those must survive too, so start
    # from the saved original and re-append anything of theirs it does not already contain.
    if ($haveBackup -and ($mode -eq "Vista" -or $left.Count -eq 0)) {
        # ALWAYS work in the live document, never in the backup.
        #
        # The obvious shortcut - load the backup, add the leftovers, save the backup over the
        # target - loses everything in the live file that is not a toolbar entry. The backup
        # predates every ribbon change the user has made since installing: their own tabs and
        # groups, and from 3.0.34 VistaType's tabs too. Saving it over the top throws all of
        # that away silently, which is the opposite of what this script promises. (It did
        # exactly that in 3.0.33.)
        [xml]$bakDoc = Get-Content -LiteralPath $bak -Raw
        $bakShared = Shared-Of $bakDoc
        if ($bakShared) {
            $keptKeys = @{}
            foreach ($c in $left) { $k = Key $doc $c; if ($k) { $keptKeys[$k] = $true } }

            # The saved entries go back at the FRONT, in their original order, ahead of
            # anything the user added on top of ours.
            $restored = 0
            $after = $null
            foreach ($c in (Get-Elements $bakShared)) {
                $k = Key $bakDoc $c
                if ($k -and $keptKeys.ContainsKey($k)) { continue }   # already there
                # Carry across the namespace declaration this entry's prefix depends on,
                # or its idQ would dangle in the live document.
                $idQ = $c.GetAttribute("idQ")
                if ($idQ) {
                    $p, $l = Split-IdQ $idQ
                    if ($p -and $p -ne "mso") {
                        $uri = Resolve-Uri $bakDoc $p
                        if ($uri -and -not $doc.DocumentElement.GetAttributeNode($p, $XMLNS)) {
                            $decl = $doc.CreateAttribute("xmlns", $p, $XMLNS)
                            $decl.Value = $uri
                            [void]$doc.DocumentElement.Attributes.Append($decl)
                        }
                    }
                }
                $imported = $doc.ImportNode($c, $true)
                if ($after) { [void]$shared.InsertAfter($imported, $after) }
                else        { [void]$shared.PrependChild($imported) }
                $after = $imported
                $restored++
            }
            Save-Xml $doc $Target
            Write-Log ("  {0} : put back {1} saved toolbar entry(ies), keeping {2} added since, and every ribbon change" -f $Target, $restored, $left.Count)
        } else {
            # The backup has no toolbar section, so there is nothing to put back. Keep the
            # live document (ours already stripped) rather than copying the backup over it.
            Save-Xml $doc $Target
            Write-Log "  $Target : nothing saved to put back; removed our entries only"
        }
    }
    elseif ($left.Count -eq 0 -and (Test-Path -LiteralPath $bak) -and (Get-Item -LiteralPath $bak).Length -eq 0) {
        # Empty sentinel: there was no toolbar file before us and nothing of theirs is left,
        # so take the file away again. (If they HAD added icons, $left would not be empty and
        # we would keep the file -- the old version deleted it regardless and lost them.)
        Remove-Item -LiteralPath $Target -Force
        Write-Log "  $Target : removed the file VistaType created (nothing of the user's in it)"
    }
    else {
        Save-Xml $doc $Target
        Write-Log "  $Target : removed $removed VistaType entry(ies); kept $($left.Count) of the user's"
    }

    if ($tabsGone -gt 0) { Write-Log "  $Target : removed $tabsGone VistaType ribbon tab(s)" }
    Cleanup-Bookkeeping $man $prev $bak $Target
}

Write-Log "Remove-Qat"

# The add-in may well still be installed (Word can disable it, or the user may reinstall), so
# tell it the user no longer has their own copies of the tabs. VtTabVisible then shows the
# built-in ones again, rather than leaving the ribbon with no VistaType tabs at all.
if (-not $TargetPaths) {
    try {
        $key = "HKCU:\Software\VistaType LP"
        if (Test-Path $key) {
            Set-ItemProperty -Path $key -Name "UserRibbonTabs" -Value "0" -Type String
            Write-Log "  UserRibbonTabs = 0 (the add-in's own tabs show again)"
        }
    } catch {
        Write-Log "  !! could not clear UserRibbonTabs: $($_.Exception.Message)"
    }
}

foreach ($t in $Targets) {
    try {
        Clean-One $t
    } catch {
        # A malformed toolbar file must not take the uninstall down with it, and must not
        # stop the other location being cleaned.
        Write-Log "  !! $t : $($_.Exception.Message)"
    }
}
