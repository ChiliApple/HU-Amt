#Requires -Version 5.1
<#
.SYNOPSIS
    Hilfe: F1 = Hilfe zum aktiven Reiter, ? = Uebersicht. Anzeige lokal im Browser.
.DESCRIPTION
    Reihenfolge (die erste vorhandene Seite gewinnt):
      Reiter:     Data\Docs\<Id>\index.html  ->  <Modulordner>\Hilfe.html  ->  Docs\Grundhilfe.html
      Uebersicht: Data\Docs\Allgemein\index.html  ->  Docs\Grundhilfe.html
    Data\Docs kommt aus dem privaten Hilfe-Repository (docsRepo in Config\app.json); ohne Zugang gilt die Grundhilfe.
.NOTES
    Dot-Source aus Main.ps1. Zielmaschine: der PC, auf dem die App laeuft.
#>

# Seite fuer einen Bereich bestimmen. $Info: Modul (aus Read-HUAddonInfo) oder $null fuer die Uebersicht
function Get-HUHelpPage($Info = $null) {
    $docs = Join-Path $script:DataRoot 'Docs'
    $cand = @()
    if ($Info) {
        $cand += (Join-Path (Join-Path $docs $Info.Id) 'index.html')
        $cand += (Join-Path $Info.Dir 'Hilfe.html')
    } else {
        $cand += (Join-Path (Join-Path $docs 'Allgemein') 'index.html')
    }
    $cand += (Join-Path $script:AppRoot 'Docs\Grundhilfe.html')
    foreach ($p in $cand) { if (Test-Path -LiteralPath $p -PathType Leaf) { return $p } }
    return ''
}

function Show-HUHelp([switch]$Overview) {
    $info = $null
    if (-not $Overview) {
        $sel = $script:Controls['tabMain'].SelectedItem
        if ($sel -and $script:AddonByTab.ContainsKey("$($sel.Name)")) { $info = $script:AddonByTab["$($sel.Name)"] }
    }
    $p = Get-HUHelpPage $info
    if (-not $p) { Show-HUMessage 'Für diesen Bereich gibt es noch keine Hilfe.' -Icon Info; return }
    Open-HUUrl $p
}

# Knopf ? wie in der Vorlage: Klick = Anleitung, Rechtsklick = Menue mit den Seiten (Info steht im Menue des Update-Knopfs).
# Abweichung: die Anleitung gilt dem aktuellen Reiter (jedes Modul hat seine eigene Hilfe), dazu die Uebersicht.
function Register-HUHelpHandlers {
    $b = $script:Controls['btnHelp']
    $b.Add_Click({ Show-HUHelp })
    $cm = New-Object System.Windows.Controls.ContextMenu
    foreach ($p in @(@('Anleitung (F1)', 'tab', 'Hilfe zum aktuellen Reiter'), @('Übersicht', 'all', 'Was die App kann und welche Module es gibt'))) {
        $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = $p[0]; $mi.Tag = $p[1]; $mi.ToolTip = $p[2]
        $mi.Add_Click({ param($s, $e) if ("$($s.Tag)" -eq 'all') { Show-HUHelp -Overview } else { Show-HUHelp } }); [void]$cm.Items.Add($mi)
    }
    # ab 1.0.4: Neuigkeiten der installierten Version (Programm und Module)
    $mn = New-Object System.Windows.Controls.MenuItem; $mn.Header = 'Neu in dieser Version'; $mn.ToolTip = 'Was sich in dieser Version von HU-Amt und den Modulen geändert hat'
    $mn.Add_Click({ Show-HUChangelog }); [void]$cm.Items.Add($mn)
    $b.ContextMenu = $cm
    $b.ToolTip = "Anleitung (F1)`nRechtsklick: Übersicht, Neu in dieser Version"
}
