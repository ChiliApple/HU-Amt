#Requires -Version 5.1
<#
.SYNOPSIS
    Fensterzustand merken: Position, Groesse, maximiert, Reiter, Reiter-Reihenfolge, Oberflaechen-Groesse.
.DESCRIPTION
    Datei: Data\Benutzer\<Name>\ui-state.json (je Windows-Benutzer, auch im gemeinsamen Netzordner getrennt).
    Wiederherstellung vor dem Anzeigen des Fensters: Position nur, wenn sie auf einem vorhandenen Bildschirm liegt.
.NOTES
    Dot-Source aus Main.ps1 (nach dem Setzen von $script:UserRoot). Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:StateFilePath = Join-Path $script:UserRoot 'ui-state.json'
$script:UIState = $null

function Get-HUUIState {
    if ($script:UIState) { return $script:UIState }
    $s = Read-HUJsonFile $script:StateFilePath
    if (-not $s) { $s = [pscustomobject]@{} }
    $script:UIState = $s
    return $s
}
function Get-HUStateValue([string]$Name, $Default = $null) {
    $s = Get-HUUIState
    if ($s.PSObject.Properties[$Name] -and $null -ne $s.$Name -and "$($s.$Name)" -ne '') { return $s.$Name }
    return $Default
}
function Set-HUStateValue([string]$Name, $Value) {
    $s = Get-HUUIState
    if ($s.PSObject.Properties[$Name]) { $s.$Name = $Value } else { $s | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
}
function Save-HUUIState {
    # nach dem Wiederherstellen einer App-Sicherung nichts mehr aus dem Speicher zurueckschreiben (bis zum Neustart)
    if ((Get-Variable -Name AbRestored -Scope Script -ErrorAction SilentlyContinue).Value) { return }
    try { Write-HUJsonFile -Path $script:StateFilePath -Object (Get-HUUIState) -Depth 5 } catch { Write-Verbose "[UIState] Speichern fehlgeschlagen: $_" }
}

# Rechteck (WPF-Einheiten) liegt zumindest teilweise auf dem virtuellen Bildschirm (alle Monitore)?
function Test-HUOnScreen([double]$Left, [double]$Top, [double]$Width, [double]$Height) {
    $vl = [System.Windows.SystemParameters]::VirtualScreenLeft
    $vt = [System.Windows.SystemParameters]::VirtualScreenTop
    $vw = [System.Windows.SystemParameters]::VirtualScreenWidth
    $vh = [System.Windows.SystemParameters]::VirtualScreenHeight
    # Titelleiste (obere 40 px, mittlere 100 px) muss erreichbar sein
    $cx = $Left + [Math]::Min(100.0, $Width / 2)
    return ($cx -ge $vl -and $cx -le ($vl + $vw - 20) -and $Top -ge ($vt - 10) -and ($Top + 40) -le ($vt + $vh))
}

function Restore-HUWindowState {
    $w = $script:Window
    $left = Get-HUStateValue 'Left'
    $top = Get-HUStateValue 'Top'
    $width = Get-HUStateValue 'Width'
    $height = Get-HUStateValue 'Height'
    try {
        if ($width -and $height) {
            $wa = [System.Windows.SystemParameters]::WorkArea
            $w.Width = [Math]::Max([double]$w.MinWidth, [Math]::Min([double]$width, [System.Windows.SystemParameters]::VirtualScreenWidth))
            $w.Height = [Math]::Max([double]$w.MinHeight, [Math]::Min([double]$height, [System.Windows.SystemParameters]::VirtualScreenHeight))
            if ($null -ne $left -and $null -ne $top -and (Test-HUOnScreen ([double]$left) ([double]$top) $w.Width $w.Height)) {
                $w.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
                $w.Left = [double]$left
                $w.Top = [double]$top
            } elseif ($w.Height -gt $wa.Height) { $w.Height = $wa.Height }
        }
        if ((Get-HUStateValue 'Maximized' $false) -eq $true) { $w.WindowState = [System.Windows.WindowState]::Maximized }
    } catch { Write-Verbose "[UIState] Fenster: $_" }

    Set-HUUiScale ([double](Get-HUStateValue 'UiScale' 1.0)) -NoSave
    try { Initialize-HUTabDrag } catch { Write-Verbose "[UIState] Reiter: $_" }
}

function Save-HUWindowState {
    $w = $script:Window
    $c = $script:Controls
    try {
        $isMax = ($w.WindowState -eq [System.Windows.WindowState]::Maximized)
        $rb = if ($w.WindowState -ne [System.Windows.WindowState]::Normal) { $w.RestoreBounds } else { [System.Windows.Rect]::new($w.Left, $w.Top, $w.Width, $w.Height) }
        if (-not $rb.IsEmpty -and $rb.Width -gt 100) {
            Set-HUStateValue 'Left' ([Math]::Round($rb.Left)); Set-HUStateValue 'Top' ([Math]::Round($rb.Top))
            Set-HUStateValue 'Width' ([Math]::Round($rb.Width)); Set-HUStateValue 'Height' ([Math]::Round($rb.Height))
        }
        Set-HUStateValue 'Maximized' $isMax
        $sel = $c['tabMain'].SelectedItem
        if ($sel -and "$($sel.Name)") { Set-HUStateValue 'LastTab' "$($sel.Name)" }
    } catch { Write-Verbose "[UIState] Erfassen: $_" }
    Save-HUUIState
}

# Oberflaechen-Groesse (Strg + Mausrad, Strg + 0) - skaliert den gesamten Fensterinhalt
function Set-HUUiScale([double]$Scale, [switch]$NoSave) {
    $Scale = [Math]::Round([Math]::Max(0.8, [Math]::Min(1.6, $Scale)), 2)
    $script:UiScale = $Scale
    $root = $script:Controls['rootPanel']
    if ($root) { $root.LayoutTransform = [System.Windows.Media.ScaleTransform]::new($Scale, $Scale) }
    if (-not $NoSave) { Set-HUStateValue 'UiScale' $Scale }
}

# ============================================================================
# Reiter verschieben (Ziehen mit der Maus), Reihenfolge in ui-state.json (TabOrder = x:Name der Reiter)
# ============================================================================
$script:HUTabDefault = @()

# Reine Funktion: gespeicherte Reihenfolge auf die vorhandenen Reiter anwenden.
# Unbekannte Namen fallen weg, neue Reiter (spaetere Versionen, neue Module) kommen an ihrer Standardstelle hinten dazu.
function Get-HUTabOrder([string[]]$Saved, [string[]]$Current) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($Saved)) { if ($n -and $Current -contains $n -and -not $out.Contains($n)) { $out.Add($n) } }
    foreach ($n in @($Current)) { if (-not $out.Contains($n)) { $out.Add($n) } }
    return @($out.ToArray())
}

function Set-HUTabOrder([string[]]$Names) {
    $tc = $script:Controls['tabMain']
    if (-not $tc) { return }
    $sel = $tc.SelectedItem
    $items = @($tc.Items | Where-Object { $_ -is [System.Windows.Controls.TabItem] })
    $order = Get-HUTabOrder -Saved $Names -Current @($items | ForEach-Object { "$($_.Name)" })
    $i = 0
    foreach ($n in $order) {
        $t = @($items | Where-Object { "$($_.Name)" -eq $n })[0]
        if (-not $t) { continue }
        if ($tc.Items.IndexOf($t) -ne $i) { $tc.Items.Remove($t); [void]$tc.Items.Insert($i, $t) }
        $i++
    }
    if ($sel) { $tc.SelectedItem = $sel }
}

function Save-HUTabOrder {
    $names = @($script:Controls['tabMain'].Items | Where-Object { $_ -is [System.Windows.Controls.TabItem] } | ForEach-Object { "$($_.Name)" })
    if (($names -join ',') -eq ($script:HUTabDefault -join ',')) {
        $s = Get-HUUIState; if ($s.PSObject.Properties['TabOrder']) { $s.PSObject.Properties.Remove('TabOrder') }
    } else { Set-HUStateValue 'TabOrder' $names }
    Save-HUUIState
}

function Initialize-HUTabDrag {
    $tc = $script:Controls['tabMain']
    if (-not $tc) { return }
    $script:HUTabDefault = @($tc.Items | Where-Object { $_ -is [System.Windows.Controls.TabItem] } | ForEach-Object { "$($_.Name)" })
    try { $saved = @(Get-HUStateValue 'TabOrder' @()); if ($saved.Count) { Set-HUTabOrder -Names $saved } } catch { Write-Verbose "[UIState] Reiter-Reihenfolge: $_" }
    $script:HUTabDragStart = $null; $script:HUTabDragItem = $null
    foreach ($t in @($tc.Items | Where-Object { $_ -is [System.Windows.Controls.TabItem] })) {
        $t.AllowDrop = $true
        if (-not $t.ToolTip) { $t.ToolTip = 'Reiter mit der Maus an eine andere Stelle ziehen; Rechtsklick: Reihenfolge zuruecksetzen' }
        # Hinweis nur ueber dem Reiter-Kopf: der Inhalt des Reiters gehoert logisch zum Reiter und wuerde ihn sonst ueberall
        # dort zeigen, wo ein Element keinen eigenen Hinweis hat (z.B. Spaltenkoepfe von Tabellen)
        $t.Add_ToolTipOpening({
                param($s, $e)
                try {
                    $p = [System.Windows.Input.Mouse]::GetPosition($s)
                    if ($p.X -lt 0 -or $p.Y -lt 0 -or $p.X -gt $s.ActualWidth -or $p.Y -gt $s.ActualHeight) { $e.Handled = $true }
                } catch { }
            })
        $t.Add_PreviewMouseLeftButtonDown({
                param($s, $e)
                # nur auf dem Reiter-Kopf ziehen - Klicks im Inhalt (Scrollbalken, Tabellen) nicht abfangen
                $script:HUTabDragItem = $null; $script:HUTabDragStart = $null
                $p = $e.GetPosition($s)
                if ($p.X -lt 0 -or $p.Y -lt 0 -or $p.X -gt $s.ActualWidth -or $p.Y -gt $s.ActualHeight) { return }
                $script:HUTabDragStart = $e.GetPosition($null); $script:HUTabDragItem = $s
            })
        $t.Add_PreviewMouseLeftButtonUp({ $script:HUTabDragItem = $null; $script:HUTabDragStart = $null })
        $t.Add_PreviewMouseMove({
                param($s, $e)
                if ($e.LeftButton -ne 'Pressed' -or -not $script:HUTabDragItem -or $null -eq $script:HUTabDragStart) { return }
                $p = $e.GetPosition($null)
                if ([Math]::Abs($p.X - $script:HUTabDragStart.X) -lt [System.Windows.SystemParameters]::MinimumHorizontalDragDistance -and
                    [Math]::Abs($p.Y - $script:HUTabDragStart.Y) -lt [System.Windows.SystemParameters]::MinimumVerticalDragDistance) { return }
                $src = $script:HUTabDragItem; $script:HUTabDragItem = $null; $script:HUTabDragStart = $null
                try { [void][System.Windows.DragDrop]::DoDragDrop($src, $src, [System.Windows.DragDropEffects]::Move) } catch { Write-Verbose "[UIState] Ziehen: $_" }
            })
        $t.Add_Drop({
                param($s, $e)
                try {
                    $src = $e.Data.GetData([System.Windows.Controls.TabItem])
                    if (-not $src -or $src -eq $s) { return }
                    $tcm = $script:Controls['tabMain']
                    $idx = $tcm.Items.IndexOf($s)
                    $tcm.Items.Remove($src); [void]$tcm.Items.Insert($idx, $src)
                    $tcm.SelectedItem = $src
                    Save-HUTabOrder
                } catch { Write-Verbose "[UIState] Ablegen: $_" }
            })
        $cm = New-Object System.Windows.Controls.ContextMenu
        $mi = New-Object System.Windows.Controls.MenuItem; $mi.Header = 'Reiter-Reihenfolge zuruecksetzen'
        $mi.Add_Click({ Set-HUTabOrder -Names $script:HUTabDefault; Save-HUTabOrder })
        [void]$cm.Items.Add($mi)
        $t.ContextMenu = $cm
    }
}

# Start-Reiter: zuletzt benutzter Reiter, sonst der erste
function Select-HUStartTab {
    $tc = $script:Controls['tabMain']
    if (-not $tc -or -not $tc.Items.Count) { return }
    $last = "$(Get-HUStateValue 'LastTab' '')"
    $t = $null
    if ($last) { $t = @($tc.Items | Where-Object { $_ -is [System.Windows.Controls.TabItem] -and "$($_.Name)" -eq $last })[0] }
    if (-not $t) { $t = @($tc.Items | Where-Object { $_ -is [System.Windows.Controls.TabItem] })[0] }
    if ($t) { $tc.SelectedItem = $t }
}

# Unterfenster: Groesse/Position merken
function Restore-HUDialogState($Window, [string]$Key) {
    $g = Get-HUStateValue $Key
    if (-not $g) { return }
    try {
        # gespeichert wird die Groesse OHNE Oberflaechengroesse (Set-HUWindowScale skaliert beim Laden) - sonst waechst das Fenster bei
        # jedem Oeffnen um den Faktor
        if ($g.Width -and $g.Height) {
            $sc = Get-HUUiScaleValue; $wa = [System.Windows.SystemParameters]::WorkArea
            $vw = $(if ($g.PSObject.Properties['Unscaled'] -and $g.Unscaled) { [double]$g.Width } else { [double]$g.Width / $sc })
            $vh = $(if ($g.PSObject.Properties['Unscaled'] -and $g.Unscaled) { [double]$g.Height } else { [double]$g.Height / $sc })
            $Window.Width = [Math]::Max([Math]::Max(200.0, $Window.MinWidth), [Math]::Min($vw, $wa.Width * 0.9 / $sc))
            $Window.Height = [Math]::Max([Math]::Max(150.0, $Window.MinHeight), [Math]::Min($vh, $wa.Height * 0.9 / $sc))
        }
        if ($null -ne $g.Left -and $null -ne $g.Top -and (Test-HUOnScreen ([double]$g.Left) ([double]$g.Top) $Window.Width $Window.Height)) {
            $Window.WindowStartupLocation = 'Manual'; $Window.Left = [double]$g.Left; $Window.Top = [double]$g.Top
        }
    } catch { }
}
function Save-HUDialogState($Window, [string]$Key) {
    try {
        $rb = if ($Window.WindowState -ne 'Normal') { $Window.RestoreBounds } else { [System.Windows.Rect]::new($Window.Left, $Window.Top, $Window.Width, $Window.Height) }
        $sc = Get-HUUiScaleValue
        Set-HUStateValue $Key ([pscustomobject]@{ Left = [Math]::Round($rb.Left); Top = [Math]::Round($rb.Top); Width = [Math]::Round($rb.Width / $sc); Height = [Math]::Round($rb.Height / $sc); Unscaled = $true })
        Save-HUUIState
    } catch { }
}
