#Requires -Version 5.1
<#
.SYNOPSIS
    Gemeinsame Fenster-Bausteine fuer Module: Modul-Fenster, Auswahllisten, dunkler Kalender, Uhrzeit, Zahlenfeld,
    Text- und Auswahl-Dialog, Hintergrund-Auftraege und Lade-Kreislauf.
.DESCRIPTION
    Modul-Fenster    New-HUModuleWindow (XAML aus dem Modulordner, Thema, Zustand merken, Besitzer)
    Auswahllisten    New-HUComboItem, Set-HUComboTag, Get-HUComboTag, New-HUListItem
    Kalender         Add-HUCalendar (Popup an TextBox + Knopf, Umschalt+Pfeil = Jahr), Get-HUMonthName,
                     Get-HUCalendarDays, ConvertFrom-HUCalendarText
    Uhrzeit, Zahl    Initialize-HUTimeCombo, Get-HUTimeSlots, Format-HUNumberInput
    Dialoge          Read-HUText (New-HUTextDialog), Read-HUChoice (New-HUChoiceDialog, -DefaultIndex), Read-HUNumberInput
    Hintergrund      Start-HUBackgroundJob, Receive-HUBackgroundJobs, Stop-HUBackgroundJobs, Get-HUBackgroundJobCount
    Lade-Kreislauf   Register-HUBackgroundLoad, Start-HUBackgroundLoad, Receive-HUBackgroundLoad,
                     Stop-HUBackgroundLoad, Unregister-HUBackgroundLoad, Test-HUBackgroundLoadRunning
.NOTES
    Wird per Dot-Source aus Main.ps1 geladen (UI-Thread). Zielmaschine: der PC, auf dem die App laeuft
    (Windows PowerShell 5.1 und PowerShell 7).
    Ereignis-Handler halten ihren Zustand am Steuerelement (Tag) oder in $script:-Tabellen, nie in Funktionsvariablen
    und ohne GetNewClosure(): nicht-modale Fenster leben laenger als die Funktion, die sie gebaut hat, und GetNewClosure()
    nimmt den Handlern den Zugriff auf die Skriptfunktionen der App.
#>

# Strict Mode: Skriptvariablen beim Laden anlegen
$script:HUCtlDe = [Globalization.CultureInfo]::GetCultureInfo('de-AT')
$script:HUCtlInv = [Globalization.CultureInfo]::InvariantCulture
# Datum lesen: zweistellige Jahre fest bis 2049 (KERN-B6), unabhaengig von der Windows-Einstellung des PCs
$script:HUCtlDateCulture = [Globalization.CultureInfo]([Globalization.CultureInfo]::InvariantCulture.Clone())
$script:HUCtlDateCal = New-Object Globalization.GregorianCalendar
$script:HUCtlDateCal.TwoDigitYearMax = 2049
$script:HUCtlDateCulture.DateTimeFormat.Calendar = $script:HUCtlDateCal
# Zahlen wie in Oesterreich geschrieben (1.234,56); Windows liefert fuer de-AT je nach Version ein Leerzeichen als Tausendertrenner
$script:HUCtlNum = [Globalization.NumberFormatInfo]$script:HUCtlDe.NumberFormat.Clone()
$script:HUCtlNum.NumberDecimalSeparator = ','; $script:HUCtlNum.NumberGroupSeparator = '.'
$script:HUBgJobs = New-Object System.Collections.Generic.List[object]
$script:HUBgJobTimer = $null
$script:HUBgJobNr = 0
$script:HULoads = @{}
# Abgebrochene Auftraege, deren Stop noch laeuft (KERN-B5): Freigabe erst, wenn Stop fertig ist
$script:HUCtlStopping = New-Object System.Collections.Generic.List[object]
$script:HUCtlStopTimer = $null

# ---------------------------------------------------------------------------------------------
# Interne Hilfen
# ---------------------------------------------------------------------------------------------
function Write-HUCtlWarn([string]$Text) {
    if (Get-Command Write-HULogWarn -ErrorAction SilentlyContinue) { try { Write-HULogWarn $Text } catch { } }
    else { Write-Verbose $Text }
}
# Gruener Punkt (Set-HUBusyDot aus UI-Addons.ps1), nur wenn vorhanden
function Set-HUCtlBusy($Dot, [bool]$On) {
    if (-not $Dot) { return }
    if (Get-Command Set-HUBusyDot -ErrorAction SilentlyContinue) { try { Set-HUBusyDot $Dot $On } catch { } }
}
# Skriptvariable der App lesen, ohne dass Strict Mode bei fehlender Variable abbricht
function Get-HUCtlScriptVar([string]$Name) {
    $v = Get-Variable -Name $Name -Scope Script -ErrorAction SilentlyContinue
    if ($v) { return $v.Value }
    return $null
}
# Fehlertext ohne die Huelle "Exception calling EndInvoke ..."
function Get-HUCtlErrorText($Exception) {
    $ex = $Exception
    while ($ex -and $ex.InnerException -and ($ex -is [System.Management.Automation.MethodInvocationException] -or $ex -is [System.Reflection.TargetInvocationException])) { $ex = $ex.InnerException }
    if (-not $ex) { return '' }
    return "$($ex.Message)"
}
function Set-HUCtlStyle($Element, $Owner, [string]$Key) {
    if (-not $Owner) { return }
    try { $Element.Style = $Owner.FindResource($Key) } catch { }
}

# ---------------------------------------------------------------------------------------------
# Modul-Fenster
# ---------------------------------------------------------------------------------------------
<#
.SYNOPSIS
    Fenster eines Moduls aus dessen Ordner laden: Thema einsetzen, New-HUWindow, Groesse und Position merken, Besitzer setzen.
.PARAMETER StatePrefix
    Vorsatz fuer den Schluessel in ui-state.json, Schluessel = Vorsatz + Dateiname ohne Endung (z. B. 'reisenFenster_').
    Ohne Angabe: Name des Modulordners + 'Fenster_'.
.PARAMETER StateKey
    Ganzer Schluessel statt Vorsatz + Dateiname.
.PARAMETER Owner
    Besitzer. Ohne Angabe das Hauptfenster, wenn es sichtbar ist. -NoOwner: ohne Besitzer.
.OUTPUTS
    @{ Window; C; StateKey } wie New-HUWindow, dazu der verwendete Schluessel.
#>
function New-HUModuleWindow {
    param([Parameter(Mandatory)][string]$ModuleDir, [Parameter(Mandatory)][string]$File, [string]$StatePrefix = '', [string]$StateKey = '',
        $Owner = $null, [switch]$NoOwner)
    $path = Join-Path $ModuleDir $File
    $x = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    if (-not $StateKey) {
        $pre = $(if ($StatePrefix) { $StatePrefix } else { "$(Split-Path -Path $ModuleDir -Leaf)Fenster_" })
        $StateKey = "$pre$([IO.Path]::GetFileNameWithoutExtension($File))"
    }
    try { Restore-HUDialogState $d.Window $StateKey } catch { }
    $d.Window.Resources['huStateKey'] = $StateKey
    $d.Window.Add_Closing({ param($s, $e) try { Save-HUDialogState $s "$($s.Resources['huStateKey'])" } catch { } })
    if ($NoOwner) { try { $d.Window.Owner = $null } catch { } }
    elseif ($Owner) { try { $d.Window.Owner = $Owner } catch { } }
    else {
        $mw = Get-HUCtlScriptVar 'Window'
        try { if ($mw -and $mw -ne $d.Window -and $mw.IsVisible) { $d.Window.Owner = $mw } } catch { }
    }
    $d.StateKey = $StateKey
    return $d
}

# ---------------------------------------------------------------------------------------------
# Auswahllisten
# ---------------------------------------------------------------------------------------------
<#
.SYNOPSIS
    Eintrag fuer eine ComboBox mit Tag (Stil DarkComboBoxItem ueber $Owner).
#>
function New-HUComboItem($Owner, [string]$Text, $Tag = '', [string]$ToolTip = '') {
    $i = New-Object System.Windows.Controls.ComboBoxItem
    $i.Content = $Text; $i.Tag = $Tag
    if ($ToolTip) { $i.ToolTip = $ToolTip }
    Set-HUCtlStyle $i $Owner 'DarkComboBoxItem'
    return $i
}
<#
.SYNOPSIS
    Eintrag mit diesem Tag auswaehlen (Vergleich als Text). Nicht gefunden: erster Eintrag, mit -NoFallback nichts aendern.
.OUTPUTS
    Nichts.
#>
function Set-HUComboTag($Combo, $Tag, [switch]$NoFallback) {
    $t = "$Tag"
    for ($i = 0; $i -lt $Combo.Items.Count; $i++) {
        $it = $Combo.Items[$i]
        $v = $(if ($it -is [System.Windows.FrameworkElement]) { "$($it.Tag)" } else { "$it" })
        if ($v -eq $t) { $Combo.SelectedIndex = $i; return }
    }
    if (-not $NoFallback -and $Combo.Items.Count) { $Combo.SelectedIndex = 0 }
}
<#
.SYNOPSIS
    Tag des gewaehlten Eintrags als Text, '' wenn nichts gewaehlt ist.
#>
function Get-HUComboTag($Combo) {
    if ($Combo.SelectedItem) { return "$($Combo.SelectedItem.Tag)" }
    return ''
}
<#
.SYNOPSIS
    Eintrag fuer eine ListBox mit Tag (beliebiges Objekt) und Tooltip (Stil DarkListBoxItem).
#>
function New-HUListItem($Owner, [string]$Text, $Tag = $null, [string]$ToolTip = '') {
    $i = New-Object System.Windows.Controls.ListBoxItem
    $i.Content = $Text; $i.Tag = $Tag
    if ($ToolTip) { $i.ToolTip = $ToolTip }
    Set-HUCtlStyle $i $Owner 'DarkListBoxItem'
    return $i
}

# ---------------------------------------------------------------------------------------------
# Kalender
# ---------------------------------------------------------------------------------------------
<#
.SYNOPSIS
    Monatsname deutsch (Oesterreich, Jaenner).
#>
function Get-HUMonthName([int]$Month) {
    if ($Month -lt 1 -or $Month -gt 12) { return '' }
    return $script:HUCtlDe.DateTimeFormat.MonthNames[$Month - 1]
}
<#
.SYNOPSIS
    Datum aus dem Feld lesen: 6.11.2026, 06.11.26, 2026-11-06 (auch mit Uhrzeit dahinter). Rueckgabe [datetime] (nur Tag) oder $null.
#>
function ConvertFrom-HUCalendarText($Text) {
    if ($Text -is [datetime]) { return ([datetime]$Text).Date }
    $t = "$Text".Trim()
    if (-not $t) { return $null }
    if ($t.Length -gt 10 -and $t -match '^\d{4}-\d{2}-\d{2}') { $t = $t.Substring(0, 10) }
    $d = [datetime]::MinValue
    $fmt = [string[]]@('d.M.yyyy', 'dd.MM.yyyy', 'd.M.yy', 'dd.MM.yy', 'yyyy-MM-dd')
    if ([datetime]::TryParseExact($t, $fmt, $script:HUCtlDateCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.Date }
    return $null
}
<#
.SYNOPSIS
    Die 42 Tage (6 Wochen, Montag zuerst) fuer die Monatsansicht.
.OUTPUTS
    Objekte { Date; Day; InMonth; IsToday; IsSelected; IsWeekend }
#>
function Get-HUCalendarDays([datetime]$Month, $Selected = $null, $Today = $null) {
    $first = New-Object DateTime $Month.Year, $Month.Month, 1
    $td = $(if ($null -eq $Today) { (Get-Date).Date } else { ([datetime]$Today).Date })
    $sel = $(if ($null -eq $Selected) { $null } else { ([datetime]$Selected).Date })
    $off = (([int]$first.DayOfWeek) + 6) % 7
    $start = $first.AddDays(-$off)
    for ($i = 0; $i -lt 42; $i++) {
        $d = $start.AddDays($i)
        [pscustomobject]@{
            Date = $d; Day = $d.Day; InMonth = ($d.Month -eq $first.Month); IsToday = ($d -eq $td)
            IsSelected = ($null -ne $sel -and $d -eq $sel); IsWeekend = ($d.DayOfWeek -eq [DayOfWeek]::Saturday -or $d.DayOfWeek -eq [DayOfWeek]::Sunday)
        }
    }
}
<#
.SYNOPSIS
    Schritt der Pfeile im Kalender: 1 Monat, mit Umschalt 12 Monate (abschaltbar ueber -NoYearStep bei Add-HUCalendar).
#>
function Get-HUCalendarStep([int]$Direction, [bool]$Shift = $false) {
    $s = [Math]::Sign($Direction)
    if ($Shift) { return 12 * $s }
    return $s
}
function Get-HUCalendarDate($State, $Text) {
    if ($State.Parse) {
        try { $r = & $State.Parse $Text; if ($r -is [datetime]) { return $r.Date } } catch { }
        return $null
    }
    return (ConvertFrom-HUCalendarText $Text)
}
# Monatsansicht neu zeichnen ($State aus Add-HUCalendar)
function Update-HUCalendarView($State) {
    $S = $State
    $S.Grid.Children.Clear()
    $S.Title.Text = "$(Get-HUMonthName $S.Month.Month) $($S.Month.Year)"
    foreach ($n in 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So') {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $n; $t.Foreground = Get-HUBrush '#858585' 'Fg'; $t.TextAlignment = 'Center'; $t.Margin = '0,0,0,4'
        [void]$S.Grid.Children.Add($t)
    }
    $sel = Get-HUCalendarDate $S $S.Box.Text
    foreach ($c in @(Get-HUCalendarDays -Month $S.Month -Selected $sel)) {
        $b = New-Object System.Windows.Controls.Button
        Set-HUCtlStyle $b $S.Owner 'DarkButton'
        $b.Content = "$($c.Day)"; $b.Tag = $c.Date; $b.Width = 34; $b.Height = 28; $b.Padding = '0'; $b.Margin = '1'
        $b.ToolTip = $c.Date.ToString('dddd, dd.MM.yyyy', $script:HUCtlDe)
        # Tage wie Felder: im hellen Schema helle Flaeche mit dunkler Schrift, nur der gewaehlte Tag bleibt blau mit weisser Schrift
        $b.Foreground = Get-HUBrush '#FFFFFF' 'Fg'
        if (-not $c.InMonth) { $b.Foreground = Get-HUBrush '#6A6A6A' 'Fg' }
        if ($c.IsToday) { $b.BorderBrush = Get-HUBrush '#FFB74D' 'Fg'; $b.BorderThickness = 1 }
        if ($c.IsSelected) { $b.Background = Get-HUBrush '#007ACC'; $b.Foreground = [System.Windows.Media.Brushes]::White }
        elseif ($c.IsWeekend) { $b.Background = Get-HUBrush '#2D2D30' 'Bg' }
        else { $b.Background = Get-HUBrush '#3E3E42' 'Bg' }
        $b.Add_Click({
                param($s, $e)
                $x = $s.Parent.Tag
                $x.Box.Text = ([datetime]$s.Tag).ToString($x.Format, $script:HUCtlInv)
                $x.Pop.IsOpen = $false
                [void]$x.Box.Focus()
            })
        [void]$S.Grid.Children.Add($b)
    }
}
<#
.SYNOPSIS
    Dunkler Kalender (Popup) an einem Datumsfeld: Knopf oeffnet und schliesst, Klick auf einen Tag traegt ihn ein.
    Pfeile = Monat, Umschalt+Pfeil = Jahr.
.PARAMETER Owner
    Element, ueber das die Stile des Themas gefunden werden (meist das Fenster). Ohne Angabe die TextBox.
.PARAMETER Parse
    Eigener Datumsleser { param($Text) ... } -> [datetime] oder $null. Ohne Angabe ConvertFrom-HUCalendarText.
.PARAMETER Format
    Format beim Eintragen, Standard dd.MM.yyyy.
.OUTPUTS
    Zustand @{ Month; Pop; Grid; Title; Box; Owner; Parse; Format; YearStep } (haengt auch an Grid.Tag und Button.Tag).
#>
function Add-HUCalendar {
    param([Parameter(Mandatory)]$TextBox, [Parameter(Mandatory)]$Button, $Owner = $null, [scriptblock]$Parse = $null,
        [string]$Format = 'dd.MM.yyyy', [switch]$NoYearStep)
    if (-not $Owner) { $Owner = $TextBox }
    $pop = New-Object System.Windows.Controls.Primitives.Popup
    $pop.PlacementTarget = $TextBox; $pop.Placement = 'Bottom'; $pop.StaysOpen = $false; $pop.AllowsTransparency = $true
    $border = New-Object System.Windows.Controls.Border
    $border.Background = Get-HUBrush '#252526' 'Bg'; $border.BorderBrush = Get-HUBrush '#007ACC'; $border.BorderThickness = 1; $border.Padding = 6; $border.CornerRadius = 3
    $sp = New-Object System.Windows.Controls.StackPanel
    $head = New-Object System.Windows.Controls.DockPanel
    $bPrev = New-Object System.Windows.Controls.Button; $bNext = New-Object System.Windows.Controls.Button
    foreach ($b in $bPrev, $bNext) { Set-HUCtlStyle $b $Owner 'DarkButton'; $b.Width = 30; $b.Padding = '0,1'; $b.FontSize = 14 }
    $bPrev.Content = [string][char]0x2039; $bNext.Content = [string][char]0x203A
    if ($NoYearStep) { $bPrev.ToolTip = 'Voriger Monat'; $bNext.ToolTip = 'Nächster Monat' }
    else { $bPrev.ToolTip = 'Voriger Monat (Umschalt: voriges Jahr)'; $bNext.ToolTip = 'Nächster Monat (Umschalt: nächstes Jahr)' }
    [System.Windows.Controls.DockPanel]::SetDock($bPrev, 'Left'); [System.Windows.Controls.DockPanel]::SetDock($bNext, 'Right')
    $title = New-Object System.Windows.Controls.TextBlock
    $title.Foreground = Get-HUBrush '#FFFFFF' 'Fg'; $title.FontWeight = 'SemiBold'; $title.TextAlignment = 'Center'; $title.VerticalAlignment = 'Center'
    [void]$head.Children.Add($bPrev); [void]$head.Children.Add($bNext); [void]$head.Children.Add($title)
    $grid = New-Object System.Windows.Controls.Primitives.UniformGrid
    $grid.Columns = 7; $grid.Margin = '0,6,0,0'
    [void]$sp.Children.Add($head); [void]$sp.Children.Add($grid)
    $border.Child = $sp; $pop.Child = $border
    $today = (Get-Date).Date
    $st = @{ Month = $today.AddDays(1 - $today.Day); Pop = $pop; Grid = $grid; Title = $title; Box = $TextBox; Owner = $Owner
        Parse = $Parse; Format = $Format; YearStep = (-not $NoYearStep)
    }
    $grid.Tag = $st; $bPrev.Tag = $st; $bNext.Tag = $st; $Button.Tag = $st
    if (-not $Button.ToolTip) { $Button.ToolTip = 'Kalender öffnen' }
    if (-not $TextBox.ToolTip) { $TextBox.ToolTip = 'Datum (TT.MM.JJJJ) eintippen oder im Kalender wählen' }
    $bPrev.Add_Click({
            param($s, $e)
            $x = $s.Tag
            $sh = [bool]$x.YearStep -and (([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0)
            $x.Month = $x.Month.AddMonths((Get-HUCalendarStep -1 $sh))
            Update-HUCalendarView $x
        })
    $bNext.Add_Click({
            param($s, $e)
            $x = $s.Tag
            $sh = [bool]$x.YearStep -and (([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0)
            $x.Month = $x.Month.AddMonths((Get-HUCalendarStep 1 $sh))
            Update-HUCalendarView $x
        })
    $Button.Add_Click({
            param($s, $e)
            $x = $s.Tag
            if ($x.Pop.IsOpen) { $x.Pop.IsOpen = $false; return }
            $d = Get-HUCalendarDate $x $x.Box.Text
            if (-not $d) { $d = (Get-Date).Date }
            $x.Month = $d.AddDays(1 - $d.Day)
            Update-HUCalendarView $x
            $x.Pop.IsOpen = $true
        })
    return $st
}

# ---------------------------------------------------------------------------------------------
# Uhrzeit und Zahlenfeld
# ---------------------------------------------------------------------------------------------
<#
.SYNOPSIS
    Uhrzeiten eines ganzen Tages ab -StartHour im Abstand -StepMinutes (nach Mitternacht weiter bis vor die Startzeit).
.OUTPUTS
    Texte 'HH:mm'.
#>
function Get-HUTimeSlots([int]$StepMinutes = 15, [int]$StartHour = 0) {
    if ($StepMinutes -lt 1) { $StepMinutes = 15 }
    $start = ((($StartHour % 24) + 24) % 24) * 60
    for ($m = 0; $m -lt 1440; $m += $StepMinutes) {
        $mm = ($start + $m) % 1440
        '{0:00}:{1:00}' -f [Math]::Floor($mm / 60), ($mm % 60)
    }
}
<#
.SYNOPSIS
    ComboBox mit Uhrzeiten fuellen (Eintraege mit Tag ''). Standard: ab 00:00 alle 15 Minuten.
.OUTPUTS
    Nichts.
#>
function Initialize-HUTimeCombo($Combo, $Owner = $null, [int]$StepMinutes = 15, [int]$StartHour = 0, [int]$MaxDropDownHeight = 260) {
    if (-not $Owner) { $Owner = $Combo }
    foreach ($t in @(Get-HUTimeSlots -StepMinutes $StepMinutes -StartHour $StartHour)) { [void]$Combo.Items.Add((New-HUComboItem $Owner $t '')) }
    $Combo.MaxDropDownHeight = $MaxDropDownHeight
    if (-not $Combo.ToolTip) { $Combo.ToolTip = 'Uhrzeit wählen oder eintippen (HH:MM)' }
}
<#
.SYNOPSIS
    Zahl fuer ein Eingabefeld: Komma, ohne Tausenderpunkt, hoechstens -Digits Nachkommastellen ohne Nullen am Ende.
    $null oder '' -> ''. 0 -> '0', mit -ZeroAsEmpty ''.
#>
function Format-HUNumberInput($Value, [int]$Digits = 2, [switch]$ZeroAsEmpty) {
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    $d = [decimal]0
    if ($Value -is [string]) {
        $txt = $Value.Trim()
        # erst wie [decimal]'12.5' (Punkt, ohne Tausender), dann oesterreichisch (1.234,5)
        if (-not [decimal]::TryParse($txt, [Globalization.NumberStyles]::Float, $script:HUCtlInv, [ref]$d) -and
            -not [decimal]::TryParse($txt, [Globalization.NumberStyles]::Number, $script:HUCtlNum, [ref]$d)) { return $txt }
    } else { $d = [decimal]$Value }
    if ($d -eq 0 -and $ZeroAsEmpty) { return '' }
    $fmt = $(if ($Digits -gt 0) { "0.$('#' * $Digits)" } else { '0' })
    return $d.ToString($fmt, $script:HUCtlNum)
}

# ---------------------------------------------------------------------------------------------
# Dialoge (modal)
# ---------------------------------------------------------------------------------------------
$script:HUCtlTextXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="560" MinWidth="460" MinHeight="170" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip"
        Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel Margin="18">
        <TextBlock x:Name="txPrompt" DockPanel.Dock="Top" Foreground="#CCCCCC" TextWrapping="Wrap" Margin="0,0,0,8"/>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button x:Name="txOk" Content="OK" Width="100" Background="#4CAF50" Style="{StaticResource DarkButton}" IsDefault="True" Margin="0,0,8,0" ToolTip="Übernehmen"/>
            <Button x:Name="txCancel" Content="Abbrechen" Width="100" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Abbrechen, nichts übernehmen"/>
        </StackPanel>
        <TextBox x:Name="txText" Style="{StaticResource DarkTextBox}" MaxLength="300" ToolTip="Text eingeben"/>
    </DockPanel>
</Window>
'@
$script:HUCtlChoiceXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="560" MinWidth="460" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip"
        Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <StackPanel Margin="18">
        <TextBlock x:Name="chPrompt" Foreground="#CCCCCC" TextWrapping="Wrap" Margin="0,0,0,14"/>
        <WrapPanel x:Name="chButtons" HorizontalAlignment="Right"/>
    </StackPanel>
</Window>
'@

<#
.SYNOPSIS
    Textdialog bauen, ohne ihn zu zeigen (Read-HUText zeigt ihn). Zustand haengt an Window.Tag und an den Knoepfen.
.OUTPUTS
    @{ Window; C; Text ($null = abgebrochen); AllowEmpty; Multiline; StateKey }
#>
function New-HUTextDialog {
    param([string]$Title = '', [string]$Prompt = '', [string]$Default = '', $Owner = $null, [Alias('Mehrzeilig')][switch]$Multiline,
        [switch]$AllowEmpty, [string]$StateKey = '', [int]$MaxLength = 0)
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $script:HUCtlTextXaml)
    $w = $d.Window; $c = $d.C
    $w.Title = $Title
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    if (-not $StateKey) { $StateKey = $(if ($Multiline) { 'huFenster_TextListe' } else { 'huFenster_Text' }) }
    $st = @{ Window = $w; C = $c; Text = $null; AllowEmpty = [bool]$AllowEmpty; Multiline = [bool]$Multiline; StateKey = $StateKey }
    $w.Tag = $st; $c.txOk.Tag = $st; $c.txCancel.Tag = $st
    $c.txPrompt.Text = $Prompt; $c.txText.Text = $Default
    if ($Multiline) {
        $c.txText.AcceptsReturn = $true; $c.txText.MaxLength = 20000; $c.txText.MinHeight = 220; $c.txText.TextWrapping = 'NoWrap'
        $c.txText.VerticalScrollBarVisibility = 'Auto'; $c.txText.HorizontalScrollBarVisibility = 'Auto'; $c.txText.VerticalContentAlignment = 'Top'
        $c.txText.ToolTip = 'Text eingeben, eine Angabe je Zeile (Enter = neue Zeile)'
        $c.txOk.IsDefault = $false
    }
    if ($MaxLength -gt 0) { $c.txText.MaxLength = $MaxLength }
    try { Restore-HUDialogState $w $StateKey } catch { }
    # einzeilig richtet sich die Hoehe weiter nach dem Inhalt; mehrzeilig bleibt eine gemerkte Hoehe
    if (-not $Multiline) { $w.SizeToContent = 'Height' }
    elseif (-not [double]::IsNaN($w.Height)) { $w.SizeToContent = 'Manual' }
    $c.txOk.Add_Click({ param($s, $e) [void](Complete-HUTextDialog $s.Tag) })
    $w.Add_Closing({ param($s, $e) try { Save-HUDialogState $s "$($s.Tag.StateKey)" } catch { } })
    $w.Add_ContentRendered({ param($s, $e) try { [void]$s.Tag.C.txText.Focus(); $s.Tag.C.txText.SelectAll() } catch { } })
    return $st
}
<#
.SYNOPSIS
    OK im Textdialog: Text uebernehmen (getrimmt). Leer ohne -AllowEmpty: nichts tun, Rueckgabe $false.
#>
function Complete-HUTextDialog($State) {
    $t = "$($State.C.txText.Text)".Trim()
    if (-not $t -and -not $State.AllowEmpty) { try { [void]$State.C.txText.Focus() } catch { }; return $false }
    $State.Text = $t
    try { $State.Window.DialogResult = $true } catch { }   # ausserhalb von ShowDialog (Test) nicht moeglich
    return $true
}
<#
.SYNOPSIS
    Kleiner Textdialog (ein- oder mehrzeilig). Groesse und Position werden unter -StateKey gemerkt.
.OUTPUTS
    Text (getrimmt) oder $null bei Abbruch. Leer nur mit -AllowEmpty.
#>
function Read-HUText {
    param([string]$Title = '', [string]$Prompt = '', [string]$Default = '', $Owner = $null, [Alias('Mehrzeilig')][switch]$Multiline,
        [switch]$AllowEmpty, [string]$StateKey = '', [int]$MaxLength = 0)
    $st = New-HUTextDialog @PSBoundParameters
    [void]$st.Window.ShowDialog()
    return $st.Text
}

<#
.SYNOPSIS
    Auswahldialog bauen, ohne ihn zu zeigen (Read-HUChoice zeigt ihn). Erster Knopf blau hervorgehoben.
.OUTPUTS
    @{ Window; C; Index (-1 = abgebrochen); Buttons; StateKey }
#>
function New-HUChoiceDialog {
    param([string]$Title = '', [string]$Prompt = '', [string[]]$Choices = @(), $Owner = $null, [string[]]$ToolTips = @(),
        [string]$CancelText = 'Abbrechen', [string]$CancelToolTip = 'Nichts ändern', [switch]$NoCancel, [string]$StateKey = '', [int]$DefaultIndex = -1)
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $script:HUCtlChoiceXaml)
    $w = $d.Window; $c = $d.C
    $w.Title = $Title
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $c.chPrompt.Text = $Prompt
    $st = @{ Window = $w; C = $c; Index = -1; Buttons = (New-Object System.Collections.Generic.List[object]); StateKey = $StateKey }
    $w.Tag = $st
    for ($i = 0; $i -lt @($Choices).Count; $i++) {
        $b = New-Object System.Windows.Controls.Button
        Set-HUCtlStyle $b $w 'DarkButton'
        $b.Content = $Choices[$i]; $b.Tag = @{ I = $i; S = $st }; $b.Margin = '0,0,8,0'; $b.Padding = '14,4'
        $b.Background = Get-HUBrush $(if ($i -eq 0) { '#1976D2' } else { '#3E3E42' }) 'Btn'
        $b.ToolTip = $(if ($i -lt @($ToolTips).Count -and "$($ToolTips[$i])") { "$($ToolTips[$i])" } else { "$($Choices[$i])" })
        if ($i -eq $DefaultIndex) { $b.IsDefault = $true }   # Enter waehlt diesen Knopf; ohne -DefaultIndex keiner
        $b.Add_Click({ param($s, $e) $t = $s.Tag; $t.S.Index = [int]$t.I; try { $t.S.Window.DialogResult = $true } catch { } })
        [void]$c.chButtons.Children.Add($b); $st.Buttons.Add($b)
    }
    if (-not $NoCancel) {
        $cb = New-Object System.Windows.Controls.Button
        Set-HUCtlStyle $cb $w 'DarkButton'
        $cb.Content = $CancelText; $cb.IsCancel = $true; $cb.Padding = '14,4'; $cb.Background = Get-HUBrush '#555555' 'Btn'; $cb.ToolTip = $CancelToolTip
        $cb.Tag = $st
        $cb.Add_Click({ param($s, $e) $s.Tag.Index = -1; try { $s.Tag.Window.DialogResult = $false } catch { } })
        [void]$c.chButtons.Children.Add($cb)
    }
    if ($StateKey) {
        try { Restore-HUDialogState $w $StateKey } catch { }
        $w.SizeToContent = 'Height'
        $w.Add_Closing({ param($s, $e) try { Save-HUDialogState $s "$($s.Tag.StateKey)" } catch { } })
    }
    return $st
}
<#
.SYNOPSIS
    Rueckfrage mit mehreren Knoepfen. -DefaultIndex: dieser Knopf reagiert auf Enter (IsDefault); ohne Angabe keiner.
.OUTPUTS
    Nummer des Knopfs (0, 1, ...) oder -1 (Abbrechen, Esc, Fenster geschlossen).
#>
function Read-HUChoice {
    param([string]$Title = '', [string]$Prompt = '', [string[]]$Choices = @(), $Owner = $null, [string[]]$ToolTips = @(),
        [string]$CancelText = 'Abbrechen', [string]$CancelToolTip = 'Nichts ändern', [switch]$NoCancel, [string]$StateKey = '', [int]$DefaultIndex = -1)
    $st = New-HUChoiceDialog @PSBoundParameters
    [void]$st.Window.ShowDialog()
    return [int]$st.Index
}

<#
.SYNOPSIS
    Betrag aus einem Eingabefeld lesen. Bei zweideutiger Eingabe (Test-HUAmbiguousNumber, z. B. 1.500) einmal nachfragen:
    "Meinst du 1.500 (tausendfuenfhundert) oder 1,5?". Sonst wie ConvertTo-HUDecimal (HU.Data).
.DESCRIPTION
    Gedacht fuer LostFocus oder Speichern eines Betragsfelds; danach das Feld mit Format-HUNumberInput neu schreiben,
    dann ist es eindeutig (1500 bzw. 1,5) und es wird nicht noch einmal gefragt.
.OUTPUTS
    [decimal], oder $null bei leerem Feld, keiner Zahl oder Abbruch der Rueckfrage.
#>
function Read-HUNumberInput {
    param([string]$Text, $Owner = $null, [string]$Title = 'Betrag prüfen')
    $t = "$Text".Trim()
    if (-not $t -or -not (Test-HUNumberText $t)) { return $null }
    if (Test-HUAmbiguousNumber $t) {
        $clean = ($t -replace '[^\d.+-]', '')
        $big = [decimal]::Parse($clean.Replace('.', ''), [Globalization.CultureInfo]::InvariantCulture)
        $small = [decimal]::Parse($clean, [Globalization.CultureInfo]::InvariantCulture)
        $bigText = $big.ToString('#,##0', $script:HUCtlNum); $smallText = $small.ToString('0.###', $script:HUCtlNum)
        $words = " (also $($big.ToString('0', $script:HUCtlInv)))"
        $i = Read-HUChoice -Title $Title -Prompt "Meinst du $bigText$words oder $($smallText)?" -Choices @($bigText, $smallText) -Owner $Owner `
            -ToolTips @("Punkt als Tausendertrenner: $bigText", "Punkt als Komma: $smallText") -DefaultIndex 0
        if ($i -eq 0) { return $big }
        if ($i -eq 1) { return $small }
        return $null
    }
    return (ConvertTo-HUDecimal $t)
}

# ---------------------------------------------------------------------------------------------
# Hintergrund-Auftraege
# ---------------------------------------------------------------------------------------------
<#
.SYNOPSIS
    Arbeit in einem eigenen Runspace. $Done bekommt ($Ergebnis, $Fehlertext, $Kontext) im Fenster-Thread.
.DESCRIPTION
    Der Skriptblock laeuft als Text in einem frischen Runspace (keine Funktionen oder Variablen der App), Werte nur ueber
    -Arguments. Ergebnis: erstes ausgegebenes Objekt, mit -AllResults alle als Array. Fehlertext: Ausnahme oder erster
    Fehler, wenn nichts ausgegeben wurde; 'Zeitüberschreitung' nach -TimeoutSec (0 = ohne Limit).
    -Sta: eigener STA-Runspace (fuer Outlook/Word ueber COM). Ein eigener Zeitgeber (300 ms) holt fertige Auftraege ab und
    haelt an, wenn keiner mehr laeuft; -NoTimer: Abholen selbst mit Receive-HUBackgroundJobs.
.PARAMETER Owner
    Gruppe (z. B. Modul-Id) fuer Stop-HUBackgroundJobs -Owner beim Beenden des Moduls.
.OUTPUTS
    Nichts, mit -PassThru das Auftragsobjekt.
#>
function Start-HUBackgroundJob {
    param([Parameter(Mandatory)][scriptblock]$Script, [object[]]$Arguments = @(), [scriptblock]$Done = $null, [Alias('Ctx')]$Context = $null,
        [switch]$Sta, [int]$TimeoutSec = 180, [switch]$AllResults, [string]$Owner = '', [string]$LogPrefix = '', [switch]$NoTimer, [switch]$PassThru)
    $rs = $null
    $ps = [powershell]::Create()
    try {
        if ($Sta) {
            $rs = [runspacefactory]::CreateRunspace()
            $rs.ApartmentState = [System.Threading.ApartmentState]::STA
            $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
            $rs.Open()
            $ps.Runspace = $rs
        }
        [void]$ps.AddScript($Script)
        foreach ($a in $Arguments) { [void]$ps.AddArgument($a) }
        $script:HUBgJobNr++
        $j = [pscustomobject]@{ Id = $script:HUBgJobNr; Ps = $ps; Rs = $rs; Handle = $null; Done = $Done; Ctx = $Context; Start = Get-Date
            Timeout = $TimeoutSec; Sta = [bool]$Sta; All = [bool]$AllResults; Owner = $Owner; Prefix = $LogPrefix
        }
        $j.Handle = $ps.BeginInvoke()
    } catch {
        try { $ps.Dispose() } catch { }
        try { if ($rs) { $rs.Dispose() } } catch { }
        throw
    }
    $script:HUBgJobs.Add($j)
    if (-not $NoTimer) { Start-HUBackgroundJobTimer }
    if ($PassThru) { return $j }
}
function Start-HUBackgroundJobTimer {
    if (-not $script:HUBgJobTimer) {
        $t = New-Object System.Windows.Threading.DispatcherTimer
        $t.Interval = [TimeSpan]::FromMilliseconds(300)
        $t.Add_Tick({ param($s, $e) Receive-HUBackgroundJobs; if (-not $script:HUBgJobs.Count) { $s.Stop() } })
        $script:HUBgJobTimer = $t
    }
    if (-not $script:HUBgJobTimer.IsEnabled) { $script:HUBgJobTimer.Start() }
}
function Close-HUBackgroundJob($Job, [switch]$Stop) {
    if ($Stop) { Stop-HUCtlPipeline $Job.Ps $Job.Rs; return }
    try { $Job.Ps.Dispose() } catch { }
    try { if ($Job.Rs) { $Job.Rs.Dispose() } } catch { }
}
<#
.SYNOPSIS
    Auftrag abbrechen, ohne den Fenster-Thread zu blockieren (KERN-B5, wie Core-Async): BeginStop statt Stop.
    PowerShell und Runspace werden erst freigegeben, wenn der Stop fertig ist (Complete-HUCtlStops im naechsten Takt).
    Kein Skriptblock als Rueckruf: der liefe in einem Thread ohne Runspace.
#>
function Stop-HUCtlPipeline($Ps, $Rs = $null) {
    if (-not $Ps) { return }
    $h = $null
    try { $h = $Ps.BeginStop($null, $null) } catch { $h = $null }
    $script:HUCtlStopping.Add([pscustomobject]@{ Ps = $Ps; Rs = $Rs; Handle = $h })
    Complete-HUCtlStops
    if ($script:HUCtlStopping.Count) { Start-HUCtlStopTimer }
}
<#
.SYNOPSIS
    Fertig abgebrochene Auftraege freigeben (vom Zeitgeber, ohne Zeitgeber selbst aufrufen).
#>
function Complete-HUCtlStops {
    foreach ($x in $script:HUCtlStopping.ToArray()) {
        if ($x.Handle -and -not $x.Handle.IsCompleted) { continue }
        [void]$script:HUCtlStopping.Remove($x)
        if ($x.Handle) { try { $x.Ps.EndStop($x.Handle) } catch { } }
        try { $x.Ps.Dispose() } catch { }
        try { if ($x.Rs) { $x.Rs.Dispose() } } catch { }
    }
}
function Start-HUCtlStopTimer {
    try {
        if (-not $script:HUCtlStopTimer) {
            $t = New-Object System.Windows.Threading.DispatcherTimer
            $t.Interval = [TimeSpan]::FromMilliseconds(250)
            $t.Add_Tick({ param($s, $e) Complete-HUCtlStops; if (-not $script:HUCtlStopping.Count) { $s.Stop() } })
            $script:HUCtlStopTimer = $t
        }
        if (-not $script:HUCtlStopTimer.IsEnabled) { $script:HUCtlStopTimer.Start() }
    } catch { }   # ohne WPF (Hintergrund, Test): freigegeben wird beim naechsten Complete-HUCtlStops
}
function Invoke-HUBackgroundDone($Job, $Result, [string]$ErrorText) {
    if (-not $Job.Done) { return }
    try { & $Job.Done $Result $ErrorText $Job.Ctx }
    catch { Write-HUCtlWarn "$(if ($Job.Prefix) { "$($Job.Prefix): " })$($_.Exception.Message)" }
}
<#
.SYNOPSIS
    Fertige oder ueberfaellige Auftraege abholen und ihr Done aufrufen (im Fenster-Thread). -Owner: nur diese Gruppe.
.OUTPUTS
    Nichts.
#>
function Receive-HUBackgroundJobs {
    param([string]$Owner = '')
    foreach ($j in $script:HUBgJobs.ToArray()) {
        if ($Owner -and $j.Owner -ne $Owner) { continue }
        if (-not $j.Handle.IsCompleted) {
            if ($j.Timeout -gt 0 -and ((Get-Date) - $j.Start).TotalSeconds -gt $j.Timeout) {
                [void]$script:HUBgJobs.Remove($j)
                Close-HUBackgroundJob $j -Stop
                Invoke-HUBackgroundDone $j $null 'Zeitüberschreitung'
            }
            continue
        }
        [void]$script:HUBgJobs.Remove($j)
        $res = $null; $msg = ''
        try {
            $out = @($j.Ps.EndInvoke($j.Handle))
            if ($j.All) { $res = $out } elseif ($out.Count) { $res = $out[0] }
            $er = @($j.Ps.Streams.Error)
            if (-not $out.Count -and $er.Count) { $msg = "$($er[0].Exception.Message)" }
        } catch { $msg = Get-HUCtlErrorText $_.Exception }
        finally { Close-HUBackgroundJob $j }
        Invoke-HUBackgroundDone $j $res $msg
    }
}
<#
.SYNOPSIS
    Auftraege beenden (ohne Done), z. B. in Stop-HUAddon<Id>. Laufende STA-Auftraege (Outlook) duerfen bis -WaitStaSec
    fertig werden, damit sie ihre COM-Objekte selbst freigeben.
#>
function Stop-HUBackgroundJobs {
    param([string]$Owner = '', [int]$WaitStaSec = 20)
    foreach ($j in $script:HUBgJobs.ToArray()) {
        if ($Owner -and $j.Owner -ne $Owner) { continue }
        if ($j.Sta -and $WaitStaSec -gt 0 -and -not $j.Handle.IsCompleted) { try { [void]$j.Handle.AsyncWaitHandle.WaitOne($WaitStaSec * 1000) } catch { } }
        [void]$script:HUBgJobs.Remove($j)
        Close-HUBackgroundJob $j -Stop
    }
    if (-not $script:HUBgJobs.Count -and $script:HUBgJobTimer) { try { $script:HUBgJobTimer.Stop() } catch { } }
}
<#
.SYNOPSIS
    Zahl der offenen Auftraege (alle oder einer Gruppe).
#>
function Get-HUBackgroundJobCount([string]$Owner = '') {
    if (-not $Owner) { return $script:HUBgJobs.Count }
    return @($script:HUBgJobs | Where-Object { $_.Owner -eq $Owner }).Count
}

# ---------------------------------------------------------------------------------------------
# Lade-Kreislauf (Daten im Hintergrund laden, alle n Sekunden abgleichen)
# ---------------------------------------------------------------------------------------------
<#
.SYNOPSIS
    Lade-Kreislauf anmelden: Laden im eigenen Runspace, Zeitgeber, Abgleich alle -RefreshSeconds, Warteschlange
    (Start waehrend eines Ladens = einmal danach neu), Zeitueberschreitung, gruener Punkt.
.PARAMETER Script
    Laeuft als Text im frischen Runspace (wie Start-HUBackgroundJob); das erste ausgegebene Objekt ist das Ergebnis.
.PARAMETER GetArguments
    { ... } im Fenster-Thread vor jedem Laden; gibt die Argumente fuer -Script aus (ein Array-Argument mit , davor).
.PARAMETER Done
    { param($Ergebnis) ... } im Fenster-Thread nach erfolgreichem Laden. Eine Ausnahme darin gilt als "Laden fehlgeschlagen".
.PARAMETER Failed
    { param($Fehlertext) ... } nach Fehler oder Zeitueberschreitung (die Warnung im Protokoll schreibt der Kern).
.PARAMETER RefreshSeconds
    Abgleich nach so vielen Sekunden seit dem letzten Start, 0 = nur auf Anforderung.
.PARAMETER Dot
    Ellipse fuer Set-HUBusyDot (falls vorhanden).
.PARAMETER OnTick
    { ... } zusaetzlich bei jedem Takt (z. B. ein eigener Abgleich des Moduls).
.PARAMETER Start
    Gleich zum ersten Mal laden.
.OUTPUTS
    Nichts.
#>
function Register-HUBackgroundLoad {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][scriptblock]$Script, [scriptblock]$GetArguments = $null,
        [Parameter(Mandatory)][scriptblock]$Done, [scriptblock]$Failed = $null, [int]$RefreshSeconds = 60, [int]$TimeoutSec = 120,
        $Dot = $null, [string]$LogPrefix = '', [int]$TickMs = 300, [scriptblock]$OnTick = $null, [switch]$NoTimer, [switch]$Start)
    if ($script:HULoads.ContainsKey($Name)) { Unregister-HUBackgroundLoad -Name $Name }
    $l = @{ Name = $Name; Script = $Script; GetArguments = $GetArguments; Done = $Done; Failed = $Failed; Refresh = $RefreshSeconds
        Timeout = $TimeoutSec; Dot = $Dot; Prefix = $(if ($LogPrefix) { $LogPrefix } else { $Name }); Job = $null; Queued = $false
        LastLoad = [datetime]::MinValue; Timer = $null; OnTick = $OnTick
    }
    $script:HULoads[$Name] = $l
    if (-not $NoTimer) {
        $t = New-Object System.Windows.Threading.DispatcherTimer
        $t.Interval = [TimeSpan]::FromMilliseconds([Math]::Max(50, $TickMs))
        $t.Tag = $Name
        $t.Add_Tick({
                param($s, $e)
                $n = "$($s.Tag)"
                Receive-HUBackgroundLoad -Name $n
                $x = $script:HULoads[$n]
                if ($x -and $x.OnTick) { try { & $x.OnTick } catch { Write-HUCtlWarn "$($x.Prefix): $($_.Exception.Message)" } }
            })
        $l.Timer = $t
        $t.Start()
    }
    if ($Start) { Start-HUBackgroundLoad -Name $Name }
}
<#
.SYNOPSIS
    Laden starten. Laeuft schon eines: vormerken, nach dessen Ende wird einmal neu geladen.
#>
function Start-HUBackgroundLoad([Parameter(Mandatory)][string]$Name) {
    $l = $script:HULoads[$Name]
    if (-not $l) { throw "Lade-Kreislauf '$Name' ist nicht angemeldet" }
    if ($l.Job) { $l.Queued = $true; return }
    $argList = @()
    $ps = $null
    Set-HUCtlBusy $l.Dot $true
    try {
        if ($l.GetArguments) { $argList = @(& $l.GetArguments) }
        $ps = [powershell]::Create()
        [void]$ps.AddScript($l.Script)
        foreach ($a in $argList) { [void]$ps.AddArgument($a) }
        $l.Job = [pscustomobject]@{ Ps = $ps; Handle = $ps.BeginInvoke(); Start = Get-Date }
    } catch {
        if ($ps) { try { $ps.Dispose() } catch { } }
        $l.Job = $null
        Set-HUCtlBusy $l.Dot $false
        Write-HUCtlWarn "$($l.Prefix): Laden nicht gestartet: $(Get-HUCtlErrorText $_.Exception)"
    }
    $l.LastLoad = Get-Date
}
<#
.SYNOPSIS
    Ein Takt des Kreislaufs (vom Zeitgeber, ohne Zeitgeber selbst aufrufen): fertiges Laden abholen, Zeitueberschreitung
    pruefen, nach -RefreshSeconds neu laden. Ohne -Name alle angemeldeten.
#>
function Receive-HUBackgroundLoad([string]$Name = '') {
    if (-not $Name) { foreach ($n in @($script:HULoads.Keys)) { Receive-HUBackgroundLoad -Name $n }; return }
    $l = $script:HULoads[$Name]
    if (-not $l) { return }
    $j = $l.Job
    if (-not $j) {
        if ($l.Refresh -gt 0 -and ((Get-Date) - $l.LastLoad).TotalSeconds -ge $l.Refresh) { Start-HUBackgroundLoad -Name $Name }
        return
    }
    if (-not $j.Handle.IsCompleted) {
        if ($l.Timeout -gt 0 -and ((Get-Date) - $j.Start).TotalSeconds -gt $l.Timeout) {
            $l.Job = $null
            Stop-HUCtlPipeline $j.Ps
            Set-HUCtlBusy $l.Dot $false
            Write-HUCtlWarn "$($l.Prefix): Laden dauert zu lange (Freigabe erreichbar?)"
            if ($l.Failed) { try { & $l.Failed 'Zeitüberschreitung' } catch { } }
        }
        return
    }
    $l.Job = $null
    try {
        $res = @($j.Ps.EndInvoke($j.Handle))
        $err = @($j.Ps.Streams.Error)
        if (-not $res.Count) { throw $(if ($err.Count) { "$($err[0].Exception.Message)" } else { 'keine Daten' }) }
        & $l.Done $res[0]
    } catch {
        $m = Get-HUCtlErrorText $_.Exception
        Write-HUCtlWarn "$($l.Prefix): Laden fehlgeschlagen: $m"
        if ($l.Failed) { try { & $l.Failed $m } catch { } }
    } finally {
        try { $j.Ps.Dispose() } catch { }
        Set-HUCtlBusy $l.Dot $false
    }
    if ($l.Queued) { $l.Queued = $false; Start-HUBackgroundLoad -Name $Name }
}
<#
.SYNOPSIS
    Laeuft gerade ein Laden?
#>
function Test-HUBackgroundLoadRunning([Parameter(Mandatory)][string]$Name) {
    $l = $script:HULoads[$Name]
    return [bool]($l -and $l.Job)
}
<#
.SYNOPSIS
    Zeitgeber anhalten und ein laufendes Laden abbrechen (z. B. in Stop-HUAddon<Id>). Die Anmeldung bleibt.
#>
function Stop-HUBackgroundLoad([Parameter(Mandatory)][string]$Name) {
    $l = $script:HULoads[$Name]
    if (-not $l) { return }
    if ($l.Timer) { try { $l.Timer.Stop() } catch { } }
    $j = $l.Job
    $l.Job = $null; $l.Queued = $false
    if ($j) {
        Stop-HUCtlPipeline $j.Ps
        Set-HUCtlBusy $l.Dot $false
    }
}
<#
.SYNOPSIS
    Wie Stop-HUBackgroundLoad und dazu die Anmeldung entfernen.
#>
function Unregister-HUBackgroundLoad([Parameter(Mandatory)][string]$Name) {
    Stop-HUBackgroundLoad -Name $Name
    [void]$script:HULoads.Remove($Name)
}
