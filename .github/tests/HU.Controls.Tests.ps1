#Requires -Version 5.1
# Pester 5: Fenster-Bausteine, Dialoge und Hintergrund-Laden, Functions\UI-Controls.ps1.
# Keine sichtbaren Fenster: WPF-Objekte nur ohne Show/ShowDialog und nur im STA-Thread, Hintergrund-Auftraege echt
# (eigene Runspaces), abgeholt ohne Dispatcher ueber Receive-* mit kurzer Wartezeit.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:Sta = ([Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA')
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $script:AppRoot = $script:Root
    $script:App = [pscustomobject]@{ title = 'Test' }
    $script:Window = $null; $script:AppIcon = $null; $script:LastDialog = $null; $script:UiScale = 1.0
    . (Join-Path $script:Root 'Functions\UI-Common.ps1')
    . (Join-Path $script:Root 'Functions\UI-Update.ps1')

    # Stubs fuer Kern-Teile, die hier nicht geprueft werden (Zustand merken, Protokoll, gruener Punkt)
    $script:StateCalls = New-Object System.Collections.Generic.List[string]
    $script:Warns = New-Object System.Collections.Generic.List[string]
    $script:BusyCalls = New-Object System.Collections.Generic.List[string]
    function Restore-HUDialogState($Window, [string]$Key) { $script:StateCalls.Add("restore:$Key") }
    function Save-HUDialogState($Window, [string]$Key) { $script:StateCalls.Add("save:$Key") }
    function Write-HULogWarn([string]$Message) { $script:Warns.Add($Message) }
    function Set-HUBusyDot($Dot, [bool]$On) { $script:BusyCalls.Add("$Dot=$On") }

    . (Join-Path $script:Root 'Functions\UI-Controls.ps1')

    function Wait-HUTest([scriptblock]$Condition, [int]$TimeoutMs = 15000) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
            if (& $Condition) { return $true }
            Start-Sleep -Milliseconds 50
        }
        return $false
    }
    function New-HUTestOwner {
        $x = '<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Window.Resources><!--HU:THEME--></Window.Resources><Grid x:Name="root"/></Window>'
        return (New-HUWindow -XamlText (Get-HUThemedXaml $x)).Window
    }
    function Invoke-HUTestClick($Button) {
        try { $Button.RaiseEvent((New-Object System.Windows.RoutedEventArgs ([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))) } catch { }
    }
    $script:TmpRoot = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtCtl-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](New-Item -ItemType Directory -Path $script:TmpRoot)
}
AfterAll {
    try { Stop-HUBackgroundJobs -WaitStaSec 0 } catch { }
    foreach ($n in @($script:HULoads.Keys)) { try { Unregister-HUBackgroundLoad -Name $n } catch { } }
    Remove-Item -LiteralPath $script:TmpRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Auswahllisten' {
    It 'Eintrag mit Tag, Stil aus dem Thema, Tag setzen und lesen' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $o = New-HUTestOwner
        $cb = New-Object System.Windows.Controls.ComboBox
        foreach ($p in @(@('Alle', ''), @('Eins', 'a'), @('Zwei', 'b'))) { [void]$cb.Items.Add((New-HUComboItem $o $p[0] $p[1] -ToolTip "Tip $($p[0])")) }
        $cb.Items[1].Style | Should -Not -BeNullOrEmpty
        $cb.Items[1].ToolTip | Should -Be 'Tip Eins'
        Set-HUComboTag $cb 'b'
        Get-HUComboTag $cb | Should -Be 'b'
        Set-HUComboTag $cb 'gibtsnicht'
        $cb.SelectedIndex | Should -Be 0
        $cb.SelectedIndex = 2
        Set-HUComboTag $cb 'gibtsnicht' -NoFallback
        $cb.SelectedIndex | Should -Be 2
        $cb.SelectedIndex = -1
        Get-HUComboTag $cb | Should -Be ''
    }
    It 'ListBox-Eintrag mit Objekt als Tag' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $o = New-HUTestOwner
        $i = New-HUListItem $o 'Text' @{ Id = 7 } 'Hinweis'
        $i.Tag.Id | Should -Be 7
        $i.ToolTip | Should -Be 'Hinweis'
        $i.Style | Should -Not -BeNullOrEmpty
    }
}

Describe 'Kalender' {
    It 'Monatsnamen deutsch (Oesterreich)' {
        Get-HUMonthName 1 | Should -Be ([string][char]0x004A + [char]0x00E4 + 'nner')
        Get-HUMonthName 3 | Should -Be ('M' + [char]0x00E4 + 'rz')
        Get-HUMonthName 12 | Should -Be 'Dezember'
        Get-HUMonthName 0 | Should -Be ''
    }
    It 'Datum aus dem Feld lesen' {
        ConvertFrom-HUCalendarText '6.11.2026' | Should -Be (Get-Date -Year 2026 -Month 11 -Day 6).Date
        ConvertFrom-HUCalendarText '06.11.26' | Should -Be (Get-Date -Year 2026 -Month 11 -Day 6).Date
        ConvertFrom-HUCalendarText '2026-11-06' | Should -Be (Get-Date -Year 2026 -Month 11 -Day 6).Date
        ConvertFrom-HUCalendarText '2026-11-06T14:30:00' | Should -Be (Get-Date -Year 2026 -Month 11 -Day 6).Date
        ConvertFrom-HUCalendarText ([datetime]'2026-11-06 10:00') | Should -Be (Get-Date -Year 2026 -Month 11 -Day 6).Date
        ConvertFrom-HUCalendarText '' | Should -BeNullOrEmpty
        ConvertFrom-HUCalendarText '31.02.2026' | Should -BeNullOrEmpty
        ConvertFrom-HUCalendarText 'morgen' | Should -BeNullOrEmpty
    }
    It '42 Tage ab Montag, Markierungen' {
        $d = @(Get-HUCalendarDays -Month ([datetime]'2026-11-15') -Selected ([datetime]'2026-11-06') -Today ([datetime]'2026-11-07'))
        $d.Count | Should -Be 42
        $d[0].Date | Should -Be ([datetime]'2026-10-26')   # 1.11.2026 ist ein Sonntag
        $d[0].Date.DayOfWeek | Should -Be 'Monday'
        @($d | Where-Object InMonth).Count | Should -Be 30
        @($d | Where-Object IsSelected)[0].Day | Should -Be 6
        @($d | Where-Object IsToday)[0].Day | Should -Be 7
        @($d | Where-Object { $_.IsWeekend -and $_.InMonth }).Count | Should -Be 9
    }
    It 'Pfeile: Monat, mit Umschalt Jahr' {
        Get-HUCalendarStep -1 | Should -Be -1
        Get-HUCalendarStep 1 | Should -Be 1
        Get-HUCalendarStep -1 $true | Should -Be -12
        Get-HUCalendarStep 1 $true | Should -Be 12
    }
    It 'Popup bauen, Monat zeichnen, Tag eintragen, blaettern (ohne Oeffnen)' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $o = New-HUTestOwner
        $tb = New-Object System.Windows.Controls.TextBox; $btn = New-Object System.Windows.Controls.Button
        $tb.Text = '15.03.2027'
        $st = Add-HUCalendar -TextBox $tb -Button $btn -Owner $o
        $btn.Tag | Should -Be $st
        $btn.ToolTip | Should -Not -BeNullOrEmpty
        $tb.ToolTip | Should -Not -BeNullOrEmpty
        $st.Month = [datetime]'2027-03-01'
        Update-HUCalendarView $st
        $st.Title.Text | Should -Be ('M' + [char]0x00E4 + 'rz 2027')
        $st.Grid.Children.Count | Should -Be 49
        $days = @($st.Grid.Children | Where-Object { $_ -is [System.Windows.Controls.Button] })
        $days.Count | Should -Be 42
        @($days | Where-Object { -not $_.ToolTip }).Count | Should -Be 0
        $sel = @($days | Where-Object { "$($_.Background)" -eq '#FF007ACC' })
        $sel.Count | Should -Be 1
        $sel[0].Content | Should -Be '15'
        # Klick auf den 20. traegt ihn ein
        $d20 = @($days | Where-Object { ([datetime]$_.Tag) -eq [datetime]'2027-03-20' })[0]
        Invoke-HUTestClick $d20
        $tb.Text | Should -Be '20.03.2027'
        $st.Pop.IsOpen | Should -BeFalse
        # Pfeile (ohne Umschalt)
        $head = $st.Pop.Child.Child.Children[0]
        $prev = $head.Children[0]; $next = $head.Children[1]
        $prev.ToolTip | Should -Match 'Umschalt'
        Invoke-HUTestClick $prev
        $st.Month | Should -Be ([datetime]'2027-02-01')
        Invoke-HUTestClick $next; Invoke-HUTestClick $next
        $st.Month | Should -Be ([datetime]'2027-04-01')
    }
    It 'eigener Datumsleser und eigenes Format' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $tb = New-Object System.Windows.Controls.TextBox; $btn = New-Object System.Windows.Controls.Button
        $tb.Text = 'X'
        $st = Add-HUCalendar -TextBox $tb -Button $btn -Parse { param($t) if ($t -eq 'X') { [datetime]'2026-05-04' } } -Format 'yyyy-MM-dd' -NoYearStep
        $st.YearStep | Should -BeFalse
        $st.Month = [datetime]'2026-05-01'
        Update-HUCalendarView $st
        $days = @($st.Grid.Children | Where-Object { $_ -is [System.Windows.Controls.Button] })
        @($days | Where-Object { "$($_.Background)" -eq '#FF007ACC' })[0].Content | Should -Be '4'
        Invoke-HUTestClick @($days | Where-Object { ([datetime]$_.Tag) -eq [datetime]'2026-05-09' })[0]
        $tb.Text | Should -Be '2026-05-09'
    }
}

Describe 'Uhrzeit und Zahlenfeld' {
    It 'Uhrzeiten: Standard 15 Minuten ab 0 Uhr' {
        $t = @(Get-HUTimeSlots)
        $t.Count | Should -Be 96
        $t[0] | Should -Be '00:00'; $t[1] | Should -Be '00:15'; $t[-1] | Should -Be '23:45'
    }
    It 'Uhrzeiten: 30 Minuten ab 6 Uhr, nach Mitternacht weiter' {
        $t = @(Get-HUTimeSlots -StepMinutes 30 -StartHour 6)
        $t.Count | Should -Be 48
        $t[0] | Should -Be '06:00'; $t[36] | Should -Be '00:00'; $t[-1] | Should -Be '05:30'
    }
    It 'ComboBox fuellen' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $o = New-HUTestOwner
        $cb = New-Object System.Windows.Controls.ComboBox
        Initialize-HUTimeCombo $cb $o -StepMinutes 60
        $cb.Items.Count | Should -Be 24
        $cb.Items[8].Content | Should -Be '08:00'
        "$($cb.Items[8].Tag)" | Should -Be ''
        $cb.MaxDropDownHeight | Should -Be 260
        $cb.ToolTip | Should -Not -BeNullOrEmpty
    }
    It 'Zahl fuer ein Eingabefeld' {
        Format-HUNumberInput 12.5 | Should -Be '12,5'
        Format-HUNumberInput ([decimal]1234567.891) | Should -Be '1234567,89'
        Format-HUNumberInput 3.14159 -Digits 3 | Should -Be '3,142'
        Format-HUNumberInput 7 -Digits 0 | Should -Be '7'
        Format-HUNumberInput 0 | Should -Be '0'
        Format-HUNumberInput 0 -ZeroAsEmpty | Should -Be ''
        Format-HUNumberInput $null | Should -Be ''
        Format-HUNumberInput '' | Should -Be ''
        Format-HUNumberInput '12.5' | Should -Be '12,5'
        Format-HUNumberInput '1.234,56' | Should -Be '1234,56'
        Format-HUNumberInput 'abc' | Should -Be 'abc'
    }
}

Describe 'Dialoge (ohne Anzeigen)' {
    It 'XAML der Dialoge: ladbar, dunkel, Anfasser, Tooltips' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        foreach ($x in @($script:HUCtlTextXaml, $script:HUCtlChoiceXaml)) {
            $w = [System.Windows.Markup.XamlReader]::Parse((Get-HUThemedXaml $x))
            $w.ResizeMode | Should -Be 'CanResizeWithGrip'
            "$($w.Background)" | Should -Be '#FF1E1E1E'
            foreach ($m in [regex]::Matches($x, '<(Button|TextBox)\b[^>]*>')) { $m.Value | Should -Match 'ToolTip=' }
        }
    }
    It 'Textdialog einzeilig: Abbruch $null, leer nur mit -AllowEmpty, Text getrimmt' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $script:StateCalls.Clear()
        $st = New-HUTextDialog -Title 'Grund' -Prompt 'Bitte Grund' -Default 'alt'
        $st.Window.Title | Should -Be 'Grund'
        $st.C.txText.Text | Should -Be 'alt'
        $st.C.txPrompt.Text | Should -Be 'Bitte Grund'
        $st.C.txOk.IsDefault | Should -BeTrue
        $st.C.txCancel.IsCancel | Should -BeTrue
        $st.Window.SizeToContent | Should -Be 'Height'
        $script:StateCalls | Should -Contain 'restore:huFenster_Text'
        $st.Text | Should -BeNullOrEmpty
        $st.C.txText.Text = '   '
        Complete-HUTextDialog $st | Should -BeFalse
        $null -eq $st.Text | Should -BeTrue
        $st.C.txText.Text = '  neuer Grund  '
        Invoke-HUTestClick $st.C.txOk
        $st.Text | Should -Be 'neuer Grund'
        $e = New-HUTextDialog -Prompt 'leer erlaubt' -AllowEmpty -StateKey 'testFenster_Text' -MaxLength 50
        $e.C.txText.MaxLength | Should -Be 50
        $script:StateCalls | Should -Contain 'restore:testFenster_Text'
        Complete-HUTextDialog $e | Should -BeTrue
        $e.Text | Should -Be ''
    }
    It 'Textdialog mehrzeilig (auch -Mehrzeilig)' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $script:StateCalls.Clear()
        $st = New-HUTextDialog -Prompt 'Liste' -Mehrzeilig
        $st.Multiline | Should -BeTrue
        $st.C.txText.AcceptsReturn | Should -BeTrue
        $st.C.txText.MaxLength | Should -Be 20000
        $st.C.txOk.IsDefault | Should -BeFalse
        $script:StateCalls | Should -Contain 'restore:huFenster_TextListe'
        $st.C.txText.Text = "a`r`nb`r`n"
        [void](Complete-HUTextDialog $st)
        $st.Text | Should -Be "a`r`nb"
    }
    It 'Auswahldialog: Knoepfe, Tooltips, Index, Abbrechen -1' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $st = New-HUChoiceDialog -Title 'Frage' -Prompt 'Was tun?' -Choices 'Nur diesen', 'Alle', 'Keinen' -ToolTips 'nur den einen', '', ''
        $st.C.chButtons.Children.Count | Should -Be 4
        $st.Buttons[0].ToolTip | Should -Be 'nur den einen'
        $st.Buttons[1].ToolTip | Should -Be 'Alle'
        "$($st.Buttons[0].Background)" | Should -Be '#FF1976D2'
        @($st.C.chButtons.Children | Where-Object { -not $_.ToolTip }).Count | Should -Be 0
        $st.Index | Should -Be -1
        Invoke-HUTestClick $st.Buttons[1]
        $st.Index | Should -Be 1
        $cancel = $st.C.chButtons.Children[3]
        $cancel.IsCancel | Should -BeTrue
        Invoke-HUTestClick $cancel
        $st.Index | Should -Be -1
        $n = New-HUChoiceDialog -Prompt 'ohne Abbrechen' -Choices 'Ja', 'Nein' -NoCancel
        $n.C.chButtons.Children.Count | Should -Be 2
    }
}

Describe 'Modul-Fenster' {
    BeforeAll {
        $script:ModDir = Join-Path $script:TmpRoot 'Probemodul'
        [void](New-Item -ItemType Directory -Path $script:ModDir)
        $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="400" Height="300" ResizeMode="CanResize" Background="#1E1E1E">
    <Window.Resources><!--HU:THEME--></Window.Resources>
    <Grid><TextBox x:Name="pmText" Style="{StaticResource DarkTextBox}" ToolTip="Probe"/></Grid>
</Window>
'@
        [IO.File]::WriteAllText((Join-Path $script:ModDir 'Probe.xaml'), $x, (New-Object Text.UTF8Encoding $true))
    }
    It 'laedt XAML mit Thema, merkt den Zustand unter Vorsatz + Dateiname, Anfasser' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $script:StateCalls.Clear()
        $d = New-HUModuleWindow -ModuleDir $script:ModDir -File 'Probe.xaml' -StatePrefix 'probeFenster_'
        $d.StateKey | Should -Be 'probeFenster_Probe'
        $d.Window.Resources['huStateKey'] | Should -Be 'probeFenster_Probe'
        $d.C.ContainsKey('pmText') | Should -BeTrue
        $d.C.pmText.Style | Should -Not -BeNullOrEmpty
        $d.Window.ResizeMode | Should -Be 'CanResizeWithGrip'
        $script:StateCalls | Should -Contain 'restore:probeFenster_Probe'
    }
    It 'ohne Vorsatz: Ordnername + Fenster_, ganzer Schluessel ueber -StateKey' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        (New-HUModuleWindow -ModuleDir $script:ModDir -File 'Probe.xaml').StateKey | Should -Be 'ProbemodulFenster_Probe'
        (New-HUModuleWindow -ModuleDir $script:ModDir -File 'Probe.xaml' -StateKey 'eigener').StateKey | Should -Be 'eigener'
    }
}

Describe 'Hintergrund-Auftraege' {
    BeforeEach { $script:Got = $null; $script:Warns.Clear() }
    AfterEach { Stop-HUBackgroundJobs -WaitStaSec 0 }

    It 'Ergebnis-Objekt, Kontext und leerer Fehlertext an Done' {
        Start-HUBackgroundJob -NoTimer -Script { param($a) [pscustomobject]@{ X = $a * 2 } } -Arguments @(21) -Context 'k1' -Done {
            param($Res, $Err, $Ctx) $script:Got = @{ R = $Res; E = $Err; C = $Ctx }
        }
        Get-HUBackgroundJobCount | Should -Be 1
        Wait-HUTest { Receive-HUBackgroundJobs; $null -ne $script:Got } | Should -BeTrue
        $script:Got.R.X | Should -Be 42
        $script:Got.E | Should -Be ''
        $script:Got.C | Should -Be 'k1'
        Get-HUBackgroundJobCount | Should -Be 0
    }
    It '-AllResults liefert alle Objekte, -Ctx als Alias' {
        Start-HUBackgroundJob -NoTimer -AllResults -Script { 1; 2; 3 } -Ctx 'k2' -Done { param($Res, $Err, $Ctx) $script:Got = @{ R = $Res; C = $Ctx } }
        Wait-HUTest { Receive-HUBackgroundJobs; $null -ne $script:Got } | Should -BeTrue
        @($script:Got.R).Count | Should -Be 3
        $script:Got.C | Should -Be 'k2'
    }
    It '-Sta laeuft in einem STA-Runspace' {
        Start-HUBackgroundJob -NoTimer -Sta -Script { "$([Threading.Thread]::CurrentThread.GetApartmentState())" } -Done { param($Res, $Err, $Ctx) $script:Got = @{ R = $Res } }
        Wait-HUTest { Receive-HUBackgroundJobs; $null -ne $script:Got } | Should -BeTrue
        $script:Got.R | Should -Be 'STA'
    }
    It 'Ausnahme und Fehler ohne Ausgabe werden zum Fehlertext' {
        Start-HUBackgroundJob -NoTimer -Script { throw 'kaputt' } -Done { param($Res, $Err, $Ctx) $script:Got = @{ R = $Res; E = $Err } }
        Wait-HUTest { Receive-HUBackgroundJobs; $null -ne $script:Got } | Should -BeTrue
        $script:Got.R | Should -BeNullOrEmpty
        $script:Got.E | Should -Be 'kaputt'
        $script:Got = $null
        Start-HUBackgroundJob -NoTimer -Script { Write-Error 'nur Fehler' } -Done { param($Res, $Err, $Ctx) $script:Got = @{ E = $Err } }
        Wait-HUTest { Receive-HUBackgroundJobs; $null -ne $script:Got } | Should -BeTrue
        $script:Got.E | Should -Be 'nur Fehler'
    }
    It 'Zeitlimit: Abbruch und Done mit Zeitueberschreitung' {
        Start-HUBackgroundJob -NoTimer -TimeoutSec 1 -Script { Start-Sleep -Seconds 20; 'zu spaet' } -Done { param($Res, $Err, $Ctx) $script:Got = @{ R = $Res; E = $Err } }
        Wait-HUTest { Receive-HUBackgroundJobs; $null -ne $script:Got } 8000 | Should -BeTrue
        $script:Got.R | Should -BeNullOrEmpty
        $script:Got.E | Should -Be ('Zeit' + [char]0x00FC + 'berschreitung')
        Get-HUBackgroundJobCount | Should -Be 0
    }
    It 'Fehler im Done landet im Protokoll mit Vorsatz' {
        Start-HUBackgroundJob -NoTimer -LogPrefix 'Probe' -Script { 1 } -Done { param($Res, $Err, $Ctx) $script:Got = 1; throw 'Done kaputt' }
        Wait-HUTest { Receive-HUBackgroundJobs; $null -ne $script:Got } | Should -BeTrue
        $script:Warns | Should -Contain 'Probe: Done kaputt'
    }
    It 'Stop je Gruppe beendet ohne Done' {
        Start-HUBackgroundJob -NoTimer -Owner 'A' -Script { Start-Sleep -Seconds 30 } -Done { $script:Got = 'A' }
        Start-HUBackgroundJob -NoTimer -Owner 'B' -Script { Start-Sleep -Seconds 30 } -Done { $script:Got = 'B' }
        $j = Start-HUBackgroundJob -NoTimer -Owner 'B' -Script { 1 } -PassThru
        $j.Owner | Should -Be 'B'
        Get-HUBackgroundJobCount -Owner 'A' | Should -Be 1
        Stop-HUBackgroundJobs -Owner 'A' -WaitStaSec 0
        Get-HUBackgroundJobCount -Owner 'A' | Should -Be 0
        Get-HUBackgroundJobCount -Owner 'B' | Should -Be 2
        Stop-HUBackgroundJobs -WaitStaSec 0
        Get-HUBackgroundJobCount | Should -Be 0
        $script:Got | Should -BeNullOrEmpty
    }
}

Describe 'Lade-Kreislauf' {
    BeforeEach { $script:Loaded = New-Object System.Collections.Generic.List[object]; $script:Failures = New-Object System.Collections.Generic.List[string]; $script:Warns.Clear(); $script:BusyCalls.Clear(); $script:N = 0 }
    AfterEach { foreach ($n in @($script:HULoads.Keys)) { Unregister-HUBackgroundLoad -Name $n } }

    It 'laedt im Hintergrund, Argumente frisch, Done mit Ergebnis, gruener Punkt an und aus' {
        Register-HUBackgroundLoad -Name 'probe' -NoTimer -RefreshSeconds 0 -Dot 'punkt' -Script { param($Mod, $Users) "$Mod|$(@($Users).Count)" } `
            -GetArguments { 'mod'; , [string[]]@('a', 'b') } -Done { param($Res) $script:Loaded.Add($Res) }
        Test-HUBackgroundLoadRunning 'probe' | Should -BeFalse
        Start-HUBackgroundLoad -Name 'probe'
        Test-HUBackgroundLoadRunning 'probe' | Should -BeTrue
        $script:BusyCalls | Should -Contain 'punkt=True'
        Wait-HUTest { Receive-HUBackgroundLoad -Name 'probe'; $script:Loaded.Count -ge 1 } | Should -BeTrue
        $script:Loaded[0] | Should -Be 'mod|2'
        $script:BusyCalls[-1] | Should -Be 'punkt=False'
        Test-HUBackgroundLoadRunning 'probe' | Should -BeFalse
    }
    It 'Start waehrend des Ladens: vorgemerkt, danach einmal neu' {
        Register-HUBackgroundLoad -Name 'warte' -NoTimer -RefreshSeconds 0 -Script { param($N) Start-Sleep -Milliseconds 300; $N } `
            -GetArguments { $script:N++; $script:N } -Done { param($Res) $script:Loaded.Add($Res) }
        Start-HUBackgroundLoad -Name 'warte'
        Start-HUBackgroundLoad -Name 'warte'
        Start-HUBackgroundLoad -Name 'warte'
        Wait-HUTest { Receive-HUBackgroundLoad; $script:Loaded.Count -ge 2 -and -not (Test-HUBackgroundLoadRunning 'warte') } | Should -BeTrue
        Start-Sleep -Milliseconds 400; Receive-HUBackgroundLoad
        $script:Loaded.Count | Should -Be 2
        "$($script:Loaded[0])|$($script:Loaded[1])" | Should -Be '1|2'
    }
    It 'Abgleich nach RefreshSeconds' {
        Register-HUBackgroundLoad -Name 'takt' -NoTimer -RefreshSeconds 1 -Script { 'x' } -Done { param($Res) $script:Loaded.Add($Res) } -Start
        Wait-HUTest { Receive-HUBackgroundLoad -Name 'takt'; $script:Loaded.Count -ge 2 } 8000 | Should -BeTrue
    }
    It 'Fehler: Warnung, Failed, kein Done' {
        Register-HUBackgroundLoad -Name 'fehler' -NoTimer -RefreshSeconds 0 -LogPrefix 'Probe' -Script { throw 'weg' } `
            -Done { param($Res) $script:Loaded.Add($Res) } -Failed { param($Err) $script:Failures.Add($Err) } -Start
        Wait-HUTest { Receive-HUBackgroundLoad -Name 'fehler'; $script:Failures.Count -ge 1 } | Should -BeTrue
        $script:Loaded.Count | Should -Be 0
        $script:Failures[0] | Should -Be 'weg'
        $script:Warns | Should -Contain 'Probe: Laden fehlgeschlagen: weg'
        $script:Failures.Clear()
        Register-HUBackgroundLoad -Name 'leer' -NoTimer -RefreshSeconds 0 -Script { } -Done { param($Res) $script:Loaded.Add($Res) } -Failed { param($Err) $script:Failures.Add($Err) } -Start
        Wait-HUTest { Receive-HUBackgroundLoad -Name 'leer'; $script:Failures.Count -ge 1 } | Should -BeTrue
        $script:Failures[0] | Should -Be 'keine Daten'
    }
    It 'Zeitueberschreitung beim Laden' {
        Register-HUBackgroundLoad -Name 'lang' -NoTimer -RefreshSeconds 0 -TimeoutSec 1 -Dot 'p' -Script { Start-Sleep -Seconds 20; 'x' } `
            -Done { param($Res) $script:Loaded.Add($Res) } -Failed { param($Err) $script:Failures.Add($Err) } -Start
        Wait-HUTest { Receive-HUBackgroundLoad -Name 'lang'; $script:Failures.Count -ge 1 } 8000 | Should -BeTrue
        $script:Failures[0] | Should -Be ('Zeit' + [char]0x00FC + 'berschreitung')
        @($script:Warns | Where-Object { $_ -like 'lang: Laden dauert zu lange*' }).Count | Should -Be 1
        $script:BusyCalls[-1] | Should -Be 'p=False'
        Test-HUBackgroundLoadRunning 'lang' | Should -BeFalse
    }
    It 'Stop bricht ab, Unregister entfernt, Zeitgeber wird angelegt' {
        Register-HUBackgroundLoad -Name 'stop' -RefreshSeconds 0 -Script { Start-Sleep -Seconds 20; 'x' } -Done { param($Res) $script:Loaded.Add($Res) } -Start
        $script:HULoads['stop'].Timer | Should -Not -BeNullOrEmpty
        $script:HULoads['stop'].Timer.Tag | Should -Be 'stop'
        Test-HUBackgroundLoadRunning 'stop' | Should -BeTrue
        Stop-HUBackgroundLoad -Name 'stop'
        Test-HUBackgroundLoadRunning 'stop' | Should -BeFalse
        $script:HULoads['stop'].Timer.IsEnabled | Should -BeFalse
        Unregister-HUBackgroundLoad -Name 'stop'
        $script:HULoads.ContainsKey('stop') | Should -BeFalse
        { Start-HUBackgroundLoad -Name 'stop' } | Should -Throw
    }
}

Describe 'Abbrechen ohne Warten (Pruefung 086, KERN-B5)' {
    BeforeEach { $script:Got = $null }
    AfterEach {
        Stop-HUBackgroundJobs -WaitStaSec 0
        foreach ($n in @($script:HULoads.Keys)) { Unregister-HUBackgroundLoad -Name $n }
        [void](Wait-HUTest { Complete-HUCtlStops; $script:HUCtlStopping.Count -eq 0 } 15000)
    }

    It 'Auftrag mit Start-Sleep 30 laeuft in die Zeitueberschreitung, Abholen kehrt sofort zurueck' {
        Start-HUBackgroundJob -NoTimer -TimeoutSec 1 -Script { Start-Sleep -Seconds 30; 'zu spaet' } -Done { param($Res, $Err, $Ctx) $script:Got = @{ R = $Res; E = $Err } }
        Start-Sleep -Milliseconds 1300
        $sw = [Diagnostics.Stopwatch]::StartNew(); Receive-HUBackgroundJobs; $sw.Stop()
        $script:Got.E | Should -Be 'Zeitüberschreitung'
        $sw.ElapsedMilliseconds | Should -BeLessThan 1000
        Get-HUBackgroundJobCount | Should -Be 0
        Wait-HUTest { Complete-HUCtlStops; $script:HUCtlStopping.Count -eq 0 } 10000 | Should -BeTrue
    }
    It 'Auftrag, der den Abbruch erst spaet annimmt (.NET-Aufruf), haelt den Fenster-Thread nicht auf' {
        Start-HUBackgroundJob -NoTimer -TimeoutSec 1 -Script { [Threading.Thread]::Sleep(4000); 'x' } -Done { param($Res, $Err, $Ctx) $script:Got = @{ E = $Err } }
        Start-Sleep -Milliseconds 1300
        $sw = [Diagnostics.Stopwatch]::StartNew(); Receive-HUBackgroundJobs; $sw.Stop()
        $sw.ElapsedMilliseconds | Should -BeLessThan 1000
        $script:Got.E | Should -Be 'Zeitüberschreitung'
        $script:HUCtlStopping.Count | Should -Be 1   # Stop laeuft noch; freigegeben wird im naechsten Takt
        Wait-HUTest { Complete-HUCtlStops; $script:HUCtlStopping.Count -eq 0 } 10000 | Should -BeTrue
    }
    It 'Stop-HUBackgroundJobs kehrt sofort zurueck' {
        Start-HUBackgroundJob -NoTimer -Owner 'B5' -Script { [Threading.Thread]::Sleep(3000) }
        Start-Sleep -Milliseconds 300
        $sw = [Diagnostics.Stopwatch]::StartNew(); Stop-HUBackgroundJobs -Owner 'B5' -WaitStaSec 0; $sw.Stop()
        $sw.ElapsedMilliseconds | Should -BeLessThan 1000
        Get-HUBackgroundJobCount -Owner 'B5' | Should -Be 0
    }
    It 'Lade-Kreislauf: Zeitueberschreitung und Stop-HUBackgroundLoad kehren sofort zurueck' {
        Register-HUBackgroundLoad -Name 'haengt' -NoTimer -RefreshSeconds 0 -TimeoutSec 1 -Script { [Threading.Thread]::Sleep(4000); 'x' } `
            -Done { param($Res) } -Failed { param($Err) $script:Got = $Err } -Start
        Start-Sleep -Milliseconds 1300
        $sw = [Diagnostics.Stopwatch]::StartNew(); Receive-HUBackgroundLoad -Name 'haengt'; $sw.Stop()
        $sw.ElapsedMilliseconds | Should -BeLessThan 1000
        $script:Got | Should -Be 'Zeitüberschreitung'
        Register-HUBackgroundLoad -Name 'stopp' -NoTimer -RefreshSeconds 0 -Script { [Threading.Thread]::Sleep(3000); 'x' } -Done { param($Res) } -Start
        Start-Sleep -Milliseconds 300
        $sw = [Diagnostics.Stopwatch]::StartNew(); Stop-HUBackgroundLoad -Name 'stopp'; $sw.Stop()
        $sw.ElapsedMilliseconds | Should -BeLessThan 1000
        Test-HUBackgroundLoadRunning 'stopp' | Should -BeFalse
        Wait-HUTest { Complete-HUCtlStops; $script:HUCtlStopping.Count -eq 0 } 10000 | Should -BeTrue
    }
    It 'kein synchrones PowerShell.Stop() mehr in den Bausteinen' {
        $src = [IO.File]::ReadAllText((Join-Path $script:Root 'Functions\UI-Controls.ps1'), [Text.Encoding]::UTF8)
        $src | Should -Not -Match '\.Ps\.Stop\(\)'
        $src | Should -Match 'BeginStop\(\$null, \$null\)'
    }
}

Describe 'Auswahl mit Standardknopf (Entscheidung Read-HUChoice -DefaultIndex)' {
    It 'genau der Knopf mit -DefaultIndex ist IsDefault, ohne Angabe keiner' {
        if (-not $script:Sta) { Set-ItResult -Skipped -Because 'WPF braucht STA'; return }
        $st = New-HUChoiceDialog -Title 'Probe' -Prompt 'Welche?' -Choices @('A', 'B', 'C') -DefaultIndex 1
        try { @($st.Buttons | ForEach-Object { $_.IsDefault }) | Should -Be @($false, $true, $false) } finally { $st.Window.Close() }
        $st2 = New-HUChoiceDialog -Title 'Probe' -Prompt 'Welche?' -Choices @('A', 'B')
        try { @($st2.Buttons | Where-Object { $_.IsDefault }).Count | Should -Be 0 } finally { $st2.Window.Close() }
    }
    It 'Read-HUChoice reicht -DefaultIndex weiter' {
        (Get-Command Read-HUChoice).Parameters.ContainsKey('DefaultIndex') | Should -BeTrue
        (Get-Command New-HUChoiceDialog).Parameters['DefaultIndex'].ParameterType | Should -Be ([int])
    }
}

Describe 'Betragsfeld mit zweideutiger Eingabe (Entscheidung zu 1.500)' {
    BeforeAll { Import-Module (Join-Path $script:Root 'Core\HU.Data.psm1') -Force -DisableNameChecking }
    BeforeEach { $script:Asked = New-Object System.Collections.Generic.List[string]; $script:Answer = -1 }
    It 'fragt bei 1.500 einmal nach und liefert die gewaehlte Zahl' {
        function Read-HUChoice { param([string]$Title, [string]$Prompt, [string[]]$Choices, $Owner, [string[]]$ToolTips, [int]$DefaultIndex = -1) $script:Asked.Add("$Prompt|$($Choices -join '/')|$DefaultIndex"); $script:Answer }
        $script:Answer = 0
        Read-HUNumberInput '1.500' | Should -Be ([decimal]1500)
        $script:Answer = 1
        Read-HUNumberInput '1.500' | Should -Be ([decimal]1.5)
        $script:Answer = -1
        Read-HUNumberInput '1.500' | Should -BeNullOrEmpty
        $script:Asked.Count | Should -Be 3
        $script:Asked[0] | Should -Be 'Meinst du 1.500 (also 1500) oder 1,5?|1.500/1,5|0'
    }
    It 'fragt bei eindeutigen Eingaben nicht' {
        function Read-HUChoice { param([string]$Title, [string]$Prompt, [string[]]$Choices, $Owner, [string[]]$ToolTips, [int]$DefaultIndex = -1) $script:Asked.Add($Prompt); 0 }
        Read-HUNumberInput '1.500,00' | Should -Be ([decimal]1500)
        Read-HUNumberInput '1,5' | Should -Be ([decimal]1.5)
        Read-HUNumberInput '12.5' | Should -Be ([decimal]12.5)
        Read-HUNumberInput '1500' | Should -Be ([decimal]1500)
        Read-HUNumberInput '' | Should -BeNullOrEmpty
        Read-HUNumberInput 'abc' | Should -BeNullOrEmpty
        $script:Asked.Count | Should -Be 0
    }
    It 'mit Format-HUNumberInput danach ist das Feld eindeutig (keine zweite Rueckfrage)' {
        Test-HUAmbiguousNumber (Format-HUNumberInput ([decimal]1500)) | Should -BeFalse
    }
}

Describe 'Kalender: zweistellige Jahre (Pruefung 086, KERN-B6)' {
    It '1.1.35 ist 2035, 1.1.50 ist 1950' {
        (ConvertFrom-HUCalendarText '1.1.35').Year | Should -Be 2035
        (ConvertFrom-HUCalendarText '31.12.49').Year | Should -Be 2049
        (ConvertFrom-HUCalendarText '1.1.50').Year | Should -Be 1950
        $script:HUCtlDateCulture.DateTimeFormat.Calendar.TwoDigitYearMax | Should -Be 2049
    }
}

Describe 'Meldung mit Symbol Information (Pruefung 086, KERN-U1)' {
    It 'Show-HUMessage nimmt Information als anderen Namen fuer Info' {
        $script:Shown = New-Object System.Collections.Generic.List[string]
        function Show-HUDialogBox { param([string]$Text, [string]$Title, [string]$Mode, [string]$Icon, $Owner) $script:Shown.Add($Icon); 'OK' }
        { Show-HUMessage 'Probe' -Icon Information } | Should -Not -Throw
        Show-HUMessage 'Probe' -Icon Info
        @($script:Shown) | Should -Be @('Info', 'Info')
    }
    It 'Kern-Aufrufe nutzen -Icon Info' {
        $hits = @(Get-ChildItem -LiteralPath (Join-Path $script:Root 'Functions') -Filter *.ps1 | Select-String -Pattern '-Icon Information\b' | ForEach-Object { "$($_.Filename):$($_.LineNumber)" })
        $hits | Should -BeNullOrEmpty
    }
}

Describe 'JSON schreiben mit eindeutiger Temp-Datei (Pruefung 086, KERN-B10)' {
    BeforeEach {
        $script:B10 = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtB10-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:B10)
    }
    AfterEach { Remove-Item -LiteralPath $script:B10 -Recurse -Force -ErrorAction SilentlyContinue }
    It 'schreibt ohne BOM, ersetzt, legt .bak an und laesst keine Temp-Datei liegen' {
        $f = Join-Path $script:B10 'probe.json'
        Write-HUJsonFile -Path $f -Object ([pscustomobject]@{ A = 1 })
        Write-HUJsonFile -Path $f -Object ([pscustomobject]@{ A = 2 }) -Backup
        ([IO.File]::ReadAllText($f) | ConvertFrom-Json).A | Should -Be 2
        ([IO.File]::ReadAllText("$f.bak") | ConvertFrom-Json).A | Should -Be 1
        [IO.File]::ReadAllBytes($f)[0] | Should -Be ([byte][char]'{')
        @(Get-ChildItem -LiteralPath $script:B10 -Force | Where-Object { $_.Name -like '*.tmp' }).Count | Should -Be 0
    }
    It 'eine fremde probe.json.tmp (alter fester Name) wird nicht benutzt und bleibt' {
        $f = Join-Path $script:B10 'probe.json'
        [IO.File]::WriteAllText("$f.tmp", 'fremd')
        Write-HUJsonFile -Path $f -Object ([pscustomobject]@{ A = 3 })
        [IO.File]::ReadAllText("$f.tmp") | Should -Be 'fremd'
        ([IO.File]::ReadAllText($f) | ConvertFrom-Json).A | Should -Be 3
    }
}
