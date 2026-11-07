#Requires -Version 5.1
<#
.SYNOPSIS
    Gefuehrter Test fuer Module (nur im Testmodus): Schritte mit Aktion und Soll, automatische Pruefung und Bestaetigung
    durch einen Menschen streng getrennt, Protokoll als CSV und HTML.
.DESCRIPTION
    Ein Modul liefert nur seine Szenarien:
        Show-HUTestAssistant -ModuleId 'Buchung' -ModuleVersion '0.1.1' -Scenarios @(
            @{ Name = 'Heimatsaal mit Teilflaeche und Rabatt'; Steps = @(
                @{ Action = 'Neue Buchung: Saal 1/2 am ...'; Expected = 'Wird gespeichert, Preis ...'; Check = { ... ; @{ Ok = $true; Text = '1 Buchung gefunden' } } }
                @{ Action = 'Zweite Buchung Saal 2/3 zur selben Zeit'; Expected = 'Wird abgelehnt (Teilflaechen ueberschneiden sich)' }
            ) }
        )
    Check ist optional. Er gibt $true/$false oder @{ Ok = ...; Text = ... } zurueck. Ein Fehler darin gilt als "Fehler".
    Das automatische Ergebnis setzt nie die Bestaetigung des Menschen und umgekehrt.
    Der Testende bedient die echten Masken des Moduls (gegen den Testbestand); das Fenster bleibt daneben offen.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7.
#>

$script:HUTaNone = [string][char]0x2013   # Gedankenstrich: nicht geprueft

# ---------------------------------------------------------------- Logik (ohne Oberflaeche, testbar) ----------------------------------------------------------------

<#
.SYNOPSIS
    Testlauf aus Szenarien anlegen. Rueckgabe: Objekt mit Kopfdaten und flacher Schrittliste.
#>
function New-HUTestRun {
    param([Parameter(Mandatory)][string]$ModuleId, [string]$ModuleVersion = '', [Parameter(Mandatory)][object[]]$Scenarios,
        [string]$AppVersion = "$($script:Version)", [string]$UserName = $env:USERNAME, [string]$PcName = $env:COMPUTERNAME)
    $steps = New-Object System.Collections.Generic.List[object]
    $si = 0
    foreach ($sc in @($Scenarios)) {
        $si++
        $name = "$($sc.Name)"; if (-not $name) { $name = "Szenario $si" }
        $n = 0
        foreach ($st in @($sc.Steps)) {
            $n++
            $chk = $st.Check
            $steps.Add([pscustomobject]@{
                    Szenario = $name; Nr = $n; Aktion = "$($st.Action)"; Soll = "$($st.Expected)"
                    HatPruefung = ($chk -is [scriptblock]); Pruefung = $chk
                    Automatisch = $script:HUTaNone; AutomatischText = ''
                    Mensch = 'offen'; Kommentar = ''; Zeit = ''
                })
        }
    }
    if ($steps.Count -eq 0) { throw 'Keine Testschritte angegeben.' }
    return [pscustomobject]@{
        ModuleId = $ModuleId; ModuleVersion = $ModuleVersion; AppVersion = $AppVersion
        Benutzer = $UserName; Pc = $PcName; Beginn = (Get-Date); Steps = $steps.ToArray()   # Array: @() ueber List[object] wirft
    }
}

<#
.SYNOPSIS
    Automatische Pruefung eines Schritts ausfuehren. Setzt nur Automatisch/AutomatischText, nie die Bestaetigung des Menschen.
#>
function Invoke-HUTestCheck {
    param([Parameter(Mandatory)]$Step)
    if (-not $Step.HatPruefung) { $Step.Automatisch = $script:HUTaNone; $Step.AutomatischText = 'keine automatische Prüfung'; return $Step }
    try {
        $r = & $Step.Pruefung
        $r = @($r) | Select-Object -Last 1
        $ok = $false; $txt = ''
        if ($r -is [bool]) { $ok = $r }
        elseif ($r -is [hashtable]) { $ok = [bool]$r['Ok']; $txt = "$($r['Text'])" }
        elseif ($null -ne $r -and $r.PSObject.Properties['Ok']) { $ok = [bool]$r.Ok; if ($r.PSObject.Properties['Text']) { $txt = "$($r.Text)" } }
        else { $ok = $false; $txt = 'Prüfung lieferte kein Ergebnis' }
        $Step.Automatisch = $(if ($ok) { 'OK' } else { 'Fehler' })
        $Step.AutomatischText = $txt
    } catch {
        $Step.Automatisch = 'Fehler'
        $Step.AutomatischText = "Prüfung abgebrochen: $($_.Exception.Message)"
    }
    return $Step
}

<#
.SYNOPSIS
    Bestaetigung des Menschen setzen: Bestanden oder Fehler (Fehler nur mit Kommentar). Aendert nie das automatische Ergebnis.
#>
function Set-HUTestVerdict {
    param([Parameter(Mandatory)]$Step, [Parameter(Mandatory)][ValidateSet('Bestanden', 'Fehler', 'offen')][string]$Verdict, [string]$Comment = '')
    if ($Verdict -eq 'Fehler' -and -not "$Comment".Trim()) { throw 'Bei „Fehler“ bitte kurz beschreiben, was passiert ist.' }
    $Step.Mensch = $Verdict
    $Step.Kommentar = "$Comment".Trim()
    $Step.Zeit = $(if ($Verdict -eq 'offen') { '' } else { (Get-Date).ToString('dd.MM.yyyy HH:mm:ss') })
    return $Step
}

<#
.SYNOPSIS
    Zusammenfassung: Zahl der Schritte, bestaetigt, offen und die offenen Fehler (automatisch oder vom Menschen).
#>
function Get-HUTestSummary {
    param([Parameter(Mandatory)]$Run)
    $s = [object[]]$Run.Steps
    $fehler = @($s | Where-Object { $_.Mensch -eq 'Fehler' -or $_.Automatisch -eq 'Fehler' })
    return [pscustomobject]@{
        Schritte = $s.Count
        Bestanden = @($s | Where-Object { $_.Mensch -eq 'Bestanden' }).Count
        FehlerMensch = @($s | Where-Object { $_.Mensch -eq 'Fehler' }).Count
        Offen = @($s | Where-Object { $_.Mensch -eq 'offen' }).Count
        AutoOk = @($s | Where-Object { $_.Automatisch -eq 'OK' }).Count
        AutoFehler = @($s | Where-Object { $_.Automatisch -eq 'Fehler' }).Count
        OffeneFehler = $fehler
    }
}

<#
.SYNOPSIS
    Protokoll speichern: CSV (Excel-tauglich) und HTML (zum Ansehen und Drucken). Rueckgabe: Csv, Html.
    Automatische Ergebnisse und Bestaetigungen des Menschen stehen in getrennten Spalten.
#>
function Export-HUTestProtocol {
    param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][string]$Folder, [datetime]$Now = (Get-Date))
    if (-not (Test-Path -LiteralPath $Folder)) { [void](New-Item -ItemType Directory -Path $Folder -Force) }
    $base = 'Testprotokoll_{0}_{1}' -f $Run.ModuleId, $Now.ToString('yyyyMMdd-HHmmss')
    $rows = @(foreach ($s in $Run.Steps) {
            [pscustomobject]@{
                App = $Run.AppVersion; Modul = $Run.ModuleId; Modulversion = $Run.ModuleVersion; Szenario = $s.Szenario; Schritt = $s.Nr
                Aktion = $s.Aktion; Soll = $s.Soll
                AutomatischErgebnis = $s.Automatisch; AutomatischIst = $s.AutomatischText
                MenschErgebnis = $s.Mensch; MenschIstKommentar = $s.Kommentar; Zeit = $s.Zeit; Benutzer = $Run.Benutzer; Pc = $Run.Pc
            }
        })
    $cols = @(
        @{ Name = 'App-Version'; Source = 'App' }, @{ Name = 'Modul' }, @{ Name = 'Modulversion' }, @{ Name = 'Szenario' }, @{ Name = 'Schritt'; Type = 'number' },
        @{ Name = 'Aktion' }, @{ Name = 'Soll' }, @{ Name = 'Automatisch geprüft'; Source = 'AutomatischErgebnis' }, @{ Name = 'Automatisch Ist'; Source = 'AutomatischIst' },
        @{ Name = 'Bestätigung Mensch'; Source = 'MenschErgebnis' }, @{ Name = 'Ist laut Mensch / Kommentar'; Source = 'MenschIstKommentar' },
        @{ Name = 'Zeit' }, @{ Name = 'Benutzer' }, @{ Name = 'PC'; Source = 'Pc' })
    $csv = Export-HUTable -Rows $rows -TableProfile (New-HUTableProfile -Columns $cols) -Path (Join-Path $Folder "$base.csv")
    $sum = Get-HUTestSummary $Run
    $e = { param($t) [System.Net.WebUtility]::HtmlEncode("$t") }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><title>' + (& $e $base) + '</title><style>')
    [void]$sb.Append('body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:24px;color:#222}h1{font-size:20px}table{border-collapse:collapse;width:100%}')
    [void]$sb.Append('th,td{border:1px solid #bbb;padding:4px 6px;vertical-align:top;text-align:left}th{background:#eee}.auto{color:#555;background:#f6f6f6}')
    [void]$sb.Append('.ok{color:#1b5e20;font-weight:bold}.err{color:#b71c1c;font-weight:bold}.box{border:2px solid #e65100;padding:8px;margin:10px 0}</style></head><body>')
    [void]$sb.Append('<h1>Testprotokoll ' + (& $e $Run.ModuleId) + '</h1>')
    [void]$sb.Append('<p>App ' + (& $e $Run.AppVersion) + ' &middot; Modul ' + (& $e $Run.ModuleId) + ' ' + (& $e $Run.ModuleVersion) + ' &middot; ' + (& $e $Run.Benutzer) + ' am ' + (& $e $Run.Pc) + ' &middot; ' + (& $e $Now.ToString('dd.MM.yyyy HH:mm')) + '</p>')
    [void]$sb.Append('<div class="box">Testmodus mit Testbestand. Preise und Regeln im Test sind Beispielwerte. Der geführte Test unterstützt den Praxistest mit den Bediensteten, er ersetzt ihn nicht.</div>')
    [void]$sb.Append("<p><b>Bestätigt durch Menschen:</b> $($sum.Bestanden) bestanden, $($sum.FehlerMensch) Fehler, $($sum.Offen) offen von $($sum.Schritte) Schritten. <b>Automatisch geprüft:</b> $($sum.AutoOk) OK, $($sum.AutoFehler) Fehler.</p>")
    if ($sum.OffeneFehler.Count) {
        [void]$sb.Append('<h2>Offene Fehler</h2><ul>')
        foreach ($f in $sum.OffeneFehler) {
            $was = @(); if ($f.Mensch -eq 'Fehler') { $was += 'Mensch: ' + $f.Kommentar }; if ($f.Automatisch -eq 'Fehler') { $was += 'automatisch: ' + $f.AutomatischText }
            [void]$sb.Append('<li>' + (& $e "$($f.Szenario), Schritt $($f.Nr): $($was -join ' | ')") + '</li>')
        }
        [void]$sb.Append('</ul>')
    }
    [void]$sb.Append('<table><tr><th>Szenario</th><th>Nr</th><th>Aktion</th><th>Soll</th><th class="auto">Automatisch geprüft</th><th>Bestätigung Mensch</th><th>Ist laut Mensch / Kommentar</th><th>Zeit</th></tr>')
    foreach ($s in $Run.Steps) {
        $ac = $(if ($s.Automatisch -eq 'OK') { 'ok' } elseif ($s.Automatisch -eq 'Fehler') { 'err' } else { '' })
        $mc = $(if ($s.Mensch -eq 'Bestanden') { 'ok' } elseif ($s.Mensch -eq 'Fehler') { 'err' } else { '' })
        [void]$sb.Append('<tr><td>' + (& $e $s.Szenario) + '</td><td>' + $s.Nr + '</td><td>' + (& $e $s.Aktion) + '</td><td>' + (& $e $s.Soll) + '</td>')
        [void]$sb.Append('<td class="auto"><span class="' + $ac + '">' + (& $e $s.Automatisch) + '</span> ' + (& $e $s.AutomatischText) + '</td>')
        [void]$sb.Append('<td><span class="' + $mc + '">' + (& $e $s.Mensch) + '</span></td><td>' + (& $e $s.Kommentar) + '</td><td>' + (& $e $s.Zeit) + '</td></tr>')
    }
    [void]$sb.Append('</table></body></html>')
    $html = Join-Path $Folder "$base.html"
    [IO.File]::WriteAllText($html, $sb.ToString(), (New-Object System.Text.UTF8Encoding $true))
    return [pscustomobject]@{ Csv = $csv; Html = $html }
}

# ---------------------------------------------------------------- Fenster ----------------------------------------------------------------

<#
.SYNOPSIS
    Gefuehrten Test oeffnen (nur im Testmodus). Das Fenster bleibt neben der App offen, damit die echten Masken bedient werden koennen.
#>
function Show-HUTestAssistant {
    param([Parameter(Mandatory)][string]$ModuleId, [string]$ModuleVersion = '', [Parameter(Mandatory)][object[]]$Scenarios, [string]$Title = '')
    if (-not (Test-HUTestMode)) { Show-HUMessage 'Den geführten Test gibt es nur im Testmodus (Einstellungen › Verwalter › Testmodus starten).' -Icon Information; return }
    if ($script:HUTaWindow -and $script:HUTaWindow.IsVisible) { [void]$script:HUTaWindow.Activate(); return }
    $run = New-HUTestRun -ModuleId $ModuleId -ModuleVersion $ModuleVersion -Scenarios $Scenarios
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Geführter Test" Width="1040" Height="660" MinWidth="760" MinHeight="460" Background="#1E1E1E" ResizeMode="CanResizeWithGrip" WindowStartupLocation="CenterOwner">
    <Window.Resources><ResourceDictionary><!--HU:THEME--></ResourceDictionary></Window.Resources>
    <DockPanel Margin="12">
        <Border DockPanel.Dock="Top" Background="#E65100" Padding="10,6" Margin="0,0,0,10" ToolTip="Alles läuft gegen den Testbestand. Echte Daten, Outlook, Word und Mails bleiben unberührt.">
            <TextBlock x:Name="txtHead" Foreground="White" FontWeight="Bold" TextWrapping="Wrap"/>
        </Border>
        <DockPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
            <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
                <Button x:Name="btnExport" Content="Protokoll speichern" Background="#1565C0" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Protokoll als CSV und HTML im Testbestand speichern und den Ordner öffnen"/>
                <Button x:Name="btnClose" Content="Schließen" Background="#555555" Style="{StaticResource DarkButton}" ToolTip="Fenster schließen (nicht gespeicherte Ergebnisse gehen verloren)"/>
            </StackPanel>
            <TextBlock x:Name="txtSummary" Foreground="#CCCCCC" VerticalAlignment="Center" TextWrapping="Wrap" ToolTip="Stand des Testlaufs"/>
        </DockPanel>
        <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="3*" MinWidth="320"/><ColumnDefinition Width="8"/><ColumnDefinition Width="2*" MinWidth="300"/></Grid.ColumnDefinitions>
            <DataGrid x:Name="grdSteps" Grid.Column="0" Style="{StaticResource DarkDataGrid}" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" HeadersVisibility="Column"
                      ToolTip="Alle Schritte des Testlaufs. Schritt anklicken, um ihn zu bearbeiten.">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Szenario" Binding="{Binding Szenario}"/>
                    <DataGridTextColumn Header="Nr" Binding="{Binding Nr}"/>
                    <DataGridTextColumn Header="Aktion" Binding="{Binding Aktion}" Width="*"/>
                    <DataGridTextColumn Header="Automatisch" Binding="{Binding Automatisch}"/>
                    <DataGridTextColumn Header="Mensch" Binding="{Binding Mensch}"/>
                </DataGrid.Columns>
            </DataGrid>
            <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
                <StackPanel>
                    <TextBlock x:Name="txtStepTitle" Style="{StaticResource SectionTitle}" TextWrapping="Wrap"/>
                    <TextBlock Text="Aktion" Style="{StaticResource FieldLabel}" Margin="0,8,0,2"/>
                    <TextBlock x:Name="txtAction" Foreground="White" TextWrapping="Wrap" ToolTip="Das machst du jetzt in den echten Masken des Moduls"/>
                    <TextBlock Text="Soll-Ergebnis" Style="{StaticResource FieldLabel}" Margin="0,10,0,2"/>
                    <TextBlock x:Name="txtExpected" Foreground="White" TextWrapping="Wrap" ToolTip="So soll es danach aussehen"/>
                    <Border Background="#2A2A2A" BorderBrush="#3E3E42" BorderThickness="1" Padding="8" Margin="0,14,0,0" ToolTip="Ergebnis des Programms, getrennt von deiner Bestätigung">
                        <StackPanel>
                            <TextBlock Text="Automatisch geprüft (vom Programm)" Foreground="#9E9E9E" FontWeight="Bold"/>
                            <TextBlock x:Name="txtAuto" Foreground="#BDBDBD" TextWrapping="Wrap" Margin="0,4,0,6"/>
                            <Button x:Name="btnCheck" Content="Automatisch prüfen" Style="{StaticResource ToolButton}" HorizontalAlignment="Left" ToolTip="Technisch prüfbare Zustände kontrollieren (ersetzt deine Bestätigung nicht)"/>
                        </StackPanel>
                    </Border>
                    <Border Background="#2A2A2A" BorderBrush="#E65100" BorderThickness="1" Padding="8" Margin="0,10,0,0" ToolTip="Deine Beobachtung als testende Person">
                        <StackPanel>
                            <TextBlock Text="Deine Bestätigung" Foreground="#FFB74D" FontWeight="Bold"/>
                            <TextBlock x:Name="txtHuman" Foreground="#DDDDDD" Margin="0,4,0,6" TextWrapping="Wrap"/>
                            <TextBlock Text="Kommentar / Ist-Ergebnis (bei Fehler Pflicht)" Style="{StaticResource FieldLabel}"/>
                            <TextBox x:Name="txtComment" Style="{StaticResource DarkTextBox}" Height="70" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"
                                     ToolTip="Was ist tatsächlich passiert? Bei Fehler bitte beschreiben."/>
                            <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                                <Button x:Name="btnPass" Content="Bestanden" Background="#2E7D32" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Ergebnis entspricht dem Soll"/>
                                <Button x:Name="btnFail" Content="Fehler" Background="#C62828" Style="{StaticResource DarkButton}" ToolTip="Ergebnis weicht ab (Kommentar nötig)"/>
                            </StackPanel>
                        </StackPanel>
                    </Border>
                    <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                        <Button x:Name="btnPrev" Content="&#x2190; Zurück" Style="{StaticResource ToolButton}" Margin="0,0,8,0" ToolTip="Vorheriger Schritt"/>
                        <Button x:Name="btnNext" Content="Weiter &#x2192;" Style="{StaticResource ToolButton}" ToolTip="Nächster Schritt"/>
                    </StackPanel>
                </StackPanel>
            </ScrollViewer>
        </Grid>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    # Fenster bleibt offen, waehrend die Funktion endet: Zustand im Skriptbereich (Ereignisse laufen im Skriptbereich der App)
    $script:HUTa = @{ Run = $run; Index = 0; C = $c; W = $w }
    $w.Title = $(if ($Title) { $Title } else { "Geführter Test $([char]0x2013) $ModuleId" })
    $c.txtHead.Text = "$([char]::ConvertFromUtf32(0x1F9EA)) TESTMODUS $([char]0x2013) $ModuleId $ModuleVersion" + ': Masken des Moduls bedienen, hier je Schritt bestätigen. Preise und Regeln sind Beispielwerte.'
    $c.grdSteps.ItemsSource = $run.Steps
    $c.grdSteps.Add_SelectionChanged({ $t = $script:HUTa; $i = $t.C.grdSteps.SelectedIndex; if ($i -ge 0 -and $i -ne $t.Index) { $t.Index = $i; Update-HUTestAssistantView } })
    $c.btnCheck.Add_Click({ $t = $script:HUTa; [void](Invoke-HUTestCheck $t.Run.Steps[$t.Index]); Update-HUTestAssistantView })
    $c.btnPass.Add_Click({ $t = $script:HUTa; [void](Set-HUTestVerdict $t.Run.Steps[$t.Index] 'Bestanden' $t.C.txtComment.Text); if ($t.Index -lt $t.Run.Steps.Count - 1) { $t.Index++ }; Update-HUTestAssistantView })
    $c.btnFail.Add_Click({
            $t = $script:HUTa
            try { [void](Set-HUTestVerdict $t.Run.Steps[$t.Index] 'Fehler' $t.C.txtComment.Text) } catch { Show-HUMessage $_.Exception.Message -Icon Warning -Owner $t.W; return }
            Update-HUTestAssistantView
        })
    $c.btnPrev.Add_Click({ $t = $script:HUTa; if ($t.Index -gt 0) { $t.Index--; Update-HUTestAssistantView } })
    $c.btnNext.Add_Click({ $t = $script:HUTa; if ($t.Index -lt $t.Run.Steps.Count - 1) { $t.Index++; Update-HUTestAssistantView } })
    $c.btnExport.Add_Click({
            $t = $script:HUTa
            try {
                $r = Export-HUTestProtocol -Run $t.Run -Folder (Join-Path $script:DataRoot 'Testprotokolle')
                Write-HULogOK "Testprotokoll gespeichert: $($r.Html)"
                Start-Process -FilePath (Split-Path -Parent $r.Html) | Out-Null
            } catch { Show-HUMessage "Protokoll nicht gespeichert:`n$($_.Exception.Message)" -Icon Error -Owner $t.W }
        })
    $c.btnClose.Add_Click({ $script:HUTa.W.Close() })
    $w.Add_Closed({ $script:HUTaWindow = $null })
    $script:HUTaWindow = $w
    Update-HUTestAssistantView
    $w.Show()
}

function Update-HUTestAssistantView {
    $t = $script:HUTa; $c = $t.C
    $s = $t.Run.Steps[$t.Index]
    $c.txtStepTitle.Text = "$($s.Szenario) $([char]0x2013) Schritt $($s.Nr)"
    $c.txtAction.Text = $s.Aktion
    $c.txtExpected.Text = $s.Soll
    $c.txtAuto.Text = $(if (-not $s.HatPruefung) { 'Für diesen Schritt gibt es keine automatische Prüfung.' } elseif ($s.Automatisch -eq $script:HUTaNone) { 'Noch nicht geprüft.' } else { "$($s.Automatisch): $($s.AutomatischText)" })
    $c.btnCheck.IsEnabled = [bool]$s.HatPruefung
    $c.txtHuman.Text = $(if ($s.Mensch -eq 'offen') { 'Noch nicht bestätigt.' } else { "$($s.Mensch) ($($s.Zeit))" })
    $c.txtComment.Text = $s.Kommentar
    $c.btnPrev.IsEnabled = ($t.Index -gt 0)
    $c.btnNext.IsEnabled = ($t.Index -lt $t.Run.Steps.Count - 1)
    $sum = Get-HUTestSummary $t.Run
    $c.txtSummary.Text = "Bestätigt: $($sum.Bestanden) bestanden, $($sum.FehlerMensch) Fehler, $($sum.Offen) offen  |  Automatisch: $($sum.AutoOk) OK, $($sum.AutoFehler) Fehler"
    $c.grdSteps.Items.Refresh()
    if ($c.grdSteps.SelectedIndex -ne $t.Index) { $c.grdSteps.SelectedIndex = $t.Index }
}
