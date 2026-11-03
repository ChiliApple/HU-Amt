#Requires -Version 5.1
<#
.SYNOPSIS
    Sicherung der Daten (Data\) als verschluesselte .husicherung: Fenster, Zeitplan, Wiederherstellen.
.DESCRIPTION
    Einstellungen: Data\sicherung.json (gemeinsam fuer alle, im Netzordner fuer die ganze Installation):
      Folder (Ziel, z. B. Netzlaufwerk), Mode (aus | taeglich | woechentlich), Keep (je PC), Last (letzte Sicherung), LastBy.
    Kennwort fuer die Automatik: je Windows-Benutzer verschluesselt (DPAPI) in Data\Benutzer\<Name>\sicherung-kennwort.dat.
    Die Automatik laeuft beim Start bzw. stuendlich, wenn sie faellig ist und ein Kennwort gespeichert ist. Im Netzordner
    sichert der Erste, dem sie faellig wird (Last steht in der gemeinsamen Datei; Listen-Sperre waehrend des Laufs).
    Einstellen und Wiederherstellen nur fuer Verwalter.
.NOTES
    Dot-Source aus Main.ps1. Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:AbBusy = $false
$script:AbTimer = $null
$script:AbSettingsFile = ''

function Get-HUAbSettings {
    $script:AbSettingsFile = Join-Path $script:DataRoot 'sicherung.json'
    $s = Read-HUJsonFile $script:AbSettingsFile
    $o = [pscustomobject]@{ Folder = ''; Mode = 'aus'; Keep = 10; Last = ''; LastBy = '' }
    if ($s) { foreach ($p in $s.PSObject.Properties) { if ($o.PSObject.Properties[$p.Name]) { $o.($p.Name) = $p.Value } } }
    if ("$($o.Mode)" -notin 'aus', 'taeglich', 'woechentlich') { $o.Mode = 'aus' }
    try { $o.Keep = [Math]::Max(1, [Math]::Min(100, [int]$o.Keep)) } catch { $o.Keep = 10 }
    return $o
}
# Speichern unter der Listen-Sperre (Regel 3). Rueckgabe: '' oder Grund
function Save-HUAbSettings($S) {
    [void](Get-HUAbSettings)
    $l = Lock-HUDataFile -Path $script:AbSettingsFile -UserName $script:UserName
    if (-not $l.Ok) { return (Get-HULockInfoText $l) }
    try { Write-HUJsonFile -Path $script:AbSettingsFile -Object $S -Depth 5; return '' } finally { Unlock-HUDataFile $l }
}

# Ziel pruefen: nicht im Programmordner (die Sicherung wuerde sich selbst sichern bzw. mit dem Programm verschwinden)
function Test-HUAbFolder([string]$Folder) {
    if (-not "$Folder".Trim()) { return 'Bitte einen Ordner für die Sicherungen wählen.' }
    try {
        $f = [IO.Path]::GetFullPath($Folder).TrimEnd('\') + '\'
        $a = [IO.Path]::GetFullPath($script:AppRoot).TrimEnd('\') + '\'
        if ($f.StartsWith($a, [StringComparison]::OrdinalIgnoreCase)) { return 'Der Ordner liegt im Programmordner. Bitte einen Ordner außerhalb wählen, am besten auf einem Netzlaufwerk.' }
    } catch { return "Ordner ungültig: $($_.Exception.Message)" }
    return ''
}

# Kopfzeile: Stand der Sicherung (nur Verwalter sehen sie)
function Update-HUAbStatus {
    $t = $script:Controls['txtBackupStatus']
    if (-not $t) { return }
    if (-not (Test-HUVerwalter)) { $t.Visibility = 'Collapsed'; return }
    $s = Get-HUAbSettings
    $last = ''; try { if ("$($s.Last)") { $last = ([datetime]::Parse("$($s.Last)", [Globalization.CultureInfo]::InvariantCulture)).ToString('dd.MM. HH:mm') } } catch { }
    $t.Text = $(if ($s.Mode -eq 'aus') { 'Sicherung: aus' } elseif ($last) { "Sicherung: $last" } else { 'Sicherung: noch keine' })
    $t.Foreground = Get-HUBrush $(if ($s.Mode -eq 'aus' -or -not $last) { '#FFA726' } else { '#858585' })
    $t.ToolTip = "Sicherung der Daten$(if ($s.LastBy) { " - zuletzt von $($s.LastBy)" }). Klick: Sicherung einstellen"
    $t.Visibility = 'Visible'
}

# Sicherung im Hintergrund. $Password: SecureString. -Auto: still, Ergebnis als Hinweis unten rechts
function Start-HUAppBackup([object]$Password, [switch]$Auto, [scriptblock]$OnDone = $null) {
    if ($script:AbBusy) { return }
    $s = Get-HUAbSettings
    $why = Test-HUAbFolder $s.Folder
    if ($why) { if (-not $Auto) { Show-HUMessage $why -Icon Info }; return }
    $l = Lock-HUDataFile -Path $script:AbSettingsFile -UserName $script:UserName
    if (-not $l.Ok) { if (-not $Auto) { Show-HUMessage "Es läuft gerade eine Sicherung.`n`n$(Get-HULockInfoText $l)" -Icon Info }; return }
    $script:AbBusy = $true
    $script:AbLock = $l
    $script:AbOnDone = $OnDone
    $script:AbAuto = [bool]$Auto
    $mods = @(Get-HUAbModuleList @((Join-Path $script:AppRoot 'Modules'), (Join-Path $script:AppRoot 'Addons')))
    Write-HULogInfo "Sicherung läuft ... ($($s.Folder))"
    Invoke-AsyncCommand -ScriptBlock {
        param($mod, $data, $pw, $folder, $ver, $keep, $mods)
        try {
            Import-Module $mod -DisableNameChecking
            $src = @(Get-HUAppBackupSources -DataRoot $data)
            $r = New-HUAppBackup -Sources $src -Password $pw -Folder $folder -AppVersion $ver -Keep $keep -Modules $mods
            return "OK:$($r.Files)|$([Math]::Round($r.Bytes / 1MB, 1))|$(Split-Path -Leaf $r.Path)|$(@($r.Removed).Count)"
        } catch { return "ERR:$($_.Exception.Message)" }
    } -ArgumentList @((Join-Path $script:AppRoot 'Core\HU.AppBackup.psm1'), $script:DataRoot, $Password, $s.Folder, "$($script:Version)", [int]$s.Keep, $mods) -TimeoutSec 900 -OnComplete {
        param($res)
        $r = "$res".Trim()
        $ok = $r -match '^OK:(\d+)\|([\d\.,]+)\|([^|]+)\|(\d+)$'
        try {
            if ($ok) {
                $st = Get-HUAbSettings
                $st.Last = (Get-Date).ToString('s'); $st.LastBy = "$($script:UserName) ($env:COMPUTERNAME)"
                try { Write-HUJsonFile -Path $script:AbSettingsFile -Object $st -Depth 5 } catch { }
            }
        } finally { Unlock-HUDataFile $script:AbLock; $script:AbLock = $null; $script:AbBusy = $false }
        if ($ok) {
            $msg = "Sicherung fertig: $($Matches[3]) ($($Matches[1]) Dateien, $($Matches[2]) MB)$(if ([int]$Matches[4]) { ", $($Matches[4]) alte entfernt" })"
            Write-HULogOK $msg
            if ($script:AbAuto) { Show-HUToast -Title 'Sicherung fertig' -Text "$($Matches[1]) Dateien gesichert." -Icon Success }
        } else {
            $msg = "Sicherung fehlgeschlagen: $($r -replace '^(ERR:|FEHLER:)\s*', '')"
            Write-HULogError $msg
            if ($script:AbAuto) { Show-HUToast -Title 'Sicherung fehlgeschlagen' -Text "$($r -replace '^(ERR:|FEHLER:)\s*', '')" -Icon Error } else { Show-HUMessage $msg -Icon Error }
        }
        try { Update-HUAbStatus } catch { }
        if ($script:AbOnDone) { try { & $script:AbOnDone $ok } catch { } }
    }
}

# Automatik: beim Start und dann stuendlich pruefen
function Start-HUAppBackupScheduler {
    $run = {
        try {
            $s = Get-HUAbSettings
            if (-not (Test-HUAppBackupDue -Mode $s.Mode -Last $s.Last)) { return }
            $pw = Get-HUAppBackupPassword -DataRoot $script:UserRoot
            if (-not $pw) { return }
            Start-HUAppBackup -Password $pw -Auto
        } catch { Write-HULogWarn "Sicherung (Automatik): $($_.Exception.Message)" }
    }
    & $run
    $script:AbTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:AbTimer.Interval = [TimeSpan]::FromHours(1)
    $script:AbTimer.Add_Tick($run)
    $script:AbTimer.Start()
}

# kleines Fenster: Kennwort eingeben. Rueckgabe SecureString oder $null
function Read-HUAbPassword([string]$Text, $Owner = $null) {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Kennwort der Sicherung" Width="460" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="NoResize" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel>
        <Border DockPanel.Dock="Bottom" Background="#252526" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="16,10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOk" Content="OK" MinWidth="96" Background="#1976D2" Style="{StaticResource DarkButton}" IsDefault="True" Margin="0,0,8,0" ToolTip="Kennwort übernehmen"/>
                <Button x:Name="btnCancel" Content="Abbrechen" MinWidth="96" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Abbrechen"/>
            </StackPanel>
        </Border>
        <StackPanel Margin="18">
            <TextBlock x:Name="lbl" Foreground="#E0E0E0" TextWrapping="Wrap" Margin="0,0,0,10"/>
            <PasswordBox x:Name="pwd" Background="#2D2D30" Foreground="White" BorderBrush="#3E3E42" Padding="4,3" ToolTip="Kennwort der Sicherung"/>
        </StackPanel>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $c.lbl.Text = $Text
    $st = @{ P = $null }
    $c.btnOk.Add_Click({ if ($c.pwd.SecurePassword.Length) { $st.P = $c.pwd.SecurePassword.Copy(); $w.Close() } })
    $w.Add_ContentRendered({ [void]$c.pwd.Focus() })
    [void]$w.ShowDialog()
    return $st.P
}

# Wiederherstellen: Datei pruefen (Hintergrund), Bereiche waehlen, vorher aktuellen Stand sichern, dann neu starten
function Start-HUAppBackupRestore([string]$File, $Owner) {
    $others = Get-HUOtherSessionsText
    if ($others) { Show-HUMessage "Wiederherstellen geht erst, wenn alle anderen $($script:AppTitle) geschlossen haben.`n`nBitte zuerst schließen:`n$others" -Icon Warning -Owner $Owner; return }
    $pw = Read-HUAbPassword "Kennwort der Sicherung $(Split-Path -Leaf $File):" $Owner
    if (-not $pw) { return }
    $st = [hashtable]::Synchronized(@{ Opened = $null })
    Write-HULogInfo "Sicherung wird geprüft: $(Split-Path -Leaf $File) ..."
    $script:AbRestoreOwner = $Owner
    $script:AbRestoreFile = $File
    Invoke-AsyncCommand -ScriptBlock {
        param($mod, $file, $pw, $ver, $app, $data, $st)
        try {
            Import-Module $mod -DisableNameChecking
            $o = Open-HUAppBackup -Path $file -Password $pw -AppVersion $ver -AppRoot $app -DataRoot $data
            $st.Opened = $o
            if ($o.Errors.Count) { return "ERR:$(@($o.Errors)[0..4] -join ' ')" }
            return "OK:$(@($o.Rows).Count)"
        } catch { return "ERR:$($_.Exception.Message)" }
    } -ArgumentList @((Join-Path $script:AppRoot 'Core\HU.AppBackup.psm1'), $File, $pw, "$($script:Version)", $script:AppRoot, $script:DataRoot, $st) -State $st -TimeoutSec 900 -OnComplete {
        param($res, $st)
        $r = "$res".Trim()
        if ($r -notmatch '^OK:') { Show-HUMessage "Die Sicherung lässt sich nicht wiederherstellen:`n`n$($r -replace '^(ERR:|FEHLER:)\s*', '')" -Icon Error -Owner $script:AbRestoreOwner; return }
        $o = $st.Opened
        $areas = @($o.Rows | Group-Object Area | Sort-Object Name | ForEach-Object {
                [pscustomobject]@{ Area = $_.Name; Count = $_.Count; Pre = -not (@($_.Group | Where-Object { Test-HUAppBackupCodeName $_.Name }).Count) } })
        $sel = Select-HUAbAreas $areas $o.Header $script:AbRestoreOwner
        if (-not $sel -or -not @($sel).Count) { return }
        if (-not (Confirm-HU -Warning "Diese Bereiche jetzt wiederherstellen?`n`n  $($sel -join "`n  ")`n`nDer aktuelle Stand dieser Dateien kommt vorher nach Data\Wiederherstellung. Danach startet $($script:AppTitle) neu." 'Wiederherstellen' -Owner $script:AbRestoreOwner)) { return }
        try {
            $w = Restore-HUAppBackup -Opened $o -Areas @($sel) -AppRoot $script:AppRoot -DataRoot $script:DataRoot
            $script:AbRestored = $true
            Write-HULogOK "Wiederhergestellt: $($w.Written) Dateien aus $(Split-Path -Leaf $script:AbRestoreFile). Vorheriger Stand: $($w.Before)"
            Show-HUMessage "$($w.Written) Dateien wiederhergestellt.`n`nDer vorherige Stand liegt in:`n$($w.Before)`n`n$($script:AppTitle) startet jetzt neu." -Owner $script:AbRestoreOwner
            Restart-HUApp
        } catch { Show-HUMessage "Wiederherstellen fehlgeschlagen:`n$($_.Exception.Message)" -Icon Error -Owner $script:AbRestoreOwner }
    }
}

# Bereiche zum Wiederherstellen waehlen. Rueckgabe: Namen der Bereiche oder $null
function Select-HUAbAreas([object[]]$Areas, $Header, $Owner) {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Wiederherstellen" Width="520" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel>
        <Border DockPanel.Dock="Bottom" Background="#252526" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="16,10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOk" Content="Wiederherstellen" MinWidth="140" Background="#C62828" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Gewählte Bereiche wiederherstellen"/>
                <Button x:Name="btnCancel" Content="Abbrechen" MinWidth="96" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" IsDefault="True" ToolTip="Nichts ändern"/>
            </StackPanel>
        </Border>
        <StackPanel Margin="18">
            <TextBlock x:Name="lbl" Foreground="#E0E0E0" TextWrapping="Wrap" Margin="0,0,0,10"/>
            <StackPanel x:Name="pnl"/>
        </StackPanel>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $created = ''; try { $created = ([datetime]::Parse("$($Header.Created)")).ToLocalTime().ToString('dd.MM.yyyy HH:mm') } catch { }
    $c.lbl.Text = "Sicherung vom $created (Version $($Header.App)). Welche Bereiche sollen zurück? Nicht gewählte bleiben, wie sie sind."
    $boxes = @()
    foreach ($a in $Areas) {
        $cb = New-Object System.Windows.Controls.CheckBox
        $cb.Content = "$($a.Area) ($($a.Count) Dateien)"; $cb.Tag = $a.Area; $cb.IsChecked = [bool]$a.Pre
        $cb.Foreground = Get-HUBrush '#E0E0E0'; $cb.Margin = [System.Windows.Thickness]::new(0, 3, 0, 3)
        try { $cb.Style = $w.FindResource('DarkCheckBox') } catch { }
        $cb.ToolTip = $(if ($a.Pre) { 'Bereich wiederherstellen' } else { 'Nicht vorgewählt: legt fest, wer Verwalter ist' })
        [void]$c.pnl.Children.Add($cb); $boxes += $cb
    }
    $st = @{ Sel = $null }
    $c.btnOk.Add_Click({ $st.Sel = @($boxes | Where-Object { $_.IsChecked } | ForEach-Object { "$($_.Tag)" }); $w.Close() })
    [void]$w.ShowDialog()
    return $st.Sel
}

# Fenster: Sicherung einstellen, jetzt sichern, wiederherstellen (nur Verwalter)
function Show-HUAppBackupDialog {
    if (-not (Test-HUVerwalter)) { Show-HUMessage 'Die Sicherung stellt ein Verwalter ein.' -Icon Info; return }
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sicherung" Width="760" Height="620" MinWidth="640" MinHeight="520" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel Margin="16">
        <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" Foreground="#CCCCCC" FontSize="14" Margin="0,0,0,12"
                   Text="Sichert alle Daten der App verschlüsselt in einen Ordner, am besten auf ein Netzlaufwerk. Ohne Kennwort lässt sich eine Sicherung nicht öffnen."/>
        <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Text="Ordner" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtFolder" Grid.Row="0" Grid.Column="1" Style="{StaticResource DarkTextBox}" IsReadOnly="True" ToolTip="Hier landen die Sicherungen"/>
            <Button x:Name="btnFolder" Grid.Row="0" Grid.Column="2" Content="Wählen ..." Style="{StaticResource DarkButton}" Background="#3E3E42" Margin="8,0,0,0" ToolTip="Ordner für die Sicherungen wählen"/>
            <TextBlock Grid.Row="1" Text="Automatisch" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,10,10,0"/>
            <StackPanel Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,10,0,0">
                <ComboBox x:Name="cmbMode" Style="{StaticResource DarkComboBox}" Width="160" ToolTip="Wie oft automatisch gesichert wird (braucht ein gespeichertes Kennwort)">
                    <ComboBoxItem Content="aus" Tag="aus"/>
                    <ComboBoxItem Content="täglich" Tag="taeglich"/>
                    <ComboBoxItem Content="wöchentlich" Tag="woechentlich"/>
                </ComboBox>
                <TextBlock Text="behalten" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="18,0,8,0"/>
                <TextBox x:Name="txtKeep" Style="{StaticResource DarkTextBox}" Width="50" ToolTip="So viele Sicherungen je PC bleiben, ältere werden gelöscht"/>
                <TextBlock Text="je PC" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="8,0,0,0"/>
            </StackPanel>
            <TextBlock Grid.Row="2" Text="Kennwort" Style="{StaticResource FieldLabel}" VerticalAlignment="Center" Margin="0,10,10,0"/>
            <StackPanel Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,10,0,0">
                <PasswordBox x:Name="pwd1" Width="200" Background="#2D2D30" Foreground="White" BorderBrush="#3E3E42" Padding="4,3" ToolTip="Kennwort der Sicherung (mindestens 10 Zeichen)"/>
                <PasswordBox x:Name="pwd2" Width="200" Background="#2D2D30" Foreground="White" BorderBrush="#3E3E42" Padding="4,3" Margin="8,0,0,0" ToolTip="Kennwort wiederholen"/>
            </StackPanel>
            <CheckBox x:Name="chkRemember" Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="2" Style="{StaticResource DarkCheckBox}" Margin="0,8,0,0"
                      Content="Kennwort für die automatische Sicherung merken (verschlüsselt, nur für meinen Windows-Benutzer)"
                      ToolTip="Ohne gespeichertes Kennwort sichert die Automatik nicht"/>
        </Grid>
        <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,4,0,10">
            <Button x:Name="btnNow" Content="&#x1F4BE; Jetzt sichern" Style="{StaticResource DarkButton}" Background="#388E3C" FontSize="15" Padding="22,10" ToolTip="Einstellungen speichern und sofort sichern"/>
            <Ellipse x:Name="dotAb" Width="10" Height="10" Fill="#4CAF50" Margin="14,0,6,0" VerticalAlignment="Center" Visibility="Collapsed" ToolTip="Sicherung läuft"/>
            <TextBlock x:Name="txtState" Foreground="#858585" VerticalAlignment="Center" Margin="6,0,0,0" ToolTip="Stand der Sicherung"/>
        </StackPanel>
        <Border DockPanel.Dock="Bottom" Background="#252526" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,10,0,0" Margin="0,10,0,0">
            <DockPanel>
                <Button x:Name="btnClose" DockPanel.Dock="Right" Content="Schließen" MinWidth="96" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Einstellungen speichern und schließen"/>
                <Button x:Name="btnOpen" DockPanel.Dock="Right" Content="Ordner öffnen" Background="#3E3E42" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Ordner mit den Sicherungen öffnen"/>
                <StackPanel Orientation="Horizontal">
                    <Button x:Name="btnRestore" Content="Wiederherstellen ..." Background="#C62828" Style="{StaticResource DarkButton}" ToolTip="Markierte Sicherung wiederherstellen (fragt nach dem Kennwort und den Bereichen)"/>
                    <Button x:Name="btnRestoreFile" Content="Andere Datei ..." Background="#3E3E42" Style="{StaticResource DarkButton}" Margin="8,0,0,0" ToolTip="Sicherung aus einem anderen Ordner wiederherstellen"/>
                </StackPanel>
            </DockPanel>
        </Border>
        <DataGrid x:Name="grid" Style="{StaticResource DarkDataGrid}" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" SelectionMode="Single"
                  ToolTip="Vorhandene Sicherungen (neueste oben)">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Datum" Binding="{Binding Datum}"/>
                <DataGridTextColumn Header="PC" Binding="{Binding PC}"/>
                <DataGridTextColumn Header="Größe" Binding="{Binding Groesse}"/>
                <DataGridTextColumn Header="Datei" Binding="{Binding Datei}"/>
            </DataGrid.Columns>
        </DataGrid>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    Restore-HUDialogState $w 'dlgBackup'
    $s = Get-HUAbSettings
    $c.txtFolder.Text = "$($s.Folder)"
    $c.txtKeep.Text = "$($s.Keep)"
    foreach ($it in $c.cmbMode.Items) { if ("$($it.Tag)" -eq "$($s.Mode)") { $c.cmbMode.SelectedItem = $it } }
    $c.chkRemember.IsChecked = [bool](Get-HUAppBackupPassword -DataRoot $script:UserRoot)
    $fill = {
        $rows = @(Get-HUAppBackupList "$($c.txtFolder.Text)" | ForEach-Object {
                $pc = ''; if ($_.Name -match '^HU-Sicherung_(.+)_\d{4}-\d{2}-\d{2}_\d{6}\.husicherung$') { $pc = $Matches[1] }
                [pscustomobject]@{ Datum = $_.LastWriteTime.ToString('dd.MM.yyyy HH:mm'); PC = $pc; Groesse = ('{0:N1} MB' -f ($_.Length / 1MB)); Datei = $_.Name; Path = $_.FullName } })
        $c.grid.ItemsSource = $rows
        $cur = Get-HUAbSettings
        $last = ''; try { if ("$($cur.Last)") { $last = ([datetime]::Parse("$($cur.Last)", [Globalization.CultureInfo]::InvariantCulture)).ToString('dd.MM.yyyy HH:mm') } } catch { }
        $hasPw = [bool](Get-HUAppBackupPassword -DataRoot $script:UserRoot)
        $c.txtState.Text = $(if ($last) { "Zuletzt gesichert: $last$(if ($cur.LastBy) { " von $($cur.LastBy)" })" } else { 'Noch keine Sicherung.' }) + $(if ($hasPw) { '  |  Kennwort gespeichert' } else { '' })
    }
    & $fill
    # Einstellungen aus dem Fenster uebernehmen. Rueckgabe: '' oder Grund
    $save = {
        $n = Get-HUAbSettings
        $n.Folder = "$($c.txtFolder.Text)".Trim()
        $n.Mode = "$($c.cmbMode.SelectedItem.Tag)"
        $k = 10; [void][int]::TryParse("$($c.txtKeep.Text)", [ref]$k); $n.Keep = [Math]::Max(1, [Math]::Min(100, $k))
        if ($n.Mode -ne 'aus' -or $n.Folder) { $why = Test-HUAbFolder $n.Folder; if ($why) { return $why } }
        $why = Save-HUAbSettings $n
        if ($why) { return $why }
        Update-HUAbStatus
        return ''
    }
    $c.btnFolder.Add_Click({
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = 'Ordner für die Sicherungen (am besten ein Netzlaufwerk)'
            if ("$($c.txtFolder.Text)") { $dlg.SelectedPath = "$($c.txtFolder.Text)" }
            if ($dlg.ShowDialog() -eq 'OK') {
                $why = Test-HUAbFolder $dlg.SelectedPath
                if ($why) { Show-HUMessage $why -Icon Warning -Owner $w; return }
                $c.txtFolder.Text = $dlg.SelectedPath; & $fill
            }
        })
    # Haken weg = gespeichertes Kennwort loeschen, aber nur nach Rueckfrage
    $c.chkRemember.Add_Unchecked({
            if (-not (Get-HUAppBackupPassword -DataRoot $script:UserRoot)) { return }
            if (Confirm-HU -Warning "Gespeichertes Kennwort der Sicherung löschen?

Dann sichert die Automatik für deinen Benutzer nicht mehr, bis du es wieder speicherst." 'Sicherung' -Owner $w) {
                Remove-HUAppBackupPassword -DataRoot $script:UserRoot; Write-HULogInfo 'Gespeichertes Kennwort der Sicherung gelöscht.'; & $fill
            } else { $c.chkRemember.IsChecked = $true }
        })
    $c.btnOpen.Add_Click({ if ("$($c.txtFolder.Text)") { Open-HUPath "$($c.txtFolder.Text)" } })
    $c.btnNow.Add_Click({
            $why = & $save
            if ($why) { Show-HUMessage $why -Icon Warning -Owner $w; return }
            $pw = $null
            if ($c.pwd1.SecurePassword.Length) {
                $why = Test-HUAbPassword $c.pwd1.SecurePassword $c.pwd2.SecurePassword
                if ($why) { Show-HUMessage $why -Icon Warning -Owner $w; return }
                $pw = $c.pwd1.SecurePassword.Copy()
                if ($c.chkRemember.IsChecked) { Save-HUAppBackupPassword -DataRoot $script:UserRoot -Password $pw; Write-HULogOK 'Kennwort der Sicherung gespeichert (nur für diesen Windows-Benutzer).' }
            } else {
                $pw = Get-HUAppBackupPassword -DataRoot $script:UserRoot
                if (-not $pw) { Show-HUMessage 'Bitte das Kennwort der Sicherung zweimal eingeben.' -Icon Info -Owner $w; return }
            }
            $c.pwd1.Clear(); $c.pwd2.Clear()
            Set-HUBusyDot $c.dotAb $true; $c.btnNow.IsEnabled = $false; $c.txtState.Text = 'Sicherung läuft ...'
            Start-HUAppBackup -Password $pw -OnDone { param($ok) try { Set-HUBusyDot $c.dotAb $false; $c.btnNow.IsEnabled = $true; & $fill } catch { } }.GetNewClosure()
        })
    $c.btnRestore.Add_Click({
            $r = $c.grid.SelectedItem
            if (-not $r) { Show-HUMessage 'Bitte zuerst eine Sicherung in der Liste markieren.' -Icon Info -Owner $w; return }
            Start-HUAppBackupRestore $r.Path $w
        })
    $c.btnRestoreFile.Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Filter = 'HU-Sicherung (*.husicherung)|*.husicherung'
            if ($dlg.ShowDialog($w)) { Start-HUAppBackupRestore $dlg.FileName $w }
        })
    $w.Add_Closing({
            param($sd, $e)
            $why = & $save
            if ($why -and -not (Confirm-HU "Einstellungen nicht gespeichert:`n$why`n`nTrotzdem schließen?" 'Sicherung' -Owner $w)) { $e.Cancel = $true; return }
            Save-HUDialogState $w 'dlgBackup'
        })
    [void]$w.ShowDialog()
}
