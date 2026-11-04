#Requires -Version 5.1
<#
.SYNOPSIS
    App-Sicherung wie in der Vorlage: Reiter "Sicherung" in den Einstellungen, Jetzt sichern, Pruefen, Wiederherstellen, Automatik.
.DESCRIPTION
    Logik in Core\HU.AppBackup.psm1. Hier nur Oberflaeche und Ablauf (uebernommen aus der Vorlage).
    Abweichungen von der Vorlage (Konzept: Datenordner im Netz fuer alle):
      - Einstellungen und Stand in Data\sicherung.json (gemeinsam fuer die Installation, mit Listen-Sperre gespeichert):
        Folder, Mode (aus | taeglich | woechentlich), Keep, Last, LastBy, Result.
        Im Netzordner sichert der Erste, dem die Automatik faellig wird; waehrend des Laufs ist die Datei gesperrt.
      - Kennwort fuer die Automatik je Windows-Benutzer verschluesselt (DPAPI) in Data\Benutzer\<Name>\sicherung-kennwort.dat.
      - Einstellen, Pruefen und Wiederherstellen nur fuer Verwalter; die Diskette in der Kopfzeile sehen nur Verwalter.
      - Wiederherstellen erst, wenn keine andere Sitzung offen ist.
      - Sicherungsordner nicht im Programmordner (die Sicherung wuerde sich sonst selbst mitsichern).
.NOTES
    Dot-Source aus Main.ps1. Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:AbDlg = $null   # offener Einstellungs-Dialog: @{ C; W; Pw }
# StrictMode: alle Zustandsvariablen vorab anlegen (sonst Fehler beim ersten Lesen)
$script:AbProg = $null        # Fortschritt (synchronisierte Hashtable) waehrend Sichern/Pruefen
$script:AbProgTimer = $null   # Anzeige-Timer
$script:AbAuto = $false       # laufende Sicherung ist automatisch
$script:AbOpen = $null        # laufendes Pruefen/Lesen: @{ File; Info; CheckOnly; Owner }
$script:AbCheck = $null       # Pruefung der Automatik
$script:AbTimer = $null       # Automatik-Timer (alle 15 Minuten)
$script:AbLock = $null        # Listen-Sperre auf Data\sicherung.json waehrend der Sicherung
$script:AbSettingsFile = ''

# Ablage des Kennworts fuer die Automatik (je Windows-Benutzer)
function Get-HUAbDataRoot { return $script:UserRoot }

# Einstellungen und Stand (gemeinsam): Data\sicherung.json
function Get-HUAbConfig {
    $script:AbSettingsFile = Join-Path $script:DataRoot 'sicherung.json'
    $s = Read-HUJsonFile $script:AbSettingsFile
    $o = [pscustomobject]@{ Folder = ''; Mode = 'aus'; Keep = 10; Last = ''; LastBy = ''; Result = '' }
    if ($s) { foreach ($p in $s.PSObject.Properties) { if ($o.PSObject.Properties[$p.Name]) { $o.($p.Name) = $p.Value } } }
    if ("$($o.Mode)" -notin 'aus', 'taeglich', 'woechentlich') { $o.Mode = 'aus' }
    try { $o.Keep = [int]$o.Keep } catch { $o.Keep = 10 }
    if ($o.Keep -lt 1 -or $o.Keep -gt 999) { $o.Keep = 10 }
    $o | Add-Member -NotePropertyName Path -NotePropertyValue ([Environment]::ExpandEnvironmentVariables("$($o.Folder)".Trim())) -Force
    return $o
}

# Speichern unter der Listen-Sperre. Rueckgabe: '' oder Grund
function Save-HUAbConfig($Cfg) {
    [void](Get-HUAbConfig)
    $o = [pscustomobject][ordered]@{ Folder = "$($Cfg.Folder)"; Mode = "$($Cfg.Mode)"; Keep = [int]$Cfg.Keep; Last = "$($Cfg.Last)"; LastBy = "$($Cfg.LastBy)"; Result = "$($Cfg.Result)" }
    # waehrend der eigenen Sicherung haelt dieses Programm die Sperre schon
    if ($script:AbLock) { Write-HUJsonFile -Path $script:AbSettingsFile -Object $o -Depth 5; return '' }
    $l = Lock-HUDataFile -Path $script:AbSettingsFile -UserName $script:UserName
    if (-not $l.Ok) { return (Get-HULockInfoText $l) }
    try { Write-HUJsonFile -Path $script:AbSettingsFile -Object $o -Depth 5; return '' } finally { Unlock-HUDataFile $l }
}

# Stand nach einer Sicherung bzw. einem Fehlschlag festhalten
function Set-HUAbResult([string]$Result, [switch]$Success) {
    try {
        $cfg = Get-HUAbConfig
        $cfg.Result = $Result
        if ($Success) { $cfg.Last = (Get-Date).ToString('s'); $cfg.LastBy = "$($script:UserName) ($env:COMPUTERNAME)" }
        $why = Save-HUAbConfig $cfg
        if ($why) { Write-HULogWarn "Sicherung: Stand nicht gespeichert ($why)" }
    } catch { Write-HULogWarn "Sicherung: Stand nicht gespeichert ($($_.Exception.Message))" }
}

# Ziel pruefen: nicht im Programmordner (die Sicherung wuerde sich selbst sichern bzw. mit dem Programm verschwinden)
function Test-HUAbFolder([string]$Folder) {
    if (-not "$Folder".Trim()) { return 'Bitte zuerst einen Sicherungsordner angeben.' }
    try {
        $f = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($Folder)).TrimEnd('\') + '\'
        $a = [IO.Path]::GetFullPath($script:AppRoot).TrimEnd('\') + '\'
        if ($f.StartsWith($a, [StringComparison]::OrdinalIgnoreCase)) { return 'Der Sicherungsordner liegt im Programmordner. Bitte einen Ordner außerhalb wählen, am besten auf einem anderen Laufwerk.' }
    } catch { return "Sicherungsordner ungültig: $($_.Exception.Message)" }
    return ''
}

# ---------------------------------------------------------------- Kennwort-Dialog ----------------------------------------------------------------
# $New: festlegen (zweimal, mind. 10 Zeichen, optional fuer die Automatik merken). Sonst: einmal eingeben (zum Oeffnen).
# Rueckgabe: @{ Password (SecureString); Remember } oder $null.
function Show-HUAbPasswordDialog($Owner = $null, [switch]$New, [string]$Text = '') {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Kennwort der App-Sicherung" Width="420" MinWidth="420" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <StackPanel Margin="18">
        <TextBlock x:Name="txtInfo" Style="{StaticResource HintText}" TextWrapping="Wrap" Margin="0,0,0,10"/>
        <TextBlock Text="Kennwort" Style="{StaticResource FieldLabel}" Margin="0,0,0,3"/>
        <PasswordBox x:Name="pw1" Height="28" Background="#2D2D2D" Foreground="White" BorderBrush="#3E3E42" Padding="6,3" ToolTip="Kennwort der Sicherung"/>
        <StackPanel x:Name="pnlNew">
            <TextBlock Text="Kennwort wiederholen" Style="{StaticResource FieldLabel}" Margin="0,10,0,3"/>
            <PasswordBox x:Name="pw2" Height="28" Background="#2D2D2D" Foreground="White" BorderBrush="#3E3E42" Padding="6,3" ToolTip="Kennwort noch einmal eingeben"/>
            <CheckBox x:Name="chkRemember"  IsChecked="True" Style="{StaticResource DarkCheckBox}" Margin="0,10,0,0" ToolTip="Verschlüsselt, nur für deinen Windows-Benutzer. Ohne gespeichertes Kennwort sichert die Automatik nicht."><TextBlock Text="F&#xFC;r die automatische Sicherung merken" TextWrapping="Wrap"/></CheckBox>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnOk" Content="OK" Width="110" Background="#388E3C" Style="{StaticResource DarkButton}" IsDefault="True" Margin="0,0,8,0" ToolTip="Kennwort übernehmen"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Abbrechen"/>
        </StackPanel>
    </StackPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $c.txtInfo.Text = $(if ($Text) { $Text } elseif ($New) { 'Mindestens 10 Zeichen. Ohne dieses Kennwort ist keine Sicherung zu öffnen - bitte sicher aufbewahren (z. B. im Kennwort-Manager). Ein neues Kennwort gilt nur für neue Sicherungen.' } else { 'Kennwort der Sicherung eingeben.' })
    if (-not $New) { $c.pnlNew.Visibility = 'Collapsed' }
    $st = @{ R = $null }
    $c.btnOk.Add_Click({
            if ($New) {
                $why = Test-HUAbPassword $c.pw1.SecurePassword $c.pw2.SecurePassword
                if ($why) { Show-HUMessage $why -Icon Warning -Owner $w; return }
            } elseif (-not $c.pw1.SecurePassword.Length) { return }
            $st.R = [pscustomobject]@{ Password = $c.pw1.SecurePassword.Copy(); Remember = [bool]($New -and $c.chkRemember.IsChecked) }
            $w.Close()
        })
    $w.Add_ContentRendered({ $c.pw1.Focus() })
    [void]$w.ShowDialog()
    return $st.R
}

# ---------------------------------------------------------------- Reiter in den Einstellungen ----------------------------------------------------------------
function Initialize-HUAbSettingsTab($c, $w) {
    $script:AbDlg = @{ C = $c; W = $w; Pw = $null }
    $cfg = Get-HUAbConfig
    $c.txtAbPath.Text = "$($cfg.Folder)"
    $c.txtAbKeep.Text = "$($cfg.Keep)"
    foreach ($i in $c.cmbAbMode.Items) { if ("$($i.Tag)" -eq $cfg.Mode) { $c.cmbAbMode.SelectedItem = $i } }
    if (-not $c.cmbAbMode.SelectedItem) { $c.cmbAbMode.SelectedIndex = 0 }
    Update-HUAbTabInfo
    if ((Test-HUJobRunning 'AppBackup') -or (Test-HUJobRunning 'AppBackupOpen')) { Start-HUAbProgressTimer }
    $c.btnAbBrowse.Add_Click({
            $f = New-Object System.Windows.Forms.FolderBrowserDialog
            $f.Description = 'Ordner für die App-Sicherung (auch Netzwerkfreigabe oder USB-Laufwerk)'
            $cc = $script:AbDlg.C
            if ($cc.txtAbPath.Text -and (Test-Path -LiteralPath $cc.txtAbPath.Text)) { $f.SelectedPath = $cc.txtAbPath.Text }
            if ($f.ShowDialog() -eq 'OK') {
                $why = Test-HUAbFolder $f.SelectedPath
                if ($why) { Show-HUMessage $why -Icon Warning -Owner $script:AbDlg.W; return }
                $cc.txtAbPath.Text = $f.SelectedPath
            }
        })
    $c.btnAbPw.Add_Click({
            $r = Show-HUAbPasswordDialog $script:AbDlg.W -New
            if (-not $r) { return }
            $script:AbDlg.Pw = $r.Password
            $dr = Get-HUAbDataRoot
            try {
                if ($r.Remember) { Save-HUAppBackupPassword -DataRoot $dr -Password $r.Password; Write-HULogOK 'App-Sicherung: Kennwort gespeichert (DPAPI, dieser Windows-Benutzer).' }
                else { Remove-HUAppBackupPassword $dr; Write-HULogInfo 'App-Sicherung: Kennwort nur für diese Sitzung (nicht gespeichert) - die Automatik läuft so nicht.' }
            } catch { Show-HUMessage "Kennwort nicht gespeichert: $($_.Exception.Message)" -Icon Error -Owner $script:AbDlg.W }
            Update-HUAbTabInfo
        })
    $c.btnAbNow.Add_Click({ Start-HUAppBackupNow -FromDialog })
    $c.btnAbCheck.Add_Click({ Start-HUAppBackupOpen -CheckOnly })
    $c.btnAbRestore.Add_Click({ Start-HUAppBackupOpen })
    $c.btnAbOpen.Add_Click({
            $p = [Environment]::ExpandEnvironmentVariables("$($script:AbDlg.C.txtAbPath.Text)".Trim())
            if ($p -and (Test-Path -LiteralPath $p)) { Start-Process explorer.exe -ArgumentList "`"$p`"" } else { Show-HUMessage 'Der Sicherungsordner existiert (noch) nicht.' -Icon Info -Owner $script:AbDlg.W }
        })
}

# Kopfzeile im Hauptfenster neben dem Titel: Diskette, Tooltip mit letzter Sicherung; laeuft = blau mit Prozent, Fehler = orange.
# Nur Verwalter sehen sie (Abweichung von der Vorlage: die Sicherung stellt ein Verwalter ein).
function Update-HUAbMainStatus {
    $t = $null; try { $t = $script:Controls['txtAppBackupStatus'] } catch { }
    if (-not $t) { return }
    if (-not (Test-HUVerwalter)) { $t.Visibility = 'Collapsed'; return }
    $t.Visibility = 'Visible'
    $ic = [char]::ConvertFromUtf32(0x1F4BE)
    $run = Get-HUAbProgressText
    $cfg = Get-HUAbConfig
    $last = "$($cfg.Last)"; $res = "$($cfg.Result)"
    $lt = $(if ($last) { try { ([datetime]::Parse($last, [Globalization.CultureInfo]::InvariantCulture)).ToString('dd.MM.yyyy HH:mm') } catch { $last } } else { '' })
    $by = $(if ($cfg.LastBy) { " von $($cfg.LastBy)" } else { '' })
    $mode = switch ("$($cfg.Mode)") { 'taeglich' { 'täglich' } 'woechentlich' { 'wöchentlich' } default { 'aus' } }
    if ($run) {
        $p = $script:AbProg; $pc = $(if ($p -and [double]$p.Total -gt 0) { " $([int](100 * [double]$p.Done / [double]$p.Total)) %" } else { ' ...' })
        $t.Text = "$ic$pc"; $t.Foreground = [System.Windows.Media.Brushes]::LightSkyBlue; $t.ToolTip = "App-Sicherung $run"
    } elseif ($res -like 'FEHLER*' -or $res -like 'Automatik:*') {
        $t.Text = "$ic !"; $t.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#FF9800'); $t.ToolTip = "App-Sicherung: $res$(if ($lt) { "`nLetzte erfolgreiche: $lt$by" })`nKlick: Einstellungen > Sicherung"
    } else {
        $t.Text = "$ic"; $t.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#858585')
        $t.ToolTip = $(if ($lt) { "App-Sicherung: zuletzt $lt$by$(if ($res) { " ($res)" })`nAutomatisch: $mode`nKlick: Einstellungen > Sicherung" } else { "Noch keine App-Sicherung`nKlick: Einstellungen > Sicherung" })
    }
}

function Register-HUAbMainStatus {
    $t = $null; try { $t = $script:Controls['txtAppBackupStatus'] } catch { }
    if (-not $t) { return }
    $t.Add_MouseLeftButtonUp({ Open-HUSettings -Tab 'Sicherung' })
    Update-HUAbMainStatus
}

function Update-HUAbTabInfo {
    Update-HUAbMainStatus
    if (-not $script:AbDlg) { return }
    $c = $script:AbDlg.C
    $stored = Test-Path -LiteralPath (Join-Path (Get-HUAbDataRoot) 'sicherung-kennwort.dat')
    $c.txtAbPwInfo.Text = $(if ($stored) { 'gespeichert (Automatik möglich)' } elseif ($script:AbDlg.Pw) { 'nur für diese Sitzung' } else { 'noch keines festgelegt' })
    $c.btnAbPw.Content = $(if ($stored -or $script:AbDlg.Pw) { 'Kennwort ändern ...' } else { 'Kennwort festlegen ...' })
    $run = Get-HUAbProgressText
    foreach ($b in 'btnAbNow', 'btnAbCheck', 'btnAbRestore') { $c[$b].IsEnabled = -not $run }
    $c.btnAbNow.Content = $(if ($run -like 'sichert*') { 'Sichert ...' } else { 'Jetzt sichern' })
    if ($run) { $c.txtAbLast.Text = $run; $c.txtAbLast.Foreground = [System.Windows.Media.Brushes]::LightSkyBlue; return }
    $c.txtAbLast.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#AAAAAA')
    $cfg = Get-HUAbConfig
    $last = "$($cfg.Last)"; $res = "$($cfg.Result)"
    $c.txtAbLast.Text = $(if ($last) { "$(try { ([datetime]::Parse($last, [Globalization.CultureInfo]::InvariantCulture)).ToString('dd.MM.yyyy HH:mm') } catch { $last })$(if ($cfg.LastBy) { " von $($cfg.LastBy)" })$(if ($res) { "  -  $res" })" } elseif ($res) { $res } else { 'noch keine' })
}

# beim Speichern der Einstellungen (Rueckgabe: Fehlertext oder '')
function Save-HUAbSettingsFromForm($c) {
    $keep = 0
    if (-not [int]::TryParse("$($c.txtAbKeep.Text)".Trim(), [ref]$keep) -or $keep -lt 1 -or $keep -gt 999) { return 'Sicherung: aufbewahren als Zahl zwischen 1 und 999.' }
    $mode = "$($c.cmbAbMode.SelectedItem.Tag)"
    $path = "$($c.txtAbPath.Text)".Trim()
    if ($mode -ne 'aus' -and -not $path) { return 'Sicherung: für die automatische Sicherung bitte einen Ordner angeben.' }
    if ($path) { $why = Test-HUAbFolder $path; if ($why) { return "Sicherung: $why" } }
    $cfg = Get-HUAbConfig
    if ("$($cfg.Folder)" -eq $path -and "$($cfg.Mode)" -eq $mode -and [int]$cfg.Keep -eq $keep) { return '' }
    $cfg.Folder = $path; $cfg.Mode = $mode; $cfg.Keep = $keep
    $why = Save-HUAbConfig $cfg
    if ($why) { return "Sicherung nicht gespeichert:`n$why" }
    return ''
}

# ---------------------------------------------------------------- Sichern ----------------------------------------------------------------
# -FromDialog: Werte aus dem offenen Einstellungs-Dialog (auch ungespeichert), Kennwort notfalls abfragen.
# -Auto: aus den Einstellungen, nur mit gespeichertem Kennwort, ohne Dialoge.
function Start-HUAppBackupNow([switch]$FromDialog, [switch]$Auto) {
    $owner = $(if ($FromDialog -and $script:AbDlg) { $script:AbDlg.W } else { $script:Window })
    if ($FromDialog -and $script:AbDlg) {
        $c = $script:AbDlg.C
        $folder = [Environment]::ExpandEnvironmentVariables("$($c.txtAbPath.Text)".Trim())
        $keep = 10; [void][int]::TryParse("$($c.txtAbKeep.Text)".Trim(), [ref]$keep); if ($keep -lt 1) { $keep = 10 }
    } else {
        $cfg = Get-HUAbConfig
        $folder = $cfg.Path; $keep = $cfg.Keep
    }
    if (-not $folder) { if (-not $Auto) { Show-HUMessage 'Bitte zuerst einen Sicherungsordner angeben.' -Icon Warning -Owner $owner }; return }
    $why = Test-HUAbFolder $folder
    if ($why) { if (-not $Auto) { Show-HUMessage $why -Icon Warning -Owner $owner }; return }
    # "Jetzt sichern" uebernimmt Ordner und Automatik gleich in die Einstellungen -
    # sonst laeuft die Automatik nie, wenn das Fenster danach mit Abbrechen geschlossen wird
    if ($FromDialog -and $script:AbDlg) {
        $why = Save-HUAbSettingsFromForm $script:AbDlg.C
        if ($why) { Show-HUMessage $why -Icon Warning -Owner $owner; return }
        Write-HULogInfo 'App-Sicherung: Ordner und Automatik gespeichert.'
    }
    $dr = Get-HUAbDataRoot
    $pw = $(if ($FromDialog -and $script:AbDlg -and $script:AbDlg.Pw) { $script:AbDlg.Pw } else { Get-HUAppBackupPassword -DataRoot $dr })
    if (-not $pw) {
        if ($Auto) {
            # nur Verwalter bekommen den Hinweis (andere stellen die Sicherung nicht ein); hoechstens einmal am Tag
            if (-not (Test-HUVerwalter)) { return }
            Write-HULogWarn 'App-Sicherung (automatisch): kein gespeichertes Kennwort - übersprungen.'
            $today = (Get-Date).ToString('yyyy-MM-dd')
            if ("$(Get-HUStateValue 'appBackupNoPwHint' '')" -ne $today) {
                Set-HUStateValue 'appBackupNoPwHint' $today
                Show-HUToast -Title 'App-Sicherung übersprungen' -Text "Kein Kennwort für die automatische Sicherung gespeichert.`nKlick: Einstellungen > Sicherung" -Icon Warning -OnClick { Open-HUSettings -Tab 'Sicherung' }
            }
            Update-HUAbMainStatus; return
        }
        $r = Show-HUAbPasswordDialog $owner -New
        if (-not $r) { return }
        $pw = $r.Password
        if ($script:AbDlg) { $script:AbDlg.Pw = $pw }
        if ($r.Remember) { try { Save-HUAppBackupPassword -DataRoot $dr -Password $pw } catch { } }
        Update-HUAbTabInfo
    }
    if ((Test-HUJobRunning 'AppBackup') -or (Test-HUJobRunning 'AppBackupOpen')) { if (-not $Auto) { Show-HUMessage 'Es läuft bereits eine App-Sicherung bzw. Prüfung.' -Icon Info -Owner $owner }; return }
    # gemeinsame Datei sperren: im Netzordner sichert immer nur einer
    [void](Get-HUAbConfig)
    $l = Lock-HUDataFile -Path $script:AbSettingsFile -UserName $script:UserName
    if (-not $l.Ok) { if (-not $Auto) { Show-HUMessage "Es läuft gerade eine Sicherung.`n`n$(Get-HULockInfoText $l)" -Icon Info -Owner $owner }; return }
    $script:AbLock = $l
    $src = @()
    try { $src = @(Get-HUAppBackupSources -DataRoot $script:DataRoot) }
    catch { Unlock-HUDataFile $l; $script:AbLock = $null; if (-not $Auto) { Show-HUMessage "Dateien nicht ermittelbar: $($_.Exception.Message)" -Icon Error -Owner $owner }; return }
    $mods = @(Get-HUAbModuleList @((Join-Path $script:AppRoot 'Modules'), (Join-Path $script:AppRoot 'Addons')))
    $per = (@($src | Group-Object Area | Sort-Object Count -Descending | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', ')
    Write-HULogInfo "App-Sicherung startet im Hintergrund$(if ($Auto) { ' (automatisch)' }): $($src.Count) Dateien ($per) ..."
    $script:AbAuto = [bool]$Auto
    $script:AbProg = [hashtable]::Synchronized(@{ Phase = 'Start'; Done = 0; Total = $src.Count })
    $ok = Start-HUJob -Name 'AppBackup' -Quiet -Vars @{ Src = $src; Pw = $pw; Folder = $folder; Ver = "$($script:Version)"; Keep = $keep; Mods = $mods; Prog = $script:AbProg } -Code {
        try {
            $r = New-HUAppBackup -Sources $Src -Password $Pw -Folder $Folder -AppVersion $Ver -Keep $Keep -Modules $Mods -Progress $Prog
            [pscustomobject]@{ __Ab = $true; Ok = $true; Path = $r.Path; Files = $r.Files; Bytes = $r.Bytes; Removed = @($r.Removed); Error = '' }
        } catch { [pscustomobject]@{ __Ab = $true; Ok = $false; Error = "$($_.Exception.Message)" } }
    } -OnDone {
        param($Result, $Errors)
        $r = @($Result | Where-Object { $_ -and $_.PSObject.Properties['__Ab'] })[0]
        $auto = [bool]$script:AbAuto
        try {
            if ($r -and $r.Ok) {
                $mb = [Math]::Round($r.Bytes / 1MB, 1)
                Set-HUAbResult "ok, $($r.Files) Dateien ($mb MB)$(if ($auto) { ', automatisch' })" -Success
                Write-HULogOK "App-Sicherung: $(Split-Path -Leaf $r.Path) ($($r.Files) Dateien, $mb MB)$(if (@($r.Removed).Count) { " - alte entfernt: $(@($r.Removed).Count)" })"
                if (-not $auto) { Show-HUMessage "Sicherung erstellt:`n$($r.Path)`n`n$($r.Files) Dateien, $mb MB." -Owner $(if ($script:AbDlg) { $script:AbDlg.W } else { $script:Window }) }
                else { Show-HUToast -Title 'App-Sicherung erstellt' -Text "$($r.Files) Dateien, $mb MB - $(Split-Path -Leaf $r.Path)" -Icon Success -OnClick { Open-HUSettings -Tab 'Sicherung' } }
            } else {
                $err = $(if ($r) { $r.Error } else { (@($Errors) -join ' ') })
                Set-HUAbResult "FEHLER: $err"
                Write-HULogError "App-Sicherung fehlgeschlagen: $err"
                if (-not $auto) { Show-HUMessage "Sicherung fehlgeschlagen:`n$err" -Icon Error -Owner $(if ($script:AbDlg) { $script:AbDlg.W } else { $script:Window }) }
                else { Show-HUToast -Title 'App-Sicherung fehlgeschlagen' -Text "$err`nKlick: Einstellungen > Sicherung" -Icon Error -OnClick { Open-HUSettings -Tab 'Sicherung' } }
            }
        } finally {
            if ($script:AbLock) { Unlock-HUDataFile $script:AbLock; $script:AbLock = $null }
            $script:AbProg = $null
            Update-HUAbTabInfo
        }
    }
    if (-not $ok) { Unlock-HUDataFile $l; $script:AbLock = $null; $script:AbProg = $null; return }
    Start-HUAbProgressTimer
    Update-HUAbTabInfo
}

# Anzeige waehrend der Sicherung (Reiter Sicherung): Knopf gesperrt, Fortschritt in "Letzte Sicherung"
function Start-HUAbProgressTimer {
    if ($script:AbProgTimer) { $script:AbProgTimer.Start(); return }
    $t = [System.Windows.Threading.DispatcherTimer]::new()
    $t.Interval = [TimeSpan]::FromMilliseconds(500)
    $t.Add_Tick({
            Update-HUAbTabInfo
            if (-not $script:AbProg -and -not (Test-HUJobRunning 'AppBackup') -and -not (Test-HUJobRunning 'AppBackupOpen')) { $this.Stop() }
        })
    $script:AbProgTimer = $t
    $t.Start()
}

function Get-HUAbProgressText {
    $p = $script:AbProg
    if (-not $p -or -not ((Test-HUJobRunning 'AppBackup') -or (Test-HUJobRunning 'AppBackupOpen'))) { return '' }
    $t = [double]$p.Total; $d = [double]$p.Done
    $pc = $(if ($t -gt 0) { " ($([int](100 * $d / $t)) %)" } else { '' })
    switch ("$($p.Phase)") {
        'Dateien' { return "sichert: $([int]$d) von $([int]$t) Dateien$pc" }
        'Verschluesseln' { return 'sichert: verschlüsseln ...' }
        'Pruefen' { return 'sichert: zurücklesen und prüfen ...' }
        'Schluessel' { return 'prüft: Schlüssel aus dem Kennwort ...' }
        'Pruefsumme' { return "prüft: Prüfsumme$pc" }
        'Entschluesseln' { return "prüft: entschlüsseln$pc" }
        'Lesen' { return "prüft: $([int]$d) von $([int]$t) Dateien lesen$pc" }
        'Vergleichen' { return "prüft: Prüfsummen der Dateien$pc" }
        default { return 'läuft ...' }
    }
}

# Automatik: beim Start (verzoegert) und dann alle 15 Minuten pruefen
function Start-HUAppBackupScheduler {
    $check = {
        try {
            Update-HUAbMainStatus
            $cfg = Get-HUAbConfig
            if (-not $cfg.Path -or -not (Test-HUAppBackupDue -Mode $cfg.Mode -Last "$($cfg.Last)")) { return }
            # nach einem Fehlschlag (z. B. Netzlaufwerk weg) hoechstens einmal pro Stunde neu versuchen
            $try = $null; try { $try = [datetime]::Parse("$(Get-HUStateValue 'appBackupTry' '')", [Globalization.CultureInfo]::InvariantCulture) } catch { }
            if ($try -and ((Get-Date) - $try).TotalMinutes -lt 60) { return }
            Set-HUStateValue 'appBackupTry' ((Get-Date).ToString('s'))
            Start-HUAppBackupNow -Auto
        } catch { Write-HULogWarn "App-Sicherung (automatisch): $($_.Exception.Message)" }
    }
    $script:AbCheck = $check
    Invoke-HUDelayed 90 { & $script:AbCheck }
    $t = [System.Windows.Threading.DispatcherTimer]::new()
    # alle 15 Minuten pruefen: bei durchlaufender App kommt die taegliche Sicherung so kurz nach Mitternacht
    $t.Interval = [TimeSpan]::FromMinutes(15)
    $t.Add_Tick({ & $script:AbCheck })
    $t.Start()
    $script:AbTimer = $t
}

# ---------------------------------------------------------------- Pruefen / Wiederherstellen ----------------------------------------------------------------
function Start-HUAppBackupOpen([switch]$CheckOnly) {
    $owner = $(if ($script:AbDlg) { $script:AbDlg.W } else { $script:Window })
    # Wiederherstellen ueberschreibt gemeinsame Daten: erst, wenn niemand sonst die App offen hat
    if (-not $CheckOnly) {
        $others = Get-HUOtherSessionsText
        if ($others) { Show-HUMessage "Wiederherstellen geht erst, wenn alle anderen $($script:AppTitle) geschlossen haben.`n`nBitte zuerst schließen:`n$others" -Icon Warning -Owner $owner; return }
    }
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'HU-Sicherung (*.husicherung)|*.husicherung'
    $dlg.Title = $(if ($CheckOnly) { 'Sicherung prüfen' } else { 'Sicherung wiederherstellen' })
    $dir = $(if ($script:AbDlg) { [Environment]::ExpandEnvironmentVariables("$($script:AbDlg.C.txtAbPath.Text)".Trim()) } else { (Get-HUAbConfig).Path })
    if ($dir -and (Test-Path -LiteralPath $dir)) { $dlg.InitialDirectory = $dir }
    if (-not $dlg.ShowDialog($owner)) { return }
    $file = $dlg.FileName
    # Kopf (nicht authentisiert) nur zur Anzeige
    $info = ''
    try {
        $fs = [IO.File]::OpenRead($file)
        try { $h = (Read-HUAppBackupHeaderStream $fs).Header; $info = "Sicherung vom $(try { ([datetime]::Parse("$($h.Created)", [Globalization.CultureInfo]::InvariantCulture)).ToLocalTime().ToString('dd.MM.yyyy HH:mm') } catch { $h.Created }), $($script:AppTitle) $($h.App), $($h.Files) Dateien." } finally { $fs.Dispose() }
    } catch { Show-HUMessage "$($_.Exception.Message)" -Icon Error -Owner $owner; return }
    $r = Show-HUAbPasswordDialog $owner -Text "$info`n`nKennwort der Sicherung eingeben."
    if (-not $r) { return }
    if ((Test-HUJobRunning 'AppBackup') -or (Test-HUJobRunning 'AppBackupOpen')) { Show-HUMessage 'Es läuft bereits eine App-Sicherung bzw. Prüfung.' -Icon Info -Owner $owner; return }
    # im Hintergrund pruefen und lesen (grosse Sicherungen brauchen etwas) - Fortschritt in Kopfzeile und Reiter
    $script:AbOpen = @{ File = $file; Info = $info; CheckOnly = [bool]$CheckOnly; Owner = $owner }
    $script:AbProg = [hashtable]::Synchronized(@{ Phase = 'Schluessel'; Done = 0; Total = 0 })
    Write-HULogInfo "App-Sicherung wird $(if ($CheckOnly) { 'geprüft' } else { 'zum Wiederherstellen gelesen' }): $(Split-Path -Leaf $file) ..."
    [void](Start-HUJob -Name 'AppBackupOpen' -Quiet -Vars @{ F = $file; Pw = $r.Password; Ver = "$($script:Version)"; Dr = $script:DataRoot; Prog = $script:AbProg } -Code {
            try { $o = Open-HUAppBackup -Path $F -Password $Pw -AppVersion $Ver -AppRoot $AppRoot -DataRoot $Dr -Progress $Prog; [pscustomobject]@{ __AbO = $true; O = $o; Error = '' } }
            catch { [pscustomobject]@{ __AbO = $true; O = $null; Error = "$($_.Exception.Message)" } }
        } -OnDone {
            param($Result, $Errors)
            $script:AbProg = $null
            Update-HUAbTabInfo
            $x = @($Result | Where-Object { $_ -and $_.PSObject.Properties['__AbO'] })[0]
            $a = $script:AbOpen; $script:AbOpen = $null
            if (-not $a) { return }
            if (-not $x -or -not $x.O) { Show-HUMessage "Sicherung nicht lesbar: $(if ($x) { $x.Error } else { @($Errors) -join ' ' })" -Icon Error -Owner $a.Owner; return }
            Complete-HUAppBackupOpen $x.O $a
        })
    Start-HUAbProgressTimer
    Update-HUAbTabInfo
}

# Fortsetzung nach dem Lesen im Hintergrund: Pruefergebnis zeigen bzw. Bereiche waehlen und schreiben
function Complete-HUAppBackupOpen($o, $a) {
    $owner = $a.Owner; $file = $a.File; $info = $a.Info; $CheckOnly = $a.CheckOnly
    if ($o.Errors.Count) { foreach ($er in @($o.Errors)) { Write-HULogWarn "App-Sicherung $(Split-Path -Leaf $file): $er" }; Show-HUMessage "Sicherung nicht in Ordnung - nichts geschrieben:`n`n$((@($o.Errors) | Select-Object -First 10) -join "`n")$(if (@($o.Errors).Count -gt 10) { "`n... und $(@($o.Errors).Count - 10) weitere (Protokoll)" })" -Icon Error -Owner $owner; return }
    $areas = @($o.Rows | Group-Object Area | Sort-Object Name | ForEach-Object { "  $($_.Name): $($_.Count) Datei(en)" })
    $mods = @($o.Manifest.Modules | Where-Object { $_ } | ForEach-Object { "$($_.Title) $($_.Version)" })
    if ($CheckOnly) {
        Write-HULogOK "App-Sicherung geprüft: $(Split-Path -Leaf $file) - in Ordnung ($(@($o.Rows).Count) Dateien)."
        Show-HUMessage "Sicherung in Ordnung (Kennwort, Prüfsummen, Inhalt).`n`n$info`n`n$($areas -join "`n")$(if ($mods.Count) { "`n`nModule: $($mods -join ', ')" })`n`nEs wurde nichts geschrieben." -Owner $owner
        return
    }
    $sel = Show-HUAbRestorePick $o $owner
    if (-not $sel -or -not @($sel).Count) { return }
    $others = Get-HUOtherSessionsText
    if ($others) { Show-HUMessage "Wiederherstellen geht erst, wenn alle anderen $($script:AppTitle) geschlossen haben.`n`nBitte zuerst schließen:`n$others" -Icon Warning -Owner $owner; return }
    if (-not (Confirm-HU "Diese Bereiche jetzt wiederherstellen?`n`n$(@($sel) -join "`n")`n`nDer aktuelle Stand der betroffenen Dateien wird vorher unter`n$(Join-Path $script:DataRoot 'Wiederherstellung')`nabgelegt. Danach $($script:AppTitle) neu starten." 'Wiederherstellen' -Warning -Owner $owner)) { return }
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    try {
        $w = Restore-HUAppBackup -Opened $o -Areas @($sel) -AppRoot $script:AppRoot -DataRoot $script:DataRoot
    } catch { [System.Windows.Input.Mouse]::OverrideCursor = $null; Show-HUMessage "Wiederherstellen fehlgeschlagen: $($_.Exception.Message)" -Icon Error -Owner $owner; return }
    [System.Windows.Input.Mouse]::OverrideCursor = $null
    Write-HULogOK "App-Sicherung wiederhergestellt: $(Split-Path -Leaf $file) - $($w.Written) Dateien ($(@($sel) -join ', ')). Vorstand: $($w.Before)"
    # Hinweise fuer danach
    $notes = New-Object System.Collections.Generic.List[string]
    $inst = @(Get-HUAbModuleList @((Join-Path $script:AppRoot 'Modules'), (Join-Path $script:AppRoot 'Addons')) | ForEach-Object { "$($_.Id)" })
    $miss = @(@($w.Modules) | Where-Object { $_ -and $inst -notcontains "$($_.Id)" } | ForEach-Object { "$($_.Title) $($_.Version)" })
    if ($miss.Count) { $notes.Add("Module hier nicht installiert (Einstellungen > Module): $($miss -join ', ')") }
    $notes.Add('Das Kennwort für die automatische Sicherung ist nicht in der Sicherung - bei Bedarf neu festlegen.')
    # ab jetzt nichts mehr aus dem Speicher zurueckschreiben (Fensterzustand beim Beenden)
    $script:AbRestored = $true
    # immer neu starten: jeder offene Stand im Speicher wuerde sonst den wiederhergestellten ueberschreiben
    Show-HUMessage "$($w.Written) Datei(en) wiederhergestellt.`n`n$(@($notes | ForEach-Object { "- $_" }) -join "`n")`n`n$($script:AppTitle) startet jetzt neu, damit der wiederhergestellte Stand gilt." 'Wiederherstellen' -Owner $owner
    Write-HULogInfo "Wiederhergestellt - $($script:AppTitle) startet neu."
    if ($script:AbDlg) { try { $script:AbDlg.W.Close() } catch { } }
    # erst nach dem Schliessen des Einstellungsfensters (modal) neu starten
    Invoke-HUDelayed 1 { Restart-HUApp }
}

# Auswahl der Bereiche. Rueckgabe: Bereichsnamen oder $null.
function Show-HUAbRestorePick($Opened, $Owner) {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Wiederherstellen - Bereiche w&#xE4;hlen" Width="860" Height="460" MinWidth="620" MinHeight="300" WindowStartupLocation="CenterOwner" Background="#1E1E1E" ShowInTaskbar="False" ResizeMode="CanResizeWithGrip">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel Margin="14">
        <TextBlock DockPanel.Dock="Top" Style="{StaticResource HintText}" TextWrapping="Wrap" Margin="0,0,0,8"
                   Text="Haken = wird aus der Sicherung zur&#xFC;ckgeschrieben (&#xFC;berschreibt den aktuellen Stand). Nicht angehakte Bereiche bleiben, wie sie sind."/>
        <TextBlock x:Name="txtCode" DockPanel.Dock="Top" TextWrapping="Wrap" Margin="0,0,0,8" Foreground="#FFB74D"
                   Text="Orange = legt fest, wer Verwalter ist. Dieser Bereich ist nicht vorgew&#xE4;hlt: Verwalter sehen Protokoll, Sicherung, Update und Module. Nur anhaken, wenn die Sicherung von dir selbst stammt."/>
        <CheckBox x:Name="chkAll" DockPanel.Dock="Top" Content="Alle" IsChecked="False" Style="{StaticResource DarkCheckBox}" Margin="6,0,0,6" ToolTip="Alle Haken setzen bzw. entfernen"/>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnOk" Content="Weiter" Width="110" Background="#1976D2" Style="{StaticResource DarkButton}" IsDefault="True" Margin="0,0,8,0" ToolTip="Gewählte Bereiche wiederherstellen (fragt noch einmal nach)"/>
            <Button x:Name="btnCancel" Content="Abbrechen" Width="100" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Nichts ändern"/>
        </StackPanel>
        <DataGrid x:Name="grid" Style="{StaticResource DarkDataGrid}" AutoGenerateColumns="False" IsReadOnly="False" HeadersVisibility="Column" SelectionMode="Single" CanUserAddRows="False"
                  ToolTip="Bereiche in der Sicherung">
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Pick, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" Style="{StaticResource DarkCheckBox}"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Bereich">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <TextBlock Text="{Binding Area}" Foreground="{Binding Color}" VerticalAlignment="Center" Margin="4,0"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Dateien" Binding="{Binding Count}" IsReadOnly="True"/>
                <DataGridTextColumn Header="Stand in der Sicherung" Binding="{Binding Saved}" IsReadOnly="True"/>
                <DataGridTextColumn Header="Aktuell hier" Binding="{Binding Current}" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $fmt = { param($v) if ("$v") { try { ([datetime]::Parse("$v", [Globalization.CultureInfo]::InvariantCulture)).ToString('dd.MM.yyyy HH:mm') } catch { "$v" } } else { '' } }
    $rows = foreach ($g in @($Opened.Rows | Group-Object Area | Sort-Object Name)) {
        $sv = @($g.Group | ForEach-Object { "$($_.Changed)" } | Where-Object { $_ } | Sort-Object -Descending)[0]
        $ex = @($g.Group | Where-Object Exists)
        $cv = @($ex | ForEach-Object { "$($_.CurChanged)" } | Where-Object { $_ } | Sort-Object -Descending)[0]
        $code = (@($g.Group | Where-Object { Test-HUAppBackupCodeName "$($_.Name)" }).Count -gt 0)
        [pscustomobject]@{ Pick = (-not $code); Code = $code; Color = $(if ($code) { '#FFB74D' } else { '#E0E0E0' }); Area = $g.Name; Count = $g.Count; Saved = (& $fmt $sv); Current = $(if ($ex.Count) { "$($ex.Count) vorhanden, neueste $(& $fmt $cv)" } else { 'nicht vorhanden' }) }
    }
    $c.grid.ItemsSource = @($rows)
    $st = @{ R = $null; Busy = $false }
    $c.chkAll.IsChecked = (-not @($rows | Where-Object { -not $_.Pick }).Count)
    if (-not @($rows | Where-Object { $_.Code }).Count) { $c.txtCode.Visibility = 'Collapsed' }
    # "Alle": setzt bzw. entfernt alle Haken (ein Klick je Zeile genuegt - Haken direkt in der Zelle)
    $setAll = { param([bool]$v) if ($st.Busy) { return }; foreach ($r in @($c.grid.ItemsSource)) { $r.Pick = $v }; $c.grid.Items.Refresh() }
    $c.chkAll.Add_Checked({ & $setAll $true })
    $c.chkAll.Add_Unchecked({ & $setAll $false })
    $c.btnOk.Add_Click({ $c.grid.CommitEdit(); $st.R = @(@($c.grid.ItemsSource) | Where-Object { $_.Pick } | ForEach-Object { $_.Area }); $w.Close() })
    [void]$w.ShowDialog()
    return $st.R
}
