#Requires -Version 5.1
<#
.SYNOPSIS
    Einstellungen wie in der Vorlage: Allgemein (Verknuepfungen, Start und Darstellung), Sicherung, Update.
    Dazu (nur HU-Amt) Module und Verwalter. Verwalter-Ansicht im Hauptfenster.
.DESCRIPTION
    Persoenlich:  Data\Benutzer\<Name>\settings.json (ui.startTab, ui.splashAnimation, ui.checkUpdatesOnStart, logging.logLevel)
                  Data\Benutzer\<Name>\ui-state.json (Oberflaechengroesse)
    Installation: Config\update.json (Kanal, Signaturpflicht), Config\powershell-host.txt (auto | 5 | 7)
    Sicherung:    Data\sicherung.json (UI-AppBackup.ps1)
    Verwalter:    Data\Verwalter.txt (ein Windows-Benutzername je Zeile; mit Listen-Sperre gespeichert)
    Nur Verwalter: Reiter Sicherung, Module, Verwalter, PowerShell beim Start, Protokoll-Stufe, im Reiter Update Kanal,
    Signatur und Vorversion (gelten fuer die ganze Installation bzw. betreffen das Protokoll, das nur Verwalter sehen).
.NOTES
    Dot-Source aus Main.ps1. Zielmaschine: der PC, auf dem die App laeuft.
#>

function Set-HUProp($Object, [string]$Name, $Value) {
    if ($Object.PSObject.Properties[$Name]) { $Object.$Name = $Value } else { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
}
# Get-HUProp kommt aus Core\HU.Data.psm1 (Standard dort '', hier immer mit eigenem Standard aufgerufen)

function Save-HUUserSettings {
    Write-HUJsonFile -Path $script:SettingsPath -Object $script:Settings -Depth 5
}

# Verwalter-Liste schreiben (Regel 3: Listen-Sperre). Rueckgabe: '' oder Grund
function Save-HUVerwalterList([string[]]$Names) {
    $f = Get-HUVerwalterFile
    $l = Lock-HUDataFile -Path $f -UserName $script:UserName
    if (-not $l.Ok) { return (Get-HULockInfoText $l) }
    try {
        [void](Backup-HUDataFile -Path $f -Keep 10)
        $txt = "# Verwalter: ein Windows-Benutzername je Zeile`r`n" + ((@($Names) | Sort-Object -Unique) -join "`r`n") + "`r`n"
        [IO.File]::WriteAllText($f, $txt, (New-Object Text.UTF8Encoding $true))
        return ''
    } finally { Unlock-HUDataFile $l }
}

# Verwalter-Ansicht: Protokoll und Verwaltung nur fuer Verwalter (Konzept 1a)
function Update-HUAdminView {
    $adm = Test-HUVerwalter
    $v = $(if ($adm) { 'Visible' } else { 'Collapsed' })
    foreach ($n in 'pnlLog', 'thbLog') { if ($script:Controls[$n]) { $script:Controls[$n].Visibility = $v } }
    try { Update-HUAbMainStatus } catch { }
}

# Erststart: Gibt es noch keine Verwalter-Liste und hat niemand das Herausgeber-Zertifikat, sieht niemand die Verwaltung.
# Dann einmal fragen, ob dieser Benutzer Verwalter wird. Wer Nein sagt, wird nicht mehr gefragt (je Benutzer gemerkt).
# Sobald die Liste besteht, kommt die Frage fuer niemanden mehr. Rueckgabe: $true, wenn eingetragen.
function Test-HUFirstVerwalterNeeded {
    if ($null -ne (Get-HUVerwalterList)) { return $false }
    if (Test-HUVerwalter) { return $false }
    if ((Get-HUStateValue 'verwalterGefragt' $false) -eq $true) { return $false }
    return $true
}
function Request-HUFirstVerwalter {
    if (-not (Test-HUFirstVerwalterNeeded)) { return $false }
    $q = "In dieser Installation ist noch niemand Verwalter.`n`nVerwalter sehen Protokoll, Sicherung, Update-Einstellungen und Module und tragen weitere Verwalter ein.`n`nSoll $([char]0x201E)$($script:UserName)$([char]0x201C) Verwalter werden?"
    if (-not (Confirm-HU $q 'Verwalter' -Owner $script:Window)) {
        Set-HUStateValue 'verwalterGefragt' $true; Save-HUUIState
        Write-HULogInfo 'Verwalter: Erststart-Frage abgelehnt'
        return $false
    }
    $why = Save-HUVerwalterList @("$($script:UserName)")
    if ($why) { Show-HUMessage "Verwalter nicht gespeichert:`n$why" -Icon Warning -Owner $script:Window; return $false }
    $script:IsVerwalter = $null
    Update-HUAdminView
    Write-HULogOK "Verwalter beim Erststart eingetragen: $($script:UserName)"
    return $true
}

function Open-HUSettings([string]$Tab = '') {
    [void](Show-HUSettingsDialog -Tab $Tab)
}

# ---------------------------------------------------------------- Datenordner ----------------------------------------------------------------
# Anzeige im Reiter Verwalter
function Update-HUDataLocationInfo($Info) {
    try {
        $loc = Get-HUDataLocation -AppRoot $script:AppRoot
        $Info.Text = "Datenordner: $($loc.Path)$(if ($loc.Configured) { '' } else { '  (Standard: neben dem Programm)' })"
    } catch { $Info.Text = "Datenordner: $($_.Exception.Message)" }
}

<#
.SYNOPSIS
    Datenordner aendern (nur Verwalter, nicht im Testmodus). Leerer Ordner: alle Daten kopieren und pruefen, dann umstellen.
    Ordner mit Daten der App: nur umstellen. Anderer Inhalt: ablehnen. Danach Neustart.
.DESCRIPTION
    Vorher muessen alle anderen Fenster dieser Installation zu sein (Data\Laufend), sonst schreiben sie in den alten Ordner.
    Der alte Ordner bleibt als Stand erhalten und bekommt UMGEZOGEN.txt.
#>
function Invoke-HUDataLocationChange($Owner, $Info, $Button) {
    $t = 'Datenordner'
    if (Test-HUTestMode) { Show-HUMessage 'Im Testmodus lässt sich der Datenordner nicht ändern.' $t -Icon Warning -Owner $Owner; return }
    if (-not (Test-HUVerwalter)) { Show-HUMessage 'Den Datenordner können nur Verwalter ändern.' $t -Icon Warning -Owner $Owner; return }
    try { $loc = Get-HUDataLocation -AppRoot $script:AppRoot } catch { Show-HUMessage $_.Exception.Message $t -Icon Error -Owner $Owner; return }
    $cur = [IO.Path]::GetFullPath($loc.Path).TrimEnd('\')
    $def = [IO.Path]::GetFullPath($loc.Default).TrimEnd('\')
    $in = Read-HUText -Title $t -Owner $Owner -AllowEmpty -StateKey 'Datenordner' -Default $loc.Raw `
        -Prompt "Neuer Ordner für die Daten aller Module, am besten auf einer Freigabe als UNC-Pfad (\\server\freigabe\HU-Amt-Daten). Leer lassen = wieder neben dem Programm.`n`nJetzt: $cur"
    if ($null -eq $in) { return }
    try { $new = $(if ("$in".Trim()) { Resolve-HUDataLocationPath $in } else { $def }) } catch { Show-HUMessage $_.Exception.Message $t -Icon Warning -Owner $Owner; return }
    if ($new.Equals($cur, [StringComparison]::OrdinalIgnoreCase)) { Show-HUMessage "Das ist schon der aktuelle Datenordner:`n$cur" $t -Owner $Owner; return }
    if ($new -match '^[A-Za-z]:' -and -not $new.Equals($def, [StringComparison]::OrdinalIgnoreCase) -and
        -not (Confirm-HU -Warning "Laufwerksbuchstaben sind nicht auf jedem PC gleich verbunden. Sicherer ist ein UNC-Pfad (\\server\freigabe\...).`n`nTrotzdem $new verwenden?" $t -Owner $Owner)) { return }
    $others = @(Get-HURunningInstances -RunDir $script:Paths.RunDir -Self $script:RunLock)
    if ($others.Count) {
        Show-HUMessage ("Zuerst müssen alle anderen Fenster von $($script:AppTitle) geschlossen sein, sonst schreiben sie weiter in den alten Ordner:`n`n" +
            (($others | ForEach-Object { "  $($_.UserName) am $($_.ComputerName)" }) -join "`n")) $t -Icon Warning -Owner $Owner
        return
    }
    try { $state = Get-HUDataFolderState $new } catch { Show-HUMessage "Der Ordner lässt sich nicht lesen:`n$new`n`n$($_.Exception.Message)" $t -Icon Warning -Owner $Owner; return }
    if ($state -eq 'Fehlt') {
        if (-not (Confirm-HU "Den Ordner gibt es noch nicht:`n$new`n`nAnlegen?" $t -Owner $Owner)) { return }
        try { [void](New-Item -ItemType Directory -Path $new -Force -ErrorAction Stop); $state = 'Leer' } catch { Show-HUMessage "Ordner nicht angelegt:`n$($_.Exception.Message)" $t -Icon Error -Owner $Owner; return }
    }
    try { Assert-HUWritable $new } catch { Show-HUMessage "In diesem Ordner darf $($script:AppTitle) nicht schreiben:`n$new`n`n$($_.Exception.Message)" $t -Icon Warning -Owner $Owner; return }
    if ($state -eq 'Unvollstaendig') { Show-HUMessage "In diesem Ordner liegt eine unvollständige Kopie von einem abgebrochenen Umzug:`n$new`n`nBitte den Ordner leeren und es noch einmal versuchen." $t -Icon Warning -Owner $Owner; return }
    if ($state -eq 'Fremd') { Show-HUMessage "Im Ordner liegen schon andere Dateien. Bitte einen leeren Ordner wählen oder einen, in dem schon Daten von $($script:AppTitle) liegen:`n`n$new" $t -Icon Warning -Owner $Owner; return }
    $ctx = @{ Owner = $Owner; Info = $Info; Button = $Button; Cur = $cur; New = $new; Def = $def }
    if ($state -eq 'HU') {
        $adm = $false
        try { $adm = @([IO.File]::ReadAllLines((Join-Path $new 'Verwalter.txt')) | ForEach-Object { "$_".Trim() } | Where-Object { $_ -ieq $script:UserName -or $_ -ieq "$env:USERDOMAIN\$($script:UserName)" }).Count -gt 0 } catch { }
        $old = Test-Path -LiteralPath (Join-Path $new 'UMGEZOGEN.txt')
        $q = "Im Ordner liegen schon Daten von $($script:AppTitle):`n$new`n`nDiese verwenden? Es wird nichts kopiert, der jetzige Ordner bleibt, wie er ist." +
        $(if ($old) { "`n`nAchtung: Dieser Ordner ist als alter Stand gekennzeichnet (UMGEZOGEN.txt). Neuere Daten liegen woanders." }) +
        $(if (-not $adm) { "`n`nAchtung: Du stehst dort nicht in der Verwalter-Liste und siehst danach die Verwaltung nicht mehr." })
        if (-not (Confirm-HU -Warning $q $t -Owner $Owner)) { return }
        Complete-HUDataLocationChange $null '' $ctx -NoCopy
        return
    }
    $q = "Alle Daten von`n  $cur`nnach`n  $new`nkopieren, jede Datei prüfen und dann umstellen?`n`n" +
    "Der alte Ordner bleibt als Stand erhalten (Hinweis UMGEZOGEN.txt). Danach startet $($script:AppTitle) neu, auf den anderen PCs gilt der neue Ordner beim nächsten Start."
    if (-not (Confirm-HU $q $t -Owner $Owner)) { return }
    try { $Button.IsEnabled = $false; $Info.Text = 'Daten werden kopiert und geprüft …' } catch { }
    Write-HULogInfo "Datenordner: kopiere $cur nach $new"
    try {
        Start-HUBackgroundJob -Owner 'Datenordner' -LogPrefix 'Datenordner' -TimeoutSec 0 -Context $ctx -Arguments @($script:AppRoot, $cur, $new) -Script {
            param($AppRoot, $Src, $Dst)
            Import-Module (Join-Path $AppRoot 'Core\HU.Paths.psm1') -DisableNameChecking -ErrorAction Stop
            Copy-HUDataTree -Source $Src -Target $Dst
        } -Done { param($Res, $Err, $Ctx) Complete-HUDataLocationChange $Res $Err $Ctx }
    } catch { Complete-HUDataLocationChange $null "$($_.Exception.Message)" $ctx }
}

# Abschluss im Fenster-Thread: Einstellung schreiben, alten Ordner kennzeichnen, Neustart
function Complete-HUDataLocationChange($Res, $Err, $Ctx, [switch]$NoCopy) {
    $t = 'Datenordner'
    try { $Ctx.Button.IsEnabled = $true } catch { }
    if (-not $NoCopy) {
        if ($Err -or -not $Res) {
            try { Update-HUDataLocationInfo $Ctx.Info } catch { }
            Write-HULogError "Datenordner: Kopieren fehlgeschlagen ($Err)"
            Show-HUMessage "Kopieren fehlgeschlagen, es wurde nichts umgestellt:`n$Err`n`nIm Zielordner können Teile liegen. Vor einem neuen Versuch leeren." $t -Icon Error -Owner $Ctx.Owner
            return
        }
        $errs = @($Res.Errors)
        if ($errs.Count) {
            try { Update-HUDataLocationInfo $Ctx.Info } catch { }
            Write-HULogError "Datenordner: $($errs.Count) Dateien nicht kopiert, nichts umgestellt"
            Show-HUMessage ("Nicht alle Dateien ließen sich kopieren, es wurde nichts umgestellt:`n`n" + (($errs | Select-Object -First 8) -join "`n") +
                $(if ($errs.Count -gt 8) { "`n... und $($errs.Count - 8) weitere" }) + "`n`nIm Zielordner liegen Teile. Vor einem neuen Versuch leeren.") $t -Icon Error -Owner $Ctx.Owner
            return
        }
    }
    # unmittelbar vor dem Umstellen noch einmal: hat inzwischen jemand anderer HU-Amt geoeffnet? (Pruefung 090, KERN-B3)
    $others = @(Get-HURunningInstances -RunDir $script:Paths.RunDir -Self $script:RunLock)
    if ($others.Count) {
        # die Kopie ist nun veraltet: als unvollstaendig kennzeichnen, damit sie nicht spaeter ungefragt verwendet wird (Pruefung 091, KERN-B2)
        if (-not $NoCopy) { try { [IO.File]::WriteAllText((Join-Path $Ctx.New '.hu-kopie-unvollstaendig'), "Umzug abgebrochen $((Get-Date).ToString('s')): jemand hatte $($script:AppTitle) offen. Diese Kopie ist veraltet.") } catch { } }
        Show-HUMessage ("Inzwischen hat jemand anderer $($script:AppTitle) geöffnet. Es wurde nichts umgestellt$(if (-not $NoCopy) { ', die Kopie im neuen Ordner bleibt liegen' }).`n`n" +
            (($others | ForEach-Object { "  $($_.UserName) am $($_.ComputerName)" }) -join "`n") + "`n`nBitte alle anderen Fenster schließen lassen und es noch einmal versuchen$(if (-not $NoCopy) { ' (den neuen Ordner vorher leeren)' }).") $t -Icon Warning -Owner $Ctx.Owner
        return
    }
    try { Set-HUDataLocation -AppRoot $script:AppRoot -Path $(if ($Ctx.New -ieq $Ctx.Def) { '' } else { $Ctx.New }) }
    catch { Show-HUMessage "Einstellung nicht gespeichert, es wurde nichts umgestellt:`n$($_.Exception.Message)" $t -Icon Error -Owner $Ctx.Owner; return }
    if (-not $NoCopy) {
        try {
            $txt = "Die Daten von $($script:AppTitle) liegen seit $(Get-Date -Format 'dd.MM.yyyy HH:mm') in:`r`n$($Ctx.New)`r`n`r`n" +
            "Dieser Ordner ist ein alter Stand. Er wird nicht mehr benutzt und kann nach einer Prüfung gelöscht werden.`r`n"
            [IO.File]::WriteAllText((Join-Path $Ctx.Cur 'UMGEZOGEN.txt'), $txt, (New-Object System.Text.UTF8Encoding $true))
        } catch { Write-HULogWarn "Datenordner: UMGEZOGEN.txt nicht geschrieben ($($_.Exception.Message))" }
    }
    $info = $(if ($NoCopy) { 'Umgestellt.' } else { "$($Res.Files) Dateien ($([math]::Round([double]$Res.Bytes / 1MB, 1)) MB) kopiert und geprüft." })
    Write-HULogOK "Datenordner: $($Ctx.Cur) -> $($Ctx.New). $info"
    Show-HUMessage "$info`n`n$($script:AppTitle) startet jetzt neu und arbeitet mit`n$($Ctx.New)`n`nAuf den anderen PCs gilt der neue Ordner beim nächsten Start." $t -Owner $Ctx.Owner
    try { $Ctx.Owner.Close() } catch { }
    Restart-HUApp
}

function Show-HUSettingsDialog {
    param([string]$Tab = '')
    $adm = Test-HUVerwalter
    $d = New-HUWindow 'SettingsWindow'
    $w = $d.Window; $c = $d.C
    Restore-HUDialogState $w 'SettingsWindow'
    $S = $script:Settings
    $st = @{ Saved = $false; Names = (New-Object System.Collections.ArrayList); Changed = $false }

    # nur Verwalter: Reiter und Einstellungen, die fuer die ganze Installation gelten
    foreach ($n in 'tabAppBackup', 'tabModules', 'tabAdmins') { $c[$n].Visibility = $(if ($adm) { 'Visible' } else { 'Collapsed' }) }
    foreach ($n in 'pnlPsHost', 'lblLogLevel', 'cmbLogLevel') { $c[$n].Visibility = $(if ($adm) { 'Visible' } else { 'Collapsed' }) }
    foreach ($n in 'cmbChannel', 'chkRequireSig', 'btnOtherVersion') { $c[$n].IsEnabled = $adm }

    # ---------------- Allgemein ----------------
    $selTag = {
        param($cmb, [string]$tag)
        foreach ($i in $cmb.Items) { if ("$($i.Tag)" -eq $tag) { $cmb.SelectedItem = $i; return } }
        if ($cmb.Items.Count) { $cmb.SelectedIndex = 0 }
    }
    # Reiter beim Start: alle Reiter des Hauptfensters, dazu "Zuletzt verwendeter Reiter"
    foreach ($ti in @($script:Controls['tabMain'].Items | Where-Object { $_ -is [System.Windows.Controls.TabItem] -and $_.Visibility -eq 'Visible' })) {
        $it = New-Object System.Windows.Controls.ComboBoxItem; $it.Content = "$($ti.Header)"; $it.Tag = "$($ti.Name)"; [void]$c.cmbStartTab.Items.Add($it)
    }
    $it = New-Object System.Windows.Controls.ComboBoxItem; $it.Content = 'Zuletzt verwendeter Reiter'; $it.Tag = 'Last'; [void]$c.cmbStartTab.Items.Add($it)
    & $selTag $c.cmbStartTab (Get-HUProp $S.ui 'startTab' 'Last')
    $c.chkSplashAnim.IsChecked = [bool](Get-HUProp $S.ui 'splashAnimation' $true)
    # Demo (fuer alle): starten oder beenden, beides mit Neustart
    if (Test-HUDemoMode) { $c.btnDemo.Content = 'Demo beenden ...'; $c.btnDemo.ToolTip = 'Neu starten und wieder mit den echten Daten arbeiten' }
    elseif (Test-HUTestMode) { $c.btnDemo.IsEnabled = $false; $c.btnDemo.ToolTip = 'Im Testmodus nicht möglich. Zuerst den Testmodus beenden.' }
    $c.btnDemo.Add_Click({ $w.Close(); Switch-HUDemoMode })
    foreach ($p in @(80, 90, 100, 110, 120, 130, 140, 150)) {
        $it = New-Object System.Windows.Controls.ComboBoxItem; $it.Content = "$p %"; $it.Tag = "$($p / 100)"; [void]$c.cmbScale.Items.Add($it)
    }
    $curScale = [Math]::Round([double]$(if ($script:UiScale) { $script:UiScale } else { 1.0 }), 1)
    foreach ($i in $c.cmbScale.Items) { if ([Math]::Abs([double]::Parse("$($i.Tag)", [Globalization.CultureInfo]::InvariantCulture) - $curScale) -lt 0.01) { $c.cmbScale.SelectedItem = $i } }
    if (-not $c.cmbScale.SelectedItem) { $c.cmbScale.SelectedIndex = 2 }
    $ll = "$(Get-HUProp $S.logging 'logLevel' 'INFO')"
    & $selTag $c.cmbLogLevel $(if ($ll -match '^debug') { 'Debug' } elseif ($ll -match '^warn') { 'Warn' } else { 'Info' })

    $c.txtShortcutHint.Text = "Legt $($script:LauncherName) im Programmordner an (eigenes Logo, kein Konsolenfenster, an die Taskleiste anheftbar) und eine Verknüpfung darauf."
    $updShortcutInfo = {
        $exe = Join-Path $script:AppRoot $script:LauncherName
        $c.txtShortcutInfo.Text = "Starter: $(if (Test-Path -LiteralPath $exe) { $exe } else { 'noch nicht erstellt' })  |  Verknüpfung vorhanden: $(if (Test-HUShortcut) { 'ja' } else { 'nein' })"
    }
    & $updShortcutInfo
    # PowerShell beim Start (Config\powershell-host.txt)
    & $selTag $c.cmbPsHost (Get-HUPsHostSetting)
    $c.txtPsHostInfo.Text = "$(Get-HUPsHostInfoText) Änderung wirkt beim nächsten Start."
    $mkLink = {
        param([string]$Loc, [bool]$All)
        try {
            $l = New-HUShortcut -Location $Loc -AllUsers:$All
            Write-HULogOK "Verknüpfung erstellt: $l"
            Show-HUMessage "Verknüpfung erstellt:`n$l" -Owner $w
        } catch { Show-HUMessage "Verknüpfung fehlgeschlagen:`n$($_.Exception.Message)$(if ($All) { "`n`nFür alle Benutzer sind Administratorrechte nötig." })" -Icon Error -Owner $w }
        & $updShortcutInfo
    }
    $c.btnShortcutDesktop.Add_Click({ & $mkLink 'Desktop' $false })
    $c.btnShortcutStart.Add_Click({ & $mkLink 'StartMenu' $false })
    $c.btnShortcutAll.Add_Click({ & $mkLink 'Desktop' $true })
    $c.btnLauncher.Add_Click({
            if (New-HULauncher -Force) { Show-HUMessage "Starter bereit:`n$(Join-Path $script:AppRoot $script:LauncherName)" -Owner $w } else { Show-HUMessage 'Starter konnte nicht erstellt werden (Protokoll).' -Icon Error -Owner $w }
            & $updShortcutInfo
        })

    # ---------------- Sicherung ----------------
    if ($adm) { Initialize-HUAbSettingsTab $c $w }

    # ---------------- Update ----------------
    $ucfg = Update-UpdateConfig
    $uRaw = Read-HUJsonFile (Join-Path $script:ConfigDir 'update.json')
    & $selTag $c.cmbChannel $(if ($ucfg) { "$($ucfg.Channel)" } else { 'Stable' })
    $c.chkRequireSig.IsChecked = -not ($uRaw -and $uRaw.PSObject.Properties['AllowUnsigned'] -and $uRaw.AllowUnsigned -eq $true)
    $c.chkCheckOnStart.IsChecked = [bool](Get-HUProp $S.ui 'checkUpdatesOnStart' $true)
    $inst = Get-HUInstalledInfo
    $c.txtUpdateInfo.Text = "Installiert: v$($script:Version)$(if ($inst -and $inst.PSObject.Properties['Date']) { " (geladen $($inst.Date), $(Get-HUProp $inst 'Check' ''))" } else { ' (nicht über Pull.ps1 installiert)' })`nQuelle: github.com/$(if ($ucfg) { "$($ucfg.Owner)/$($ucfg.Repo)" })`nSignierte Releases werden mit dem Zertifikat des Herausgebers geprüft (Fingerabdruck im Programm). Eigene Quelle: Config\update.json (Owner, Repo, SignerThumbprint).$(if (-not $adm) { "`nKanal und Signatur stellt ein Verwalter ein." })"
    $c.btnCheckNow.Add_Click({ Invoke-UpdateCheck; Show-HUMessage 'Die Prüfung läuft im Hintergrund - bei neuer Version wird der Knopf Update gold.' -Owner $w })
    $c.btnOtherVersion.Add_Click({ $w.Close(); Show-HUVersionPicker })

    # ---------------- Module ----------------
    if ($adm) {
        $c.gridModules.ItemsSource = @($script:LoadedAddons | ForEach-Object {
                [pscustomobject]@{ Reiter = "$($_.Title)"; Id = $_.Id; Version = "$($_.Version)"; Herkunft = "$($_.Source)"; Ziehen = $(if (@($_.Drop).Count) { @($_.Drop) -join ' ' } else { '-' }) } })
    }
    $c.btnAddonRepos.Add_Click({ Show-HUAddonDialog })
    $c.btnAddonUpdate.Add_Click({ Start-HUAddonUpdate -Loud })
    $c.btnAddonDev.Add_Click({ $p = Join-Path $script:AppRoot 'AddonsDev'; if (-not (Test-Path -LiteralPath $p)) { [void](New-Item -ItemType Directory -Path $p -Force) }; Open-HUPath $p })

    # ---------------- Verwalter ----------------
    $list = Get-HUVerwalterList
    if ($list) { foreach ($n in $list) { [void]$st.Names.Add($n) } }
    $fillAdm = {
        $c.lstAdmins.Items.Clear(); foreach ($n in $st.Names) { [void]$c.lstAdmins.Items.Add($n) }
        $c.txtAdminInfo.Text = $(if ($st.Names.Count) { "$($st.Names.Count) Verwalter. Gespeichert wird mit $([char]0x201E)Speichern$([char]0x201C)." } else { 'Noch keine Liste: Verwalter ist nur, wer das Herausgeber-Zertifikat hat.' })
    }
    & $fillAdm
    # Rechte der Verwalter-Liste (NTFS): Warnzeile mit Details im Tooltip; nicht lesbar = still "nicht pruefbar"
    # Im Hintergrund mit Zeitlimit: Rechte einer (langsamen) Freigabe zu lesen darf den Fenster-Thread nie blockieren
    $c.txtAdminAcl.Text = 'Rechte der Verwalter-Liste werden geprüft …'
    try {
        Start-HUBackgroundJob -Owner 'Einstellungen' -LogPrefix 'Rechte' -TimeoutSec 20 -Context $c.txtAdminAcl `
            -Arguments @($script:AppRoot, $script:DataRoot) -Script {
                param($AppRoot, $DataRoot)
                Import-Module (Join-Path $AppRoot 'Core\HU.Paths.psm1') -DisableNameChecking -ErrorAction Stop
                Test-HUVerwalterAcl -DataRoot $DataRoot -AppRoot $AppRoot
            } -Done {
                param($Res, $Err, $Ctx)
                try {
                    if ($Err -or -not $Res) { $Ctx.Text = 'Rechte der Verwalter-Liste: nicht prüfbar.'; $Ctx.ToolTip = "$(if ($Err) { $Err } else { 'Keine Antwort.' })"; return }
                    $Ctx.Text = $Res.Text; $Ctx.ToolTip = $Res.ToolTip
                    if ($Res.Status -eq 'Warnung') { $Ctx.Foreground = Get-HUBrush '#FFB74D' }
                } catch { }   # Fenster schon zu
            }
    } catch { $c.txtAdminAcl.Visibility = 'Collapsed' }
    # Datenordner (nur Verwalter sehen den Reiter)
    Update-HUDataLocationInfo $c.txtDataLocation
    $c.btnDataLocation.Add_Click({ Invoke-HUDataLocationChange -Owner $w -Info $c.txtDataLocation -Button $c.btnDataLocation })
    $addName = { param($n) $n = "$n".Trim(); if (-not $n) { return }; if (@($st.Names | Where-Object { $_ -ieq $n }).Count) { return }; [void]$st.Names.Add($n); $st.Changed = $true; & $fillAdm }
    $c.btnAdminAdd.Add_Click({ & $addName $c.txtAdminName.Text; $c.txtAdminName.Text = '' })
    $c.btnAdminMe.Add_Click({ & $addName $script:UserName })
    # Testmodus (nur Verwalter sehen den Reiter): starten oder beenden, beides mit Neustart
    if (Test-HUDemoMode) { $c.btnTestMode.IsEnabled = $false; $c.btnTestMode.ToolTip = 'In der Demo nicht möglich. Zuerst die Demo beenden (Leiste oben).' }
    elseif (Test-HUTestMode) { $c.btnTestMode.Content = 'Testmodus beenden ...'; $c.btnTestMode.ToolTip = 'Neu starten und wieder mit den echten Daten arbeiten' }
    $c.btnTestMode.Add_Click({ $w.Close(); Switch-HUTestMode })
    $c.btnAdminDel.Add_Click({
            $sel = "$($c.lstAdmins.SelectedItem)"
            if (-not $sel) { return }
            $self = ($sel -ieq $script:UserName -or $sel -ieq "$env:USERDOMAIN\$($script:UserName)")
            if (-not (Confirm-HU -Warning "$([char]0x201E)$sel$([char]0x201C) als Verwalter entfernen?$(if ($self) { "`n`nDas bist du selbst: Nach dem Speichern siehst du die Verwaltung nicht mehr (außer mit dem Herausgeber-Zertifikat)." })" 'Verwalter' -Owner $w)) { return }
            $st.Names.Remove($sel); $st.Changed = $true; & $fillAdm
        })

    # ---------------- Speichern ----------------
    $c.btnSave.Add_Click({
            if ($adm) {
                $abErr = Save-HUAbSettingsFromForm $c
                if ($abErr) { $c.tabSettings.SelectedItem = $c.tabAppBackup; Show-HUMessage $abErr -Icon Warning -Owner $w; return }
                if (-not [bool]$c.chkRequireSig.IsChecked -and -not ($ucfg -and $ucfg.AllowUnsigned) -and -not (Confirm-HU -Warning "Auch Updates ohne Signatur annehmen?`n`nDas ist nur zum Testen gedacht." 'Update' -Owner $w)) { $c.tabSettings.SelectedItem = $c.tabUpdate; return }
            }
            # persoenlich
            Set-HUProp $S.ui 'startTab' "$($c.cmbStartTab.SelectedItem.Tag)"
            Set-HUProp $S.ui 'splashAnimation' ([bool]$c.chkSplashAnim.IsChecked)
            Set-HUProp $S.ui 'checkUpdatesOnStart' ([bool]$c.chkCheckOnStart.IsChecked)
            if ($adm -and $c.cmbLogLevel.SelectedItem) { Set-HUProp $S.logging 'logLevel' "$($c.cmbLogLevel.SelectedItem.Tag)".ToUpper() }
            try { Save-HUUserSettings } catch { Show-HUMessage "Einstellungen nicht gespeichert:`n$($_.Exception.Message)" -Icon Error -Owner $w; return }
            if ($c.cmbScale.SelectedItem) { Set-HUUiScale ([double]::Parse("$($c.cmbScale.SelectedItem.Tag)", [Globalization.CultureInfo]::InvariantCulture)) }
            if ($adm) {
                # Installation
                try { if ($c.cmbPsHost.SelectedItem) { Set-HUPsHostSetting "$($c.cmbPsHost.SelectedItem.Tag)" } } catch { Write-HULogWarn "PowerShell beim Start nicht gespeichert: $($_.Exception.Message)" }
                # update.json (andere Schluessel bleiben)
                try {
                    $u = if ($uRaw) { $uRaw } else { [pscustomobject]@{} }
                    Set-HUProp $u 'Channel' "$($c.cmbChannel.SelectedItem.Tag)"
                    if ($c.chkRequireSig.IsChecked) { if ($u.PSObject.Properties['AllowUnsigned']) { $u.PSObject.Properties.Remove('AllowUnsigned') } } else { Set-HUProp $u 'AllowUnsigned' $true }
                    Write-HUJsonFile -Path (Join-Path $script:ConfigDir 'update.json') -Object $u -Depth 5
                } catch { Write-HULogWarn "update.json: $($_.Exception.Message)" }
                # Kanal in der Kopfzeile sofort nachziehen
                try { [void](Update-UpdateConfig); Update-VersionDisplay } catch { }
                if ($st.Changed) {
                    $why = Save-HUVerwalterList @($st.Names)
                    if ($why) { Show-HUMessage "Verwalter nicht gespeichert:`n$why" -Icon Warning -Owner $w; return }
                    $script:IsVerwalter = $null
                    Write-HULogOK "Verwalter gespeichert: $(@($st.Names) -join ', ')"
                }
            }
            $st.Saved = $true
            Write-HULogOK 'Einstellungen gespeichert.'
            Update-HUAdminView
            $w.Close()
        })
    $c.btnCancel.Add_Click({ $w.Close() })
    $w.Add_Closing({ Save-HUDialogState $w 'SettingsWindow'; $script:AbDlg = $null })

    switch ($Tab) { 'Update' { $c.tabSettings.SelectedItem = $c.tabUpdate } 'Sicherung' { if ($adm) { $c.tabSettings.SelectedItem = $c.tabAppBackup } } 'Module' { if ($adm) { $c.tabSettings.SelectedItem = $c.tabModules } } 'Verwalter' { if ($adm) { $c.tabSettings.SelectedItem = $c.tabAdmins } } default { } }
    $c.txtSettingsInfo.Text = "Programmordner: $($script:AppRoot)  |  Datenordner: $($script:DataRoot)$(if ($adm) { "  |  Verwalter: $($script:UserName)" })"
    [void]$w.ShowDialog()
    return $st.Saved
}
