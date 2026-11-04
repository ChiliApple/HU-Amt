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
function Get-HUProp($Object, [string]$Name, $Default = $null) {
    if ($Object -and $Object.PSObject.Properties[$Name] -and $null -ne $Object.$Name) { return $Object.$Name }
    return $Default
}

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

function Open-HUSettings([string]$Tab = '') {
    [void](Show-HUSettingsDialog -Tab $Tab)
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
    $addName = { param($n) $n = "$n".Trim(); if (-not $n) { return }; if (@($st.Names | Where-Object { $_ -ieq $n }).Count) { return }; [void]$st.Names.Add($n); $st.Changed = $true; & $fillAdm }
    $c.btnAdminAdd.Add_Click({ & $addName $c.txtAdminName.Text; $c.txtAdminName.Text = '' })
    $c.btnAdminMe.Add_Click({ & $addName $script:UserName })
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
    $c.txtSettingsInfo.Text = "Programmordner: $($script:AppRoot)$(if ($adm) { "  |  Verwalter: $($script:UserName)" })"
    [void]$w.ShowDialog()
    return $st.Saved
}
