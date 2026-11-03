#Requires -Version 5.1
<#
.SYNOPSIS
    Einstellungen: Allgemein (alle), Sicherung, Update, Module, Verwalter (nur Verwalter). Verwalter-Ansicht im Hauptfenster.
.DESCRIPTION
    Persoenlich:  Data\Benutzer\<Name>\settings.json (ui.splashAnimation, ui.checkUpdatesOnStart, ui.maxLogLines)
    Installation: Config\update.json (Kanal, Signaturpflicht), Config\powershell-host.txt (auto | 5 | 7)
    Verwalter:    Data\Verwalter.txt (ein Windows-Benutzername je Zeile; mit Listen-Sperre gespeichert)
.NOTES
    Dot-Source aus Main.ps1. Zielmaschine: der PC, auf dem die App laeuft.
#>

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
    try { Update-HUAbStatus } catch { }
}

function Open-HUSettings([string]$Tab = '') {
    $adm = Test-HUVerwalter
    $d = New-HUWindow 'SettingsWindow'
    $w = $d.Window; $c = $d.C
    Restore-HUDialogState $w 'dlgSettings'
    # nur Verwalter: Reiter und PowerShell-Auswahl
    foreach ($n in 'tabBackup', 'tabUpdate', 'tabModules', 'tabAdmins') { $c[$n].Visibility = $(if ($adm) { 'Visible' } else { 'Collapsed' }) }
    $c.pnlPsHost.Visibility = $(if ($adm) { 'Visible' } else { 'Collapsed' })
    $c.txtSettingsInfo.Text = $(if ($adm) { "Angemeldet als Verwalter ($($script:UserName))." } else { "Persönliche Einstellungen von $($script:UserName)." })

    # --- Allgemein
    $c.chkSplash.IsChecked = ($script:Settings.ui.splashAnimation -ne $false)
    $c.chkCheckOnStart.IsChecked = ($script:Settings.ui.checkUpdatesOnStart -ne $false)
    $c.txtMaxLog.Text = "$($script:Settings.ui.maxLogLines)"
    $c.txtShortcutInfo.Text = $(if (Test-HUShortcut) { 'Eine Verknüpfung ist schon vorhanden.' } else { 'Noch keine Verknüpfung.' })
    $c.btnShortcutDesktop.Add_Click({ try { $l = New-HUShortcut -Location Desktop; Write-HULogOK "Verknüpfung erstellt: $l"; $c.txtShortcutInfo.Text = "Erstellt: $l" } catch { Show-HUMessage "Verknüpfung fehlgeschlagen:`n$($_.Exception.Message)" -Icon Warning -Owner $w } })
    $c.btnShortcutStart.Add_Click({ try { $l = New-HUShortcut -Location StartMenu; Write-HULogOK "Startmenü-Eintrag erstellt: $l"; $c.txtShortcutInfo.Text = "Erstellt: $l" } catch { Show-HUMessage "Startmenü-Eintrag fehlgeschlagen:`n$($_.Exception.Message)" -Icon Warning -Owner $w } })
    $ph = Get-HUPsHostSetting
    foreach ($it in $c.cmbPsHost.Items) { if ("$($it.Tag)" -eq $ph) { $c.cmbPsHost.SelectedItem = $it } }
    $c.txtPsHostInfo.Text = Get-HUPsHostInfoText

    # --- Sicherung
    $abSum = {
        $s = Get-HUAbSettings
        $last = ''; try { if ("$($s.Last)") { $last = ([datetime]::Parse("$($s.Last)", [Globalization.CultureInfo]::InvariantCulture)).ToString('dd.MM.yyyy HH:mm') } } catch { }
        $mode = switch ("$($s.Mode)") { 'taeglich' { 'täglich' } 'woechentlich' { 'wöchentlich' } default { 'aus' } }
        $c.txtAbSummary.Text = "Ordner: $(if ($s.Folder) { $s.Folder } else { '(noch keiner)' })`nAutomatisch: $mode, behalten: $($s.Keep) je PC`nZuletzt: $(if ($last) { "$last$(if ($s.LastBy) { " von $($s.LastBy)" })" } else { 'noch keine Sicherung' })`nKennwort für die Automatik: $(if (Get-HUAppBackupPassword -DataRoot $script:UserRoot) { 'gespeichert' } else { 'nicht gespeichert' })"
    }
    if ($adm) { & $abSum }
    $c.btnAbOpenDlg.Add_Click({ Show-HUAppBackupDialog; & $abSum })

    # --- Update
    $cfg = Update-UpdateConfig
    foreach ($it in $c.cmbChannel.Items) { if ($cfg -and "$($it.Tag)" -eq "$($cfg.Channel)") { $c.cmbChannel.SelectedItem = $it } }
    if (-not $c.cmbChannel.SelectedItem) { $c.cmbChannel.SelectedIndex = 0 }
    $c.chkRequireSig.IsChecked = -not ($cfg -and $cfg.AllowUnsigned)
    $c.txtUpdateInfo.Text = "Quelle: $(if ($cfg) { "$($cfg.Owner)/$($cfg.Repo)" })  |  installiert: v$($script:Version)$(if ($cfg -and $cfg.SignerThumbprint) { "  |  Zertifikat: $($cfg.SignerThumbprint.Substring(0, 8))..." })"
    $c.btnCheckNow.Add_Click({ Invoke-UpdateCheck })
    $c.btnOtherVersion.Add_Click({ Show-HUVersionPicker })

    # --- Module
    $c.gridModules.ItemsSource = @($script:LoadedAddons | ForEach-Object {
            [pscustomobject]@{ Reiter = "$($_.Title)"; Id = $_.Id; Version = "$($_.Version)"; Herkunft = "$($_.Source)"; Ziehen = $(if (@($_.Drop).Count) { @($_.Drop) -join ' ' } else { '-' }) } })
    $c.btnAddonRepos.Add_Click({ Show-HUAddonDialog })
    $c.btnAddonUpdate.Add_Click({ Start-HUAddonUpdate -Loud })
    $c.btnAddonDev.Add_Click({ $p = Join-Path $script:AppRoot 'AddonsDev'; if (-not (Test-Path -LiteralPath $p)) { [void](New-Item -ItemType Directory -Path $p -Force) }; Open-HUPath $p })

    # --- Verwalter
    $list = Get-HUVerwalterList
    $st = @{ Names = (New-Object System.Collections.ArrayList); Changed = $false }
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

    if ($Tab -and $c["tab$Tab"] -and $c["tab$Tab"].Visibility -eq 'Visible') { $c.tabSettings.SelectedItem = $c["tab$Tab"] }

    $c.btnSave.Add_Click({
            # persoenlich
            $n = 1000; if (-not [int]::TryParse("$($c.txtMaxLog.Text)", [ref]$n) -or $n -lt 100 -or $n -gt 10000) { Show-HUMessage 'Zeilen im Protokoll: bitte eine Zahl von 100 bis 10000.' -Icon Info -Owner $w; return }
            $script:Settings.ui.splashAnimation = [bool]$c.chkSplash.IsChecked
            $script:Settings.ui.checkUpdatesOnStart = [bool]$c.chkCheckOnStart.IsChecked
            $script:Settings.ui.maxLogLines = $n
            try { Save-HUUserSettings } catch { Show-HUMessage "Einstellungen nicht gespeichert:`n$($_.Exception.Message)" -Icon Error -Owner $w; return }
            if ($adm) {
                # Installation
                try { Set-HUPsHostSetting "$($c.cmbPsHost.SelectedItem.Tag)" } catch { Write-HULogWarn "PowerShell-Auswahl: $($_.Exception.Message)" }
                $uf = Join-Path $script:ConfigDir 'update.json'
                $u = Read-HUJsonFile $uf; if (-not $u) { $u = [pscustomobject]@{} }
                foreach ($kv in @(@('Channel', "$($c.cmbChannel.SelectedItem.Tag)"), @('AllowUnsigned', (-not [bool]$c.chkRequireSig.IsChecked)))) {
                    if ($u.PSObject.Properties[$kv[0]]) { $u.($kv[0]) = $kv[1] } else { $u | Add-Member -NotePropertyName $kv[0] -NotePropertyValue $kv[1] }
                }
                if (-not [bool]$c.chkRequireSig.IsChecked -and -not ($cfg -and $cfg.AllowUnsigned) -and -not (Confirm-HU -Warning "Auch Updates ohne Signatur annehmen?`n`nDas ist nur zum Testen gedacht." 'Update' -Owner $w)) { return }
                try { Write-HUJsonFile -Path $uf -Object $u -Depth 5 } catch { Show-HUMessage "Update-Einstellungen nicht gespeichert:`n$($_.Exception.Message)" -Icon Error -Owner $w; return }
                [void](Update-UpdateConfig); Update-VersionDisplay
                if ($st.Changed) {
                    $why = Save-HUVerwalterList @($st.Names)
                    if ($why) { Show-HUMessage "Verwalter nicht gespeichert:`n$why" -Icon Warning -Owner $w; return }
                    $script:IsVerwalter = $null
                    Write-HULogOK "Verwalter gespeichert: $(@($st.Names) -join ', ')"
                }
            }
            Write-HULogOK 'Einstellungen gespeichert.'
            Update-HUAdminView
            $w.Close()
        })
    $w.Add_Closing({ Save-HUDialogState $w 'dlgSettings' })
    [void]$w.ShowDialog()
}
