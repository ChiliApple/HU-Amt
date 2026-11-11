#Requires -Version 5.1
<#
.SYNOPSIS
    Oberflaeche fuer Aussenstellen: Fenster "Paket fuer Aussenstelle erstellen" (nur Verwalter der Zentrale)
    und der erste Start eines Pakets (Programm pruefen, beigelegten Zugang einlesen).
.NOTES
    Dot-Source aus Main.ps1 (nach UI-Settings.ps1). Logik in Functions\Core-Outpost.ps1.
    Zielmaschine: der PC, auf dem die App laeuft.
#>

function Show-HUOutpostDialog($Owner = $null) {
    $t = "Paket f$([char]0xFC)r Au$([char]0xDF)enstelle"
    if (-not (Test-HUVerwalter)) { Show-HUMessage 'Pakete für Außenstellen erstellen nur Verwalter.' $t -Icon Info -Owner $Owner; return }
    if (Test-HUTestMode) { Show-HUMessage 'Im Testmodus und in der Demo nicht möglich.' $t -Icon Info -Owner $Owner; return }
    try { Remove-HUOldPackageTemp } catch { }
    # eigene Update-Quelle bzw. eigenes Zertifikat: die Aussenstelle wuerde mit der offiziellen Quelle pruefen, darum ablehnen
    $cfg0 = $null; try { $cfg0 = Update-UpdateConfig } catch { }
    if ($cfg0 -and (("$($cfg0.Owner)" -and "$($cfg0.Owner)" -ne "$($script:App.repo.owner)") -or ("$($cfg0.Repo)" -and "$($cfg0.Repo)" -ne "$($script:App.repo.name)") -or ("$($cfg0.SignerThumbprint)" -and "$($cfg0.SignerThumbprint)" -ne "$($script:App.signer)"))) {
        Show-HUMessage 'Diese Installation nutzt eine eigene Update-Quelle. Pakete für Außenstellen gehen nur mit der offiziellen Quelle (Einstellungen › Update).' $t -Icon Info -Owner $Owner; return
    }
    # ins Paket kommen die signierten Module aus Addons\ - unabhaengig davon, welche Fassung gerade geladen ist (z. B. AddonsDev)
    $mods = @(Get-HUAddonsInDir (Get-HUAddonsDir) ([Version]$script:Version) 'Privat' | Where-Object { -not $_.Error } |
        Sort-Object @{ Expression = { $_.Order } }, @{ Expression = { $_.Title } })
    if (-not $mods.Count) { Show-HUMessage "In dieser Installation liegt kein Modul unter Addons.`n`nEin Paket enthielte nur das Programm. Zuerst die Module über die Zugangsdatei laden." $t -Icon Info -Owner $Owner }
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Paket f&#xFC;r Au&#xDF;enstelle" Width="760" MinWidth="700" Height="600" MinHeight="480" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel Margin="16">
        <TextBlock DockPanel.Dock="Top" Style="{StaticResource HintText}" TextWrapping="Wrap" Margin="0,0,0,12"
                   Text="Erstellt ein ZIP f&#xFC;r eine Au&#xDF;enstelle. Dort wird es ausgepackt und mit Einrichten doppelgeklickt, HU-Amt richtet sich dann f&#xFC;r den Windows-Benutzer selbst ein. Daten, Benutzer und Protokolle der Zentrale kommen nicht mit."/>
        <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <TextBlock Text="Name der Au&#xDF;enstelle:" Foreground="#AAAAAA" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <TextBox x:Name="txtName" Grid.Column="1" Style="{StaticResource DarkTextBox}" Margin="0,0,0,6" MaxLength="60" ToolTip="z. B. Kindergarten. Steht dort im Fenstertitel und im Namen des ZIP."/>
            <TextBlock Grid.Row="1" Text="Verwalter dort:" Foreground="#AAAAAA" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtAdmins" Grid.Row="1" Grid.Column="1" Style="{StaticResource DarkTextBox}" ToolTip="Windows-Benutzernamen der Verwalter in der Au&#xDF;enstelle, mehrere mit Komma getrennt. Nur Verwalter sehen dort Einstellungen, Module und Protokoll. Vorbelegt sind die Verwalter der Zentrale."/>
        </Grid>
        <StackPanel DockPanel.Dock="Bottom">
            <CheckBox x:Name="chkKeys" IsChecked="True" Style="{StaticResource DarkCheckBox}" Margin="0,10,0,4" Content="Zugang f&#xFC;r Updates der mitgelieferten Module beilegen"
                      ToolTip="Legt die hinterlegten Nur-Lese-Tokens der mitgelieferten Module bei. Wer HU-Amt in der Au&#xDF;enstelle einrichtet, bekommt damit die Updates der Module."/>
            <TextBlock x:Name="txtWarn" Foreground="#FFB74D" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,8"
                       Text="Das ZIP enth&#xE4;lt den Zugang und die mitgegebenen Einstellungen (z. B. den Schl&#xFC;ssel f&#xFC;r die Routenberechnung) im Klartext. Nur auf sicherem Weg weitergeben und nach dem Einrichten l&#xF6;schen."/>
            <WrapPanel HorizontalAlignment="Right">
                <Button x:Name="btnCreate" Content="Paket erstellen ..." Background="#388E3C" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Speicherort w&#xE4;hlen und das ZIP erstellen"/>
                <Button x:Name="btnClose" Content="Schlie&#xDF;en" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Fenster schlie&#xDF;en"/>
            </WrapPanel>
        </StackPanel>
        <TextBlock DockPanel.Dock="Top" Text="Module" Foreground="#AAAAAA" FontSize="11" Margin="0,0,0,2"/>
        <DataGrid x:Name="gridMods" Style="{StaticResource DarkDataGrid}" AutoGenerateColumns="False" IsReadOnly="False" HeadersVisibility="Column" SelectionMode="Single"
                  CanUserAddRows="False" CanUserSortColumns="False" CanUserDeleteRows="False" ToolTip="Mitliefern = Modul kommt ins Paket. Aktiv = erscheint dort sofort. Einstellungen = Einstellungen des Moduls mitgeben (nur, wenn es welche gibt).">
            <DataGrid.Columns>
                <DataGridTemplateColumn>
                    <DataGridTemplateColumn.Header>
                        <CheckBox x:Name="chkAllMit" Content="Mitliefern" Style="{StaticResource DarkCheckBox}" ToolTip="Alle Module mitliefern bzw. keines"/>
                    </DataGridTemplateColumn.Header>
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Mit, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" Style="{StaticResource DarkCheckBox}"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn>
                    <DataGridTemplateColumn.Header>
                        <CheckBox x:Name="chkAllAktiv" Content="Aktiv" Style="{StaticResource DarkCheckBox}" ToolTip="Bei allen Zeilen: sofort aktiv bzw. vorerst aus. Gilt nur f&#xFC;r Module mit Haken bei Mitliefern."/>
                    </DataGridTemplateColumn.Header>
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Aktiv, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" Style="{StaticResource DarkCheckBox}"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn>
                    <DataGridTemplateColumn.Header>
                        <CheckBox x:Name="chkAllEinst" Content="Einstellungen" Style="{StaticResource DarkCheckBox}" ToolTip="Einstellungen mitgeben bzw. nicht. Gibt es zu einem Modul keine Einstellungen, kommt nichts mit. Gilt nur f&#xFC;r Module mit Haken bei Mitliefern."/>
                    </DataGridTemplateColumn.Header>
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding Einst, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center" Style="{StaticResource DarkCheckBox}"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Modul" Binding="{Binding Reiter}" IsReadOnly="True"/>
                <DataGridTextColumn Header="Version" Binding="{Binding Version}" IsReadOnly="True"/>
                <DataGridTextColumn Header="Token hinterlegt" Binding="{Binding Token}" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    if ($Owner) { try { $w.Owner = $Owner } catch { } }
    $sources = @(); try { $sources = @(Get-HUAddonSources) } catch { }
    $tbl = New-Object System.Data.DataTable
    foreach ($col in 'Id', 'Reiter', 'Version', 'Token') { [void]$tbl.Columns.Add($col, [string]) }
    foreach ($col in 'Mit', 'Aktiv', 'Einst') { [void]$tbl.Columns.Add($col, [bool]) }
    foreach ($a in $mods) {
        $r = $tbl.NewRow()
        $r['Id'] = "$($a.Id)"; $r['Reiter'] = "$($a.Title)"; $r['Version'] = "$($a.Version)"
        $r['Token'] = $(if (Find-HUAddonSourceFor "$($a.Id)" $sources) { 'ja' } else { 'nein' })
        $hasSet = Test-Path -LiteralPath (Join-Path (Join-Path $script:DataRoot "$($a.Id)") 'einstellungen.json')
        $r['Mit'] = $false; $r['Aktiv'] = $true; $r['Einst'] = [bool]$hasSet
        [void]$tbl.Rows.Add($r)
    }
    $c.gridMods.ItemsSource = $tbl.DefaultView
    # Kopf-Haekchen: ganze Spalte setzen bzw. leeren (Haken in den Zeilen genuegt je ein Klick)
    $allCol = { param($Col, $Box) foreach ($r in @($tbl.Rows)) { $r[$Col] = [bool]$Box.IsChecked } }
    $c.chkAllMit.IsChecked = $false; $c.chkAllAktiv.IsChecked = $true; $c.chkAllEinst.IsChecked = $true
    $c.chkAllMit.Add_Click({ & $allCol 'Mit' $c.chkAllMit })
    $c.chkAllAktiv.Add_Click({ & $allCol 'Aktiv' $c.chkAllAktiv })
    $c.chkAllEinst.Add_Click({ & $allCol 'Einst' $c.chkAllEinst })
    # Vorbelegt (ab 1.0.5): alle Verwalter der Zentrale, damit jeder von ihnen auch dort verwalten kann; sonst nur man selbst
    $za = @(); try { $za = @(Get-HUVerwalterList | Where-Object { $_ }) } catch { }
    if ($za -notcontains "$($script:UserName)") { $za = @("$($script:UserName)") + $za }
    $c.txtAdmins.Text = (@($za | Select-Object -Unique) -join ', ')
    $c.btnClose.Add_Click({ $w.Close() })
    $c.btnCreate.Add_Click({
            $name = "$($c.txtName.Text)".Trim()
            if (-not $name) { Show-HUMessage 'Bitte den Namen der Außenstelle eintragen, z. B. Kindergarten.' $t -Icon Info -Owner $w; return }
            $admins = @("$($c.txtAdmins.Text)" -split '[,;]' | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
            if (-not $admins.Count) { Show-HUMessage 'Bitte mindestens einen Verwalter für die Außenstelle eintragen.' $t -Icon Info -Owner $w; return }
            # offene Bearbeitung der Tabelle uebernehmen, sonst fehlt der zuletzt gesetzte Haken
            try { [void]$c.gridMods.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch { }
            foreach ($rv in @($tbl.DefaultView)) { try { $rv.EndEdit() } catch { } }
            $sel = @($tbl.Rows | Where-Object { [bool]$_['Mit'] } | ForEach-Object { [pscustomobject]@{ Id = "$($_['Id'])"; Active = [bool]$_['Aktiv']; Settings = [bool]$_['Einst'] } })
            if (-not $sel.Count -and -not (Confirm-HU 'Kein Modul ausgewählt. Trotzdem ein Paket nur mit dem Programm erstellen?' $t -Owner $w)) { return }
            try { $folder = ConvertTo-HUPackageName $name } catch { Show-HUMessage $_.Exception.Message $t -Icon Info -Owner $w; return }
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter = 'ZIP-Datei (*.zip)|*.zip'; $dlg.FileName = "HU-Amt-$folder.zip"; $dlg.OverwritePrompt = $false
            if (-not $dlg.ShowDialog($w)) { return }
            $zip = $dlg.FileName
            # vorhandene Datei wird erst ersetzt, wenn das neue Paket fertig ist
            $over = $false
            if (Test-Path -LiteralPath $zip) {
                if (-not (Confirm-HU -Warning "Die Datei gibt es schon:`n$zip`n`nErsetzen?" $t -Owner $w)) { return }
                $over = $true
            }
            $cfg = $null; try { $cfg = Update-UpdateConfig } catch { }
            $inst = Get-HUInstalledInfo
            $ver = $(if ($inst -and $inst.PSObject.Properties['Version']) { "$($inst.Version)" } else { "$($script:Version)" })
            $tp = "$($script:App.signer)"
            $owner = $(if ($cfg -and $cfg.Owner) { "$($cfg.Owner)" } else { "$($script:App.repo.owner)" })
            $repo = $(if ($cfg -and $cfg.Repo) { "$($cfg.Repo)" } else { "$($script:App.repo.name)" })
            $useSrc = $(if ([bool]$c.chkKeys.IsChecked) { @($sources | ForEach-Object { [pscustomobject]@{ Owner = $_.Owner; Repo = $_.Repo; Token = $_.Token } }) } else { @() })
            $jobArgs = @{
                AppRoot = $script:AppRoot; DataRoot = $script:DataRoot; ConfigDir = $script:ConfigDir; ZipPath = $zip; Name = $name; Admins = $admins; Modules = $sel
                Sources = $useSrc; IncludeKeys = [bool]$c.chkKeys.IsChecked; Version = $ver; Channel = $(if ($cfg) { "$($cfg.Channel)" } else { 'Stable' })
                Thumbprint = $tp; Owner = $owner; Repo = $repo; User = "$($script:UserName)"; Overwrite = $over
                Libs = @($script:UpdateLib, (Join-Path $script:AppRoot 'Functions\Core-Addons.ps1'), (Join-Path $script:AppRoot 'Functions\Core-Outpost.ps1')); PathsMod = (Join-Path $script:AppRoot 'Core\HU.Paths.psm1')
            }
            $c.btnCreate.IsEnabled = $false; $c.btnCreate.Content = 'Paket wird erstellt ...'
            Write-HULogInfo "Paket für Außenstelle $name wird erstellt: $zip"
            Start-HUBackgroundJob -Owner 'Paket' -LogPrefix 'Paket' -TimeoutSec 600 -Context @{ Window = $w; Button = $c.btnCreate; Title = $t } -Arguments @($jobArgs) -Script {
                param($A)
                Import-Module $A.PathsMod -DisableNameChecking -ErrorAction Stop
                foreach ($l in $A.Libs) { . $l }
                try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
                $m = Get-HUVerifiedInstallManifest -ConfigDir $A.ConfigDir -Version $A.Version -Owner $A.Owner -Repo $A.Repo -Thumbprint $A.Thumbprint
                New-HUOutpostPackage -AppRoot $A.AppRoot -DataRoot $A.DataRoot -ZipPath $A.ZipPath -Name $A.Name -Admins $A.Admins -Modules $A.Modules `
                    -Sources $A.Sources -IncludeKeys:$A.IncludeKeys -Manifest $m.Manifest -Signature $m.Signature -Version $A.Version -Channel $A.Channel `
                    -Thumbprint $A.Thumbprint -CreatedBy $A.User -Overwrite:$A.Overwrite
            } -Done {
                param($Res, $Err, $Ctx)
                try { $Ctx.Button.IsEnabled = $true; $Ctx.Button.Content = 'Paket erstellen ...' } catch { }
                $r = @($Res)[-1]
                if ($Err -or -not $r -or -not $r.Path) {
                    $why = $(if ($Err) { "$Err" } else { 'keine Rückmeldung' })
                    Write-HULogError "Paket nicht erstellt: $why"
                    Show-HUMessage "Das Paket wurde nicht erstellt:`n`n$why" $Ctx.Title -Icon Error -Owner $Ctx.Window
                    return
                }
                Write-HULogOK "Paket erstellt: $($r.Path) ($($r.Files) Programmdateien, Module: $(if (@($r.Modules).Count) { @($r.Modules) -join ', ' } else { 'keine' }))"
                foreach ($x in @($r.Warnings)) { Write-HULogWarn "Paket: $x" }
                $txt = "Paket erstellt:`n$($r.Path)`n`nModule: $(if (@($r.Modules).Count) { @($r.Modules) -join ', ' } else { 'keine' })" +
                $(if (@($r.Settings).Count) { "`nEinstellungen mitgegeben: $(@($r.Settings) -join ', ')" }) +
                $(if (@($r.Tokens).Count) { "`nZugang für Updates: $(@($r.Tokens).Count) Modul(e)" }) +
                $(if (@($r.Warnings).Count) { "`n`nHinweise:`n  $(@($r.Warnings) -join "`n  ")" }) +
                "`n`nIn der Außenstelle: ZIP mit Rechtsklick und 'Alle extrahieren ...' auspacken, dann Einrichten doppelklicken. HU-Amt richtet sich für den angemeldeten Windows-Benutzer ein."
                Show-HUMessage $txt $Ctx.Title -Icon $(if (@($r.Warnings).Count) { 'Warning' } else { 'Info' }) -Owner $Ctx.Window
                try { Open-HUPath (Split-Path $r.Path -Parent) } catch { }
            }
        })
    [void]$w.ShowDialog()
}

# Start einer aus einem Paket eingerichteten Installation: Programm pruefen, bis es einmal geklappt hat (sonst bei jedem
# Start eine Warnung), danach den beigelegten Zugang (Zugang.humodule) fuer diesen Benutzer einlesen.
# Der Zugang wird ohne Verwalter-Pruefung eingelesen (die Installation liegt im eigenen Benutzerordner), aber nur nach
# bestandener Pruefung und nur, wenn seine Pruefsumme zu paket.json passt.
function Invoke-HUPackageFirstStart {
    $chk = Test-HUPackageFirstStart -AppRoot $script:AppRoot -Thumbprint "$($script:App.signer)"
    if ($chk.Status -eq 'Fehler') {
        Write-HULogError "Paket: $($chk.Text)"
        Show-HUMessage ("HU-Amt ist beschädigt oder verändert.`n`nBitte HU-Amt nicht verwenden und in der Gemeinde um ein neues Paket bitten.`n`n" +
            "Technischer Grund: $($chk.Text)") 'Paket' -Icon Warning -Owner $script:Window
        return
    }
    if ($chk.Status -eq 'OK') {
        Write-HULogOK $chk.Text
        # neues Paket: Zugaenge zu Modulen, die nicht mehr dabei sind, entfernen
        try {
            foreach ($s in @(Get-HUPackageStaleSources -AppRoot $script:AppRoot -Sources @(Get-HUAddonSources))) {
                Remove-HUAddonToken $s.Owner $s.Repo
                Write-HULogInfo "Paket: Zugang $($s.Owner)/$($s.Repo) entfernt, das Modul ist nicht mehr im Paket"
            }
        } catch { Write-HULogWarn "Paket: Zugänge nicht bereinigt: $($_.Exception.Message)" }
    }
    $kf = Get-HUPackageKeyFile $script:AppRoot
    if ($kf) { Import-HUAddonKeyFile $kf -Owner $script:Window -FromPackage }
}
