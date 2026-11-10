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
                   Text="Erstellt ein ZIP f&#xFC;r eine Au&#xDF;enstelle, z. B. auf einen USB-Stick. Dort wird es entpackt und mit Start-HU-Amt.cmd gestartet. Daten, Benutzer und Protokolle der Zentrale kommen nicht mit."/>
        <Grid DockPanel.Dock="Top" Margin="0,0,0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <TextBlock Text="Name der Au&#xDF;enstelle:" Foreground="#AAAAAA" VerticalAlignment="Center" Margin="0,0,10,6"/>
            <TextBox x:Name="txtName" Grid.Column="1" Style="{StaticResource DarkTextBox}" Margin="0,0,0,6" MaxLength="60" ToolTip="z. B. Kindergarten. Steht dort im Fenstertitel und im Namen des ZIP."/>
            <TextBlock Grid.Row="1" Text="Verwalter dort:" Foreground="#AAAAAA" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtAdmins" Grid.Row="1" Grid.Column="1" Style="{StaticResource DarkTextBox}" ToolTip="Windows-Benutzernamen, mehrere mit Komma getrennt. Verwalter sehen dort Einstellungen, Module und Protokoll."/>
        </Grid>
        <StackPanel DockPanel.Dock="Bottom">
            <CheckBox x:Name="chkKeys" IsChecked="True" Foreground="#CCCCCC" Margin="0,10,0,4" Content="Zugang f&#xFC;r Updates der mitgelieferten Module beilegen"
                      ToolTip="Legt die hinterlegten Nur-Lese-Tokens der mitgelieferten Module bei. Beim ersten Start in der Au&#xDF;enstelle werden sie eingelesen und die Datei gel&#xF6;scht."/>
            <TextBlock x:Name="txtWarn" Foreground="#FFB74D" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,8"
                       Text="Das ZIP enth&#xE4;lt den Zugang und die mitgegebenen Einstellungen (z. B. den Schl&#xFC;ssel f&#xFC;r die Routenberechnung) im Klartext. Den USB-Stick sicher verwahren und das ZIP nach dem Einrichten l&#xF6;schen."/>
            <WrapPanel HorizontalAlignment="Right">
                <Button x:Name="btnCreate" Content="Paket erstellen ..." Background="#388E3C" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Speicherort w&#xE4;hlen und das ZIP erstellen"/>
                <Button x:Name="btnClose" Content="Schlie&#xDF;en" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Fenster schlie&#xDF;en"/>
            </WrapPanel>
        </StackPanel>
        <TextBlock DockPanel.Dock="Top" Text="Module" Foreground="#AAAAAA" FontSize="11" Margin="0,0,0,2"/>
        <DataGrid x:Name="gridMods" Style="{StaticResource DarkDataGrid}" AutoGenerateColumns="False" IsReadOnly="False" HeadersVisibility="Column"
                  CanUserAddRows="False" CanUserDeleteRows="False" ToolTip="Mitliefern = Modul kommt ins Paket. Aktiv = erscheint dort sofort. Einstellungen = einstellungen.json des Moduls mitgeben.">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="Mitliefern" Binding="{Binding Mit, UpdateSourceTrigger=PropertyChanged}"/>
                <DataGridCheckBoxColumn Header="Aktiv" Binding="{Binding Aktiv, UpdateSourceTrigger=PropertyChanged}"/>
                <DataGridCheckBoxColumn Header="Einstellungen" Binding="{Binding Einst, UpdateSourceTrigger=PropertyChanged}"/>
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
        $r['Mit'] = $false; $r['Aktiv'] = $true; $r['Einst'] = $true
        [void]$tbl.Rows.Add($r)
    }
    $c.gridMods.ItemsSource = $tbl.DefaultView
    $c.txtAdmins.Text = "$($script:UserName)"
    $c.btnClose.Add_Click({ $w.Close() })
    $c.btnCreate.Add_Click({
            $name = "$($c.txtName.Text)".Trim()
            if (-not $name) { Show-HUMessage 'Bitte den Namen der Außenstelle eintragen, z. B. Kindergarten.' $t -Icon Info -Owner $w; return }
            $admins = @("$($c.txtAdmins.Text)" -split '[,;]' | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
            if (-not $admins.Count) { Show-HUMessage 'Bitte mindestens einen Verwalter für die Außenstelle eintragen.' $t -Icon Info -Owner $w; return }
            $sel = @($tbl.Rows | Where-Object { [bool]$_['Mit'] } | ForEach-Object { [pscustomobject]@{ Id = "$($_['Id'])"; Active = [bool]$_['Aktiv']; Settings = [bool]$_['Einst'] } })
            if (-not $sel.Count -and -not (Confirm-HU 'Kein Modul ausgewählt. Trotzdem ein Paket nur mit dem Programm erstellen?' $t -Owner $w)) { return }
            try { $folder = ConvertTo-HUPackageName $name } catch { Show-HUMessage $_.Exception.Message $t -Icon Info -Owner $w; return }
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter = 'ZIP-Datei (*.zip)|*.zip'; $dlg.FileName = "HU-Amt-$folder.zip"; $dlg.OverwritePrompt = $false
            if (-not $dlg.ShowDialog($w)) { return }
            $zip = $dlg.FileName
            if (Test-Path -LiteralPath $zip) {
                if (-not (Confirm-HU -Warning "Die Datei gibt es schon:`n$zip`n`nErsetzen?" $t -Owner $w)) { return }
                try { Remove-Item -LiteralPath $zip -Force -ErrorAction Stop } catch { Show-HUMessage "Die Datei lässt sich nicht ersetzen:`n$($_.Exception.Message)" $t -Icon Warning -Owner $w; return }
            }
            $cfg = $null; try { $cfg = Update-UpdateConfig } catch { }
            $inst = Get-HUInstalledInfo
            $ver = $(if ($inst -and $inst.PSObject.Properties['Version']) { "$($inst.Version)" } else { "$($script:Version)" })
            $tp = $(if ($cfg -and $cfg.SignerThumbprint) { "$($cfg.SignerThumbprint)" } else { "$($script:App.signer)" })
            $owner = $(if ($cfg -and $cfg.Owner) { "$($cfg.Owner)" } else { "$($script:App.repo.owner)" })
            $repo = $(if ($cfg -and $cfg.Repo) { "$($cfg.Repo)" } else { "$($script:App.repo.name)" })
            $useSrc = $(if ([bool]$c.chkKeys.IsChecked) { @($sources | ForEach-Object { [pscustomobject]@{ Owner = $_.Owner; Repo = $_.Repo; Token = $_.Token } }) } else { @() })
            $jobArgs = @{
                AppRoot = $script:AppRoot; DataRoot = $script:DataRoot; ConfigDir = $script:ConfigDir; ZipPath = $zip; Name = $name; Admins = $admins; Modules = $sel
                Sources = $useSrc; IncludeKeys = [bool]$c.chkKeys.IsChecked; Version = $ver; Channel = $(if ($cfg) { "$($cfg.Channel)" } else { 'Stable' })
                Thumbprint = $tp; Owner = $owner; Repo = $repo; User = "$($script:UserName)"
                Libs = @($script:UpdateLib, (Join-Path $script:AppRoot 'Functions\Core-Addons.ps1'), (Join-Path $script:AppRoot 'Functions\Core-Outpost.ps1')); PathsMod = (Join-Path $script:AppRoot 'Core\HU.Paths.psm1')
            }
            $c.btnCreate.IsEnabled = $false; $c.btnCreate.Content = 'Paket wird erstellt …'
            Write-HULogInfo "Paket für Außenstelle $name wird erstellt: $zip"
            Start-HUBackgroundJob -Owner 'Paket' -LogPrefix 'Paket' -TimeoutSec 600 -Context @{ Window = $w; Button = $c.btnCreate; Title = $t } -Arguments @($jobArgs) -Script {
                param($A)
                Import-Module $A.PathsMod -DisableNameChecking -ErrorAction Stop
                foreach ($l in $A.Libs) { . $l }
                try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
                $m = Get-HUVerifiedInstallManifest -ConfigDir $A.ConfigDir -Version $A.Version -Owner $A.Owner -Repo $A.Repo -Thumbprint $A.Thumbprint
                New-HUOutpostPackage -AppRoot $A.AppRoot -DataRoot $A.DataRoot -ZipPath $A.ZipPath -Name $A.Name -Admins $A.Admins -Modules $A.Modules `
                    -Sources $A.Sources -IncludeKeys:$A.IncludeKeys -Manifest $m.Manifest -Signature $m.Signature -Version $A.Version -Channel $A.Channel `
                    -Thumbprint $A.Thumbprint -CreatedBy $A.User
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
                "`n`nIn der Außenstelle: ZIP an einen festen Platz entpacken (z. B. C:\), dann im Ordner $($r.Folder) Start-HU-Amt.cmd doppelklicken."
                Show-HUMessage $txt $Ctx.Title -Icon $(if (@($r.Warnings).Count) { 'Warning' } else { 'Info' }) -Owner $Ctx.Window
                try { Open-HUPath (Split-Path $r.Path -Parent) } catch { }
            }
        })
    [void]$w.ShowDialog()
}

# Erster Start aus einem Paket: Programm einmal pruefen, beigelegten Zugang (Zugang.humodule) fuer diesen Benutzer einlesen.
# Der Zugang wird ohne Verwalter-Pruefung eingelesen, aber nur, wenn seine Pruefsumme zu paket.json passt.
function Invoke-HUPackageFirstStart {
    $cfg = $null; try { $cfg = Update-UpdateConfig } catch { }
    $tp = $(if ($cfg -and $cfg.SignerThumbprint) { "$($cfg.SignerThumbprint)" } else { "$($script:App.signer)" })
    $chk = Test-HUPackageFirstStart -AppRoot $script:AppRoot -Thumbprint $tp
    if ($chk.Status -eq 'OK') { Write-HULogOK $chk.Text }
    elseif ($chk.Status -eq 'Fehler') {
        Write-HULogError "Paket: $($chk.Text)"
        Show-HUMessage "Das Programm passt nicht zum Paket:`n$($chk.Text)`n`nBitte das Paket neu entpacken oder ein neues Paket erstellen lassen." 'Paket' -Icon Warning -Owner $script:Window
    }
    $kf = Get-HUPackageKeyFile $script:AppRoot
    if ($kf) { Import-HUAddonKeyFile $kf -Owner $script:Window -FromPackage }
}
