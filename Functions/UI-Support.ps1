#Requires -Version 5.1
<#
.SYNOPSIS
    Support: ZIP mit Versionen, Umgebung und (auf Wunsch) den eigenen Protokollen, damit der Verwalter helfen kann.
.DESCRIPTION
    Nie im ZIP: Daten der Module (Data\<Id>\), Token, Kennwoerter, Sicherungen, Fenster-Einstellungen anderer Benutzer.
    Benutzer- und Rechnernamen sowie das Profil in Pfaden werden ersetzt (<Benutzer>, <PC>, <Profil>, <Programmordner>).
    Ziel: Downloads\<App>-Support_<JJJJ-MM-TT_HHMMSS>.zip, danach oeffnet sich der Ordner.
.NOTES
    Dot-Source aus Main.ps1. Zielmaschine: der PC, auf dem die App laeuft.
#>

# Pfade und Namen unkenntlich machen (Text aus Protokollen und Umgebung)
function Protect-HUSupportText([string]$Text) {
    $t = "$Text"
    try { if ($script:AppRoot) { $t = $t.Replace("$($script:AppRoot)", '<Programmordner>') } } catch { }
    $t = $t -replace '(?i)[A-Z]:\\Users\\[^\\\s"]+', '<Profil>'
    foreach ($p in @(@("$env:USERNAME", '<Benutzer>'), @("$env:COMPUTERNAME", '<PC>'), @("$env:USERDOMAIN", '<Domaene>'))) {
        if ($p[0].Length -ge 2) { $t = [regex]::Replace($t, [regex]::Escape($p[0]), $p[1], 'IgnoreCase') }
    }
    return $t
}

# Text mit Versionen und Umgebung
function Get-HUSupportInfo {
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add("$($script:AppTitle) v$($script:Version)  ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))")
    $l.Add("PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)), .NET $([Environment]::Version)")
    try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop; $l.Add("Windows: $($os.Caption) $($os.Version)") } catch { $l.Add("Windows: $([Environment]::OSVersion.VersionString)") }
    $l.Add("Programmordner: $(if ("$($script:AppRoot)" -match '^(\\\\|//)') { 'Netzordner (UNC)' } else { 'lokal' })")
    $cfg = $null; try { $cfg = Update-UpdateConfig } catch { }
    if ($cfg) { $l.Add("Update: $($cfg.Owner)/$($cfg.Repo), Kanal $(Format-HUChannel $cfg.Channel)$(if ($cfg.RequireSignature) { ', nur signiert' })$(if ($cfg.UseBranch) { ", Branch $($cfg.Branch)" })") }
    $inst = $null; try { $inst = Get-HUInstalledInfo } catch { }
    if ($inst) { $l.Add("Installiert: $($inst.Version), $($inst.Date), geprüft: $($inst.Check)") }
    $wv = ''
    foreach ($k in 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'HKCU:\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}') {
        try { $v = (Get-ItemProperty -LiteralPath $k -ErrorAction Stop).pv; if ($v) { $wv = "$v"; break } } catch { }
    }
    $l.Add("WebView2-Runtime: $(if ($wv) { $wv } else { 'nicht gefunden' })")
    $l.Add('')
    $l.Add('Module:')
    if (@($script:LoadedAddons).Count) { foreach ($a in @($script:LoadedAddons)) { $l.Add("  $($a.Id) $($a.Version) ($($a.Source))") } } else { $l.Add('  (keine)') }
    try { $o = @(Get-HURunningInstances -RunDir $script:Paths.RunDir -Self $script:RunLock); $l.Add(''); $l.Add("Weitere offene Sitzungen: $($o.Count)") } catch { }
    return (Protect-HUSupportText ($l -join "`r`n"))
}

# ZIP schreiben. -Logs: eigene Protokolle der letzten 7 Tage (unkenntlich gemacht). Rueckgabe: Pfad
function New-HUSupportZip([switch]$Logs, [string]$Description = '', [string]$Folder = '') {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    if (-not $Folder) { $Folder = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads' }
    if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $zipPath = Join-Path $Folder ("$($script:App.id)-Support_$((Get-Date).ToString('yyyy-MM-dd_HHmmss')).zip")
    $fs = [IO.File]::Open($zipPath, [IO.FileMode]::CreateNew)
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
        try {
            $put = {
                param($name, [string]$text)
                $e = $zip.CreateEntry($name); $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding $true))
                try { $w.Write($text) } finally { $w.Dispose() }
            }
            & $put 'info.txt' (Get-HUSupportInfo)
            if ("$Description".Trim()) { & $put 'beschreibung.txt' (Protect-HUSupportText $Description) }
            if ($Logs -and $script:LogDir -and (Test-Path -LiteralPath $script:LogDir)) {
                $lim = (Get-Date).AddDays(-7)
                foreach ($f in @(Get-ChildItem -LiteralPath $script:LogDir -Filter '*.log' -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $lim })) {
                    $txt = ''
                    try {
                        $s = New-Object IO.FileStream($f.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
                        try { $txt = (New-Object IO.StreamReader($s, [Text.Encoding]::UTF8)).ReadToEnd() } finally { $s.Dispose() }
                    } catch { $txt = "(nicht lesbar: $($_.Exception.Message))" }
                    & $put "logs/$($f.Name)" (Protect-HUSupportText $txt)
                }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
    return $zipPath
}

# Fenster: Support-ZIP erstellen
function Show-HUSupport {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Support" Width="620" Height="520" MinWidth="520" MinHeight="420" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <DockPanel Margin="16">
        <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" Foreground="#CCCCCC" FontSize="14" Margin="0,0,0,12"
                   Text="Packt Versionsangaben und auf Wunsch deine Protokolle in eine ZIP-Datei. Die schickst du an den Verwalter. Daten, Kennwörter und Token sind nie dabei."/>
        <TextBlock DockPanel.Dock="Top" Text="Was ist passiert? (freiwillig)" Style="{StaticResource FieldLabel}" Margin="0,0,0,4"/>
        <TextBox x:Name="txtDesc" DockPanel.Dock="Top" Style="{StaticResource DarkTextBox}" Height="90" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"
                 ToolTip="Kurz beschreiben, was du gemacht hast und was passiert ist"/>
        <CheckBox x:Name="chkLogs" DockPanel.Dock="Top" Style="{StaticResource DarkCheckBox}" IsChecked="True" Margin="0,10,0,0"
                  Content="Meine Protokolle der letzten 7 Tage mitschicken (Namen und Pfade werden unkenntlich gemacht)"
                  ToolTip="Protokolle helfen bei der Fehlersuche; sie können Dateinamen enthalten"/>
        <Border DockPanel.Dock="Bottom" Background="#252526" BorderBrush="#3E3E42" BorderThickness="0,1,0,0" Padding="0,10,0,0" Margin="0,10,0,0">
            <DockPanel>
                <Button x:Name="btnClose" DockPanel.Dock="Right" Content="Schließen" MinWidth="96" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Fenster schließen"/>
                <Button x:Name="btnZip" Content="&#x1F4E6; ZIP erstellen" Background="#1976D2" Style="{StaticResource DarkButton}" FontSize="14" Padding="18,8" HorizontalAlignment="Left"
                        ToolTip="ZIP in den Ordner Downloads schreiben und den Ordner öffnen"/>
            </DockPanel>
        </Border>
        <DockPanel Margin="0,10,0,0">
            <TextBlock DockPanel.Dock="Top" Text="Das kommt in die Datei info.txt:" Style="{StaticResource FieldLabel}" Margin="0,0,0,4"/>
            <TextBox x:Name="txtInfo" Style="{StaticResource DarkTextBox}" IsReadOnly="True" FontFamily="Consolas" FontSize="11" TextWrapping="NoWrap"
                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" ToolTip="Inhalt von info.txt (Namen und Pfade unkenntlich)"/>
        </DockPanel>
    </DockPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    Restore-HUDialogState $w 'dlgSupport'
    $c.txtInfo.Text = Get-HUSupportInfo
    $c.btnZip.Add_Click({
            try {
                $z = New-HUSupportZip -Logs:([bool]$c.chkLogs.IsChecked) -Description "$($c.txtDesc.Text)"
                Write-HULogOK "Support-ZIP erstellt: $z"
                Show-HUMessage "Die Datei liegt im Ordner Downloads:`n$(Split-Path -Leaf $z)`n`nBitte an den Verwalter schicken." -Owner $w
                Start-Process explorer.exe -ArgumentList "/select,`"$z`""
            } catch { Show-HUMessage "ZIP nicht erstellt:`n$($_.Exception.Message)" -Icon Error -Owner $w }
        })
    $w.Add_Closing({ Save-HUDialogState $w 'dlgSupport' })
    [void]$w.ShowDialog()
}
