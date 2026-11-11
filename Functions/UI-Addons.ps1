#Requires -Version 5.1
<#
.SYNOPSIS
    Module in der Oberflaeche: Reiter einbauen, Dateien hineinziehen, beim Beenden stoppen, Neustart.
    Dazu gemeinsame Bausteine fuer Module (pulsierender gruener Punkt beim Laden).
.NOTES
    Dot-Source aus Main.ps1 (nach Core-Addons.ps1). Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:LoadedAddons = @()
$script:AllAddons = @()         # alle gueltigen Module dieser Installation, auch abgeschaltete (Einstellungen > Module)
$script:AddonByTab = @{}
$script:AddonHeader = @{}       # Kennung -> urspruengliche Reiter-Beschriftung (fuer die Zahl im Postfach)
$script:AddonInboxTimer = $null

# Reiter eines (bereits gepruefen) Moduls in das Hauptfenster einbauen und die benannten Elemente registrieren
function Add-HUAddonTab($Info) {
    $x = [IO.File]::ReadAllText((Join-Path $Info.Dir $Info.Xaml), [Text.Encoding]::UTF8)
    $theme = Get-HUXaml 'Theme'
    $mt = [regex]::Match($theme, '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>')
    $x = $x.Replace('<!--HU:THEME-->', $mt.Groups[1].Value)
    try { $x = Convert-HUXamlTheme $x } catch { try { Write-HULogWarn "Farbschema ($($Info.Id)): $($_.Exception.Message)" } catch { } }
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($x))
    try { $tab = [System.Windows.Markup.XamlReader]::Load($reader) } finally { $reader.Close() }
    if (-not ($tab -is [System.Windows.Controls.TabItem])) { throw 'Tab.xaml: Wurzel muss <TabItem> sein' }
    $names = @{}
    foreach ($m in [regex]::Matches($x, 'x:Name="([A-Za-z0-9_]+)"')) {
        $n = $m.Groups[1].Value
        if ($names.ContainsKey($n)) { continue }
        if ($script:Controls.ContainsKey($n)) { throw "Tab.xaml: Name '$n' gibt es im Programm schon" }
        $el = $tab.FindName($n); if ($el) { $names[$n] = $el }
    }
    foreach ($n in $names.Keys) { $script:Controls[$n] = $names[$n] }
    if (-not $tab.Name) { $tab.Name = "tabAddon$($Info.Id)" }
    if ($Info.Header -and -not $tab.Header) { $tab.Header = $Info.Header }
    if (-not $tab.Style) { try { $tab.Style = $script:Window.FindResource('DarkTabItem') } catch { } }
    $tab.Tag = 'addon'
    try { Set-HUButtonFit $tab } catch { }   # Knoepfe wachsen mit ihrer Beschriftung (global)
    try { Set-HUOutputTools $tab } catch { }  # Ausgaben: Umkehren/Loeschen, neuestes oben (global)
    $script:Controls[$tab.Name] = $tab
    $script:AddonByTab[$tab.Name] = $Info
    if ($tab.Header -is [string]) { $script:AddonHeader[$Info.Id] = $tab.Header }
    # Platzhalter "Start" weg, sobald ein Modul da ist
    $tc = $script:Controls['tabMain']
    if ($script:Controls['tabStart'] -and $tc.Items.Contains($script:Controls['tabStart'])) { $tc.Items.Remove($script:Controls['tabStart']) }
    [void]$tc.Items.Add($tab)
}

# Beim Start: wartende Fassungen uebernehmen und die Module sammeln, in Reiter-Reihenfolge.
# Main.ps1 laedt die Skripte selbst per Dot-Source, damit ihre Funktionen im Programm-Bereich liegen.
function Get-HUAddonStartList {
    $dir = Get-HUAddonsDir
    try { Use-HUAddonPending $dir } catch { Write-HULogWarn "Module: Uebernahme fehlgeschlagen: $($_.Exception.Message)" }
    $tp = "$($script:App.signer)"
    $ver = [Version]$script:Version
    # mitgelieferte Module gelten wie Functions\ (ohne eigene Signatur); Abschalten je Modul: settings.json "modulesOff"
    $off = @(); try { if ($script:Settings -and $script:Settings.PSObject.Properties['modulesOff']) { $off = @($script:Settings.modulesOff | ForEach-Object { "$_" }) } } catch { }
    $lists = @()
    foreach ($i in @(Get-HUAddonsInDir (Join-Path $script:AppRoot 'Modules') $ver 'Programm')) {
        if ($i.Error) { Write-HULogWarn "Modul $(Split-Path $i.Dir -Leaf) (mitgeliefert) nicht geladen: $($i.Error)"; continue }
        if ($off -contains $i.Id) { Write-HULogInfo "Modul $($i.Id) abgeschaltet (Einstellung modulesOff)"; continue }
        $lists += $i
    }
    foreach ($i in @(Get-HUAddonsInDir $dir $ver 'Privat' -Signed -Thumbprint $tp)) {
        if ($i.Error) { Write-HULogWarn "Modul $(if ($i.Id) { $i.Id } else { Split-Path $i.Dir -Leaf }) nicht geladen: $($i.Error)"; continue }
        $lists += $i
    }
    # Entwicklung: AddonsDev\<Id>\ OHNE Signatur, ersetzt ein Modul gleicher Kennung. Ordner mit _ oder . am Anfang = abgeschaltet.
    # Kein zusaetzliches Risiko: wer in den Programmordner schreiben kann, kann auch Main.ps1 aendern.
    foreach ($i in @(Get-HUAddonsInDir (Join-Path $script:AppRoot 'AddonsDev') $ver 'Entwicklung' -SkipDisabled)) {
        if ($i.Error) { Write-HULogWarn "Modul $(Split-Path $i.Dir -Leaf) (Entwicklung) nicht geladen: $($i.Error)"; continue }
        Write-HULogWarn "Modul $($i.Id) $($i.Version) aus AddonsDev geladen - Entwicklung, ohne Signaturpruefung"
        $lists += $i
    }
    $all = @(Merge-HUAddonList $lists)
    $script:AllAddons = $all
    # fuer diese Installation abgeschaltet (Einstellungen > Module, Config\installation.json "moduleAus")
    $instOff = @(); try { if ($script:Installation) { $instOff = @($script:Installation.ModulesOff) } } catch { }
    foreach ($i in @($all | Where-Object { $instOff -contains $_.Id })) { Write-HULogInfo "Modul $($i.Title) ist für diese Installation abgeschaltet" }
    return @($all | Where-Object { $instOff -notcontains $_.Id })
}

# Beim Beenden: Module mit Hintergrundarbeit stoppen (Stop-HUAddon<Id>, optional)
function Stop-HUAddons {
    foreach ($a in @($script:LoadedAddons)) {
        $fn = "Stop-HUAddon$($a.Id)"
        if (Get-Command $fn -CommandType Function -ErrorAction SilentlyContinue) {
            try { & $fn } catch { try { Write-HULogWarn "Modul $($a.Id) beim Beenden: $($_.Exception.Message)" } catch { } }
        }
    }
}

# ============================================================================
# Dateien hineinziehen: das Hauptfenster nimmt Dateien an und gibt sie an das Modul im aktiven Reiter
# (Invoke-HUAddonDrop<Id> -Paths). Passt der aktive Reiter nicht, aber genau ein anderes Modul, wechselt die App dorthin.
# ============================================================================
function Get-HUDropFiles($e) {
    try {
        if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return @() }
        return @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    } catch { return @() }
}
# Anhaenge direkt aus Outlook (ab 1.0.4): in einen eigenen Temp-Ordner schreiben, alte Ordner (aelter als 1 Tag) wegraeumen
function Save-HUDropOutlookFiles($e) {
    if (-not (Get-Command Save-HUDroppedFiles -ErrorAction SilentlyContinue)) { return @() }
    $base = [IO.Path]::GetTempPath()
    try {
        foreach ($old in @(Get-ChildItem -LiteralPath $base -Directory -Filter 'HU-Amt-Drop-*' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) })) {
            Remove-Item -LiteralPath $old.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { }
    $dir = Join-Path $base ('HU-Amt-Drop-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $prob = New-Object 'System.Collections.Generic.List[string]'
    $files = @(Save-HUDroppedFiles -Data $e.Data -TargetDir $dir -Problems $prob)
    if ($prob.Count) { Show-HUMessage ("Aus Outlook nicht übernommen:`n" + ($prob -join "`n")) -Icon Warning }
    return $files
}
function Test-HUAddonCanDrop($Info) {
    return ($Info -and @($Info.Drop).Count -and (Get-Command "Invoke-HUAddonDrop$($Info.Id)" -CommandType Function -ErrorAction SilentlyContinue))
}
# Ziel fuer die Dateien bestimmen: @{ Info; Files; Switch } oder $null
function Resolve-HUDropTarget([string[]]$Files) {
    if (-not @($Files).Count) { return $null }
    $sel = $script:Controls['tabMain'].SelectedItem
    $cur = $(if ($sel -and $script:AddonByTab.ContainsKey("$($sel.Name)")) { $script:AddonByTab["$($sel.Name)"] } else { $null })
    if (Test-HUAddonCanDrop $cur) {
        $f = @($Files | Where-Object { Test-HUAddonDropMatch $cur $_ })
        if ($f.Count) { return @{ Info = $cur; Files = $f; Switch = $false } }
    }
    $cands = @()
    foreach ($a in @($script:LoadedAddons)) {
        if ((Test-HUAddonCanDrop $a) -and @($Files | Where-Object { Test-HUAddonDropMatch $a $_ }).Count) { $cands += $a }
    }
    if ($cands.Count -eq 1) { return @{ Info = $cands[0]; Files = @($Files | Where-Object { Test-HUAddonDropMatch $cands[0] $_ }); Switch = $true } }
    return $null
}
function Register-HUFileDrop {
    $w = $script:Window
    $w.AllowDrop = $true
    $over = {
        param($s, $e)
        # nur Dateien behandeln: alles andere (z. B. ein gezogener Reiter) muss bis zum Ziel durchkommen,
        # sonst sperrt Effects=None + Handled das Verschieben der Reiter
        $outlook = $false
        if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            # Anhang direkt aus Outlook: nur die Namen sind schon bekannt, der Inhalt kommt erst beim Loslassen
            if (-not ((Get-Command Get-HUDropNames -ErrorAction SilentlyContinue) -and $e.Data.GetDataPresent('FileGroupDescriptorW'))) { return }
            $outlook = $true
        }
        $fl = $(if ($outlook) { @(Get-HUDropNames $e.Data) } else { @(Get-HUDropFiles $e) })
        $t = $(if (@($fl | Where-Object { Test-HUAddonKeyFile $_ }).Count) { $true } else { Resolve-HUDropTarget $fl })
        $e.Effects = $(if ($t) { [System.Windows.DragDropEffects]::Copy } else { [System.Windows.DragDropEffects]::None })
        $e.Handled = $true
    }
    $w.Add_PreviewDragEnter($over)
    $w.Add_PreviewDragOver($over)
    $w.Add_PreviewDrop({
            param($s, $e)
            $files = @(Get-HUDropFiles $e)
            if (-not $files.Count -and -not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop) -and $e.Data.GetDataPresent('FileGroupDescriptorW')) {
                # erst pruefen, ob ein Reiter die Namen nimmt, dann die Anhaenge aus Outlook als Dateien holen
                $names = @(Get-HUDropNames $e.Data)
                if (-not $names.Count -or -not (@($names | Where-Object { Test-HUAddonKeyFile $_ }).Count -or (Resolve-HUDropTarget $names))) { return }
                $e.Handled = $true
                $files = @(Save-HUDropOutlookFiles $e)
            }
            if (-not $files.Count) { return }
            $e.Handled = $true
            # Modul-Dateien (*.humodule) behandelt der Kern selbst, kein Modul bekommt sie
            $keys = @($files | Where-Object { Test-HUAddonKeyFile $_ })
            if ($keys.Count) {
                foreach ($k in $keys) { Import-HUAddonKeyFile $k }
                $files = @($files | Where-Object { -not (Test-HUAddonKeyFile $_) })
                if (-not $files.Count) { return }
            }
            $t = Resolve-HUDropTarget $files
            if (-not $t) {
                $ext = @($files | ForEach-Object { [IO.Path]::GetExtension($_).ToLowerInvariant() } | Sort-Object -Unique) -join ', '
                Show-HUMessage "Diese Dateien ($ext) kann hier kein Reiter verarbeiten.`n`nBitte den passenden Reiter wählen und die Dateien dort hineinziehen." -Icon Info
                return
            }
            if ($t.Switch) {
                $tab = @($script:Controls['tabMain'].Items | Where-Object { $script:AddonByTab["$($_.Name)"] -eq $t.Info })[0]
                if ($tab) { $script:Controls['tabMain'].SelectedItem = $tab }
            }
            $skip = $files.Count - @($t.Files).Count
            if ($skip -gt 0) { Write-HULogWarn "$skip Datei(en) passen nicht zu $($t.Info.Title) und wurden übergangen." }
            try { & "Invoke-HUAddonDrop$($t.Info.Id)" -Paths @($t.Files) }
            catch { Write-HULogError "$($t.Info.Title): Dateien nicht übernommen: $($_.Exception.Message)" }
        })
}

# ============================================================================
# Weitergeben zwischen Modulen (Postfach im gemeinsamen Datenordner, Logik in Core-Addons.ps1)
#   Get-HUAddonTargets -Extension '.pdf' [-Exclude <Id>]     Module, die diese Endung annehmen: @{ Id; Title }
#   Send-HUAddonFiles -To <Id> -Paths <...> -From <Id> [-Open] Anzahl der weitergegebenen Dateien
#       ohne -Open: Kopie ins Postfach (geht auch an Bediener auf anderen PCs), Zahl am Reiter
#       mit -Open:  zum Reiter wechseln und die Dateien sofort uebergeben (Invoke-HUAddonDrop<Id>)
#   Get-HUAddonInbox <Id>                Eintraege @{ Key; Path; Name; From; Pc; User; Time }
#   Remove-HUAddonInboxItem <Eintrag>    Eintrag samt Kopie entfernen (erst nach gelungener Uebernahme)
#   Update-HUAddonInboxBadge [<Id>]      Zahl am Reiter neu zaehlen (der Kern prueft zusaetzlich jede Minute)
# ============================================================================
function Get-HUAddonLoaded([string]$Id) { return @($script:LoadedAddons | Where-Object { $_.Id -ceq $Id })[0] }

function Get-HUAddonTargets([string]$Extension, [string]$Exclude = '') {
    $x = "$Extension".Trim().ToLowerInvariant(); if ($x -and -not $x.StartsWith('.')) { $x = ".$x" }
    $out = @()
    foreach ($a in @($script:LoadedAddons)) {
        if ($Exclude -and $a.Id -ceq $Exclude) { continue }
        if (-not (Test-HUAddonCanDrop $a)) { continue }
        if ($x -and @($a.Drop) -notcontains $x) { continue }
        $out += [pscustomobject]@{ Id = $a.Id; Title = $(if ($a.Title) { $a.Title } else { $a.Id }) }
    }
    return $out
}

function Send-HUAddonFiles([string]$To, [string[]]$Paths, [string]$From = '', [switch]$Open) {
    $t = Get-HUAddonLoaded $To
    if (-not (Test-HUAddonCanDrop $t)) { throw "Modul '$To' ist nicht geladen oder nimmt keine Dateien an" }
    $files = @($Paths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) -and (Test-HUAddonDropMatch $t $_) })
    $skip = @($Paths).Count - $files.Count
    if ($skip -gt 0) { Write-HULogWarn "$skip Datei(en) passen nicht zu $($t.Title) oder fehlen und wurden nicht weitergegeben." }
    if (-not $files.Count) { return 0 }
    if ($Open) {
        $tab = @($script:Controls['tabMain'].Items | Where-Object { $script:AddonByTab["$($_.Name)"] -eq $t })[0]
        if ($tab) { $script:Controls['tabMain'].SelectedItem = $tab }
        & "Invoke-HUAddonDrop$($t.Id)" -Paths @($files)
        return $files.Count
    }
    $n = 0
    foreach ($f in $files) {
        try { [void](Add-HUAddonInboxItem -DataRoot $script:DataRoot -To $t.Id -Path $f -From $From -User $script:UserName); $n++ }
        catch { Write-HULogWarn "$([IO.Path]::GetFileName($f)) nicht an $($t.Title) weitergegeben: $($_.Exception.Message)" }
    }
    if ($n) { Write-HULogInfo "$n Datei(en) an $($t.Title) weitergegeben$(if ($From) { " (von $From)" })" }
    Update-HUAddonInboxBadge $t.Id
    return $n
}

function Get-HUAddonInbox([string]$Id) {
    $t = Get-HUAddonLoaded $Id
    if (-not $t) { return @() }
    return @(Get-HUAddonInboxItems -DataRoot $script:DataRoot -Id $t.Id -Extensions @($t.Drop))
}

function Remove-HUAddonInboxItem($Item) {
    Remove-HUAddonInboxEntry -DataRoot $script:DataRoot -Item $Item
    try { Update-HUAddonInboxBadge "$($Item.Inbox)" } catch { }
}

# Zahl der wartenden Dateien an den Reiter schreiben ("Titel (3)"). Nur bei Reitern mit Text als Beschriftung.
function Update-HUAddonInboxBadge([string]$Id = '') {
    foreach ($a in @($script:LoadedAddons)) {
        if ($Id -and $a.Id -cne $Id) { continue }
        if (-not $script:AddonHeader.ContainsKey($a.Id) -or -not (Test-HUAddonCanDrop $a)) { continue }
        $tab = @($script:Controls['tabMain'].Items | Where-Object { $script:AddonByTab["$($_.Name)"] -eq $a })[0]
        if (-not $tab) { continue }
        $n = 0; try { $n = @(Get-HUAddonInbox $a.Id).Count } catch { }
        $h = $script:AddonHeader[$a.Id]
        $tab.Header = $(if ($n) { "$h ($n)" } else { $h })
    }
}

# Beim Start: Zahl einmal setzen, dann jede Minute (andere PCs legen auch ins Postfach)
function Start-HUAddonInboxWatch {
    try { Update-HUAddonInboxBadge } catch { }
    if ($script:AddonInboxTimer) { return }
    $script:AddonInboxTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:AddonInboxTimer.Interval = [TimeSpan]::FromSeconds(60)
    $script:AddonInboxTimer.Add_Tick({ try { Update-HUAddonInboxBadge } catch { } })
    $script:AddonInboxTimer.Start()
}

# ============================================================================
# Gemeinsame Bausteine fuer Module
# ============================================================================
# Gruener Punkt pulsiert, solange etwas laedt (HU-Standard). $Dot: Ellipse aus dem Tab.xaml des Moduls.
function Set-HUBusyDot($Dot, [bool]$On) {
    if (-not $Dot) { return }
    if ($On) {
        if ($Dot.Visibility -eq 'Visible') { return }
        $an = New-Object System.Windows.Media.Animation.DoubleAnimation
        $an.From = 1.0; $an.To = 0.2; $an.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(700))
        $an.AutoReverse = $true; $an.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
        $Dot.Visibility = 'Visible'
        $Dot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $an)
    } else {
        $Dot.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $null)
        $Dot.Opacity = 1.0; $Dot.Visibility = 'Collapsed'
    }
}

# ============================================================================
# Private Module: im Hintergrund die neueste signierte Fassung holen (wirkt nach dem Neustart)
# ============================================================================
function Start-HUAddonUpdate([switch]$Loud) {
    $src = @(Get-HUAddonSources)
    if (-not $src.Count) { if ($Loud) { Write-HULogInfo 'Module: kein privates Modul hinterlegt' }; return }
    $cfg = $null; try { $cfg = Update-UpdateConfig } catch { }
    $tp = if ($cfg -and $cfg.SignerThumbprint) { $cfg.SignerThumbprint } else { "$($script:App.signer)" }
    $chan = if ($cfg) { "$($cfg.Channel)" } else { 'Stable' }
    $script:AddonUpdateLoud = [bool]$Loud
    # sichtbar machen, dass geladen wird (kann bei mehreren Modulen einige Minuten dauern)
    if ($Loud) {
        Write-HULogInfo "Module: $($src.Count) werden geprüft und geladen ..."
        try { Show-HUToast -Title 'Module werden geladen' -Text "$($src.Count) Module werden geprüft und geladen. Das kann einige Minuten dauern, danach kommt eine Meldung." -Icon Info } catch { }
    }
    Invoke-AsyncCommand -ScriptBlock {
        param($lib, $addonLib, $pathsMod, $sources, $dir, $ver, $chan, $tp)
        Import-Module $pathsMod -DisableNameChecking
        . $lib; . $addonLib
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
        $out = @()
        foreach ($s in @($sources)) {
            $r = Install-HUAddon -Owner $s.Owner -Repo $s.Repo -Token $s.Token -AddonsDir $dir -AppVersion ([Version]$ver) -Channel $chan -Thumbprint $tp
            $out += "$($r.Status)|$($s.Owner)/$($s.Repo)|$($r.Text)"
        }
        return ($out -join "`n")
    } -ArgumentList @($script:UpdateLib, (Join-Path $script:AppRoot 'Functions\Core-Addons.ps1'), (Join-Path $script:AppRoot 'Core\HU.Paths.psm1'), $src, (Get-HUAddonsDir), "$($script:Version)", $chan, $tp) -TimeoutSec 300 -BusyText 'Module werden geladen ...' -OnComplete {
        param($result)
        $neu = @()
        foreach ($ln in @("$result" -split "`n" | Where-Object { $_ })) {
            $p = $ln -split '\|', 3
            switch ($p[0]) {
                'Neu' { Write-HULogOK "Modul: $($p[2])"; $neu += ($p[2] -replace ' geladen - .*$', '') }
                'Aktuell' { if ($script:AddonUpdateLoud) { Write-HULogInfo "Modul: $($p[2])" } }
                default { Write-HULogWarn "Modul $($p[1]): $($p[2])" }
            }
        }
        if ($neu.Count) { Show-HUAddonRestartPrompt $neu }
        elseif ($script:AddonUpdateLoud) {
            $warn = @("$result" -split "`n" | Where-Object { $_ -and $_ -notmatch '^(Neu|Aktuell)\|' }).Count
            try { Show-HUToast -Title 'Module geladen' -Text $(if ($warn) { "Nicht alle Module ließen sich laden. Details im Protokoll." } else { 'Alle Module sind aktuell.' }) -Icon $(if ($warn) { 'Warning' } else { 'Success' }) } catch { }
        }
    }
}

# Fenster (nur Verwalter): privates Modul hinterlegen (Repository + Nur-Lese-Token) oder entfernen
function Show-HUAddonDialog {
    $x = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Module" Width="660" MinWidth="660" SizeToContent="Height" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#1E1E1E" ShowInTaskbar="False">
    <Window.Resources>
        <!--HU:THEME-->
    </Window.Resources>
    <StackPanel Margin="16">
        <TextBlock Style="{StaticResource HintText}" TextWrapping="Wrap" Margin="0,0,0,10"
                   Text="Private Module liegen in privaten Repositories. Mit dem Nur-Lese-Token, das du vom Herausgeber bekommst, lädt die App das Modul und hält es aktuell. Geladen wird nur, was mit dem Zertifikat des Herausgebers signiert ist. Der Token wird verschlüsselt nur für deinen Windows-Benutzer gespeichert."/>
        <TextBlock Text="Hinterlegt" Foreground="#AAAAAA" FontSize="11"/>
        <ListBox x:Name="lstSrc" Height="80" Margin="0,2,0,10" Background="#252526" Foreground="#CCCCCC" BorderBrush="#3E3E42" ToolTip="Hinterlegte Repositories"/>
        <TextBlock Text="Repository (Besitzer/Name)" Foreground="#AAAAAA" FontSize="11"/>
        <TextBox x:Name="txtRepo" Style="{StaticResource DarkTextBox}" Margin="0,2,0,8" ToolTip="z. B. Besitzer/Modulname"/>
        <TextBlock Text="Nur-Lese-Token" Foreground="#AAAAAA" FontSize="11"/>
        <PasswordBox x:Name="pwdTok" Margin="0,2,0,12" Background="#2D2D30" Foreground="White" BorderBrush="#3E3E42" Padding="4,3" ToolTip="Token vom Herausgeber (wird verschlüsselt gespeichert)"/>
        <WrapPanel HorizontalAlignment="Right">
            <Button x:Name="btnFile" Content="Datei einlesen ..." Width="140" Background="#0E639C" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Modul-Datei (*.humodule) mit Repositories und Token einlesen. Geht auch per Hineinziehen ins Hauptfenster."/>
            <Button x:Name="btnDel" Content="Markiertes entfernen" Width="160" Background="#C62828" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Token des markierten Repositories löschen"/>
            <Button x:Name="btnAdd" Content="Hinterlegen und laden" Width="170" Background="#388E3C" Style="{StaticResource DarkButton}" Margin="0,0,8,0" ToolTip="Token speichern und das Modul laden"/>
            <Button x:Name="btnClose" Content="Schließen" Width="100" Background="#555555" Style="{StaticResource DarkButton}" IsCancel="True" ToolTip="Fenster schließen"/>
        </WrapPanel>
    </StackPanel>
</Window>
'@
    $d = New-HUWindow -XamlText (Get-HUThemedXaml $x)
    $w = $d.Window; $c = $d.C
    $fill = {
        $c.lstSrc.Items.Clear()
        $loaded = @($script:LoadedAddons | ForEach-Object { "$($_.Title) $($_.Version)" })
        foreach ($s in @(Get-HUAddonSources)) { [void]$c.lstSrc.Items.Add("$($s.Owner)/$($s.Repo)") }
        if ($loaded.Count) { $w.Title = "Module - geladen: $($loaded -join ', ')" }
    }
    & $fill
    $c.btnAdd.Add_Click({
            $parts = "$($c.txtRepo.Text)".Trim() -split '/', 2
            if ($parts.Count -ne 2 -or -not (Test-HUAddonRepoName $parts[0] $parts[1])) { Show-HUMessage 'Repository bitte als Besitzer/Name eingeben.' -Icon Info -Owner $w; return }
            if (-not $c.pwdTok.SecurePassword.Length) { Show-HUMessage 'Bitte den Token eingeben.' -Icon Info -Owner $w; return }
            try { Save-HUAddonToken $parts[0] $parts[1] $c.pwdTok.SecurePassword } catch { Show-HUMessage "Nicht gespeichert: $($_.Exception.Message)" -Icon Warning -Owner $w; return }
            $c.pwdTok.Clear(); $c.txtRepo.Text = ''
            & $fill
            Write-HULogInfo "Modul $($parts[0])/$($parts[1]) hinterlegt - wird geladen ..."
            Start-HUAddonUpdate -Loud
        })
    $c.btnFile.Add_Click({
            $dlg = New-Object Microsoft.Win32.OpenFileDialog
            $dlg.Filter = "Modul-Datei (*$($script:HUAddonKeyFileExt))|*$($script:HUAddonKeyFileExt)"
            if (-not $dlg.ShowDialog($w)) { return }
            Import-HUAddonKeyFile $dlg.FileName -Owner $w
            $w.Close()
        })
    $c.btnDel.Add_Click({
            $sel = "$($c.lstSrc.SelectedItem)"
            if (-not $sel) { return }
            if (-not (Confirm-HU -Warning "Token für $sel entfernen? Das Modul wird dann nicht mehr aktualisiert; der Ordner unter Addons bleibt, bis du ihn löschst." 'Module')) { return }
            $p = $sel -split '/', 2; Remove-HUAddonToken $p[0] $p[1]; & $fill
        })
    [void]$w.ShowDialog()
}

# Modul-Datei (*.humodule) einlesen (nur Verwalter): alle Repositories mit Token pruefen und hinterlegen.
# Die Datei wird geloescht, wenn alle Eintraege gespeichert sind; sonst bleibt sie und die Meldung nennt, was fehlt.
# Tokens erscheinen nie in Meldungen oder im Protokoll.
# -FromPackage: Zugang aus einem Paket fuer eine Aussenstelle (Pruefsumme schon geprueft) - ohne Verwalter-Pruefung und ohne Rueckfrage.
function Import-HUAddonKeyFile([string]$Path, $Owner = $null, [switch]$FromPackage) {
    $name = [IO.Path]::GetFileName($Path)
    if (-not $FromPackage -and -not (Test-HUVerwalter)) { Show-HUMessage "Module eintragen dürfen nur Verwalter.`n`nDie Datei $name bleibt unverändert." -Icon Info -Owner $Owner; return }
    try { $items = @(Read-HUAddonKeyFile $Path) }
    catch { Write-HULogWarn "Modul-Datei $($name): $($_.Exception.Message)"; Show-HUMessage "$name kann nicht gelesen werden:`n$($_.Exception.Message)" -Icon Warning -Owner $Owner; return }
    $list = ($items | ForEach-Object { if ($_.Error) { "  - $($_.Error)" } else { "  - $($_.Owner)/$($_.Repo)" } }) -join "`n"
    if (-not $FromPackage -and -not (Confirm-HU "Diese Module aus $name eintragen?`n`n$list`n`nVorhandene Einträge für dieselben Repositories werden ersetzt. Ist alles eingetragen, wird die Datei gelöscht." 'Module' -Owner $Owner)) { return }
    $ok = @($items | Where-Object { -not $_.Error })
    if (-not $ok.Count) { Show-HUMessage "In $name ist kein brauchbarer Eintrag.`n`n$list" -Icon Warning -Owner $Owner; return }
    Write-HULogInfo "Modul-Datei $($name): $($ok.Count) Repository(s) werden geprüft ..."
    $src = @($ok | ForEach-Object { [pscustomobject]@{ Owner = $_.Owner; Repo = $_.Repo; Token = $_.Token } })
    Invoke-AsyncCommand -ScriptBlock {
        param($addonLib, $lib, $pathsMod, $sources)
        Import-Module $pathsMod -DisableNameChecking
        . $lib; . $addonLib
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
        $out = @()
        foreach ($s in @($sources)) { $out += "$($s.Owner)/$($s.Repo)|$(Test-HUAddonTokenAccess $s.Owner $s.Repo $s.Token)" }
        return ($out -join "`n")
    } -ArgumentList @((Join-Path $script:AppRoot 'Functions\Core-Addons.ps1'), $script:UpdateLib, (Join-Path $script:AppRoot 'Core\HU.Paths.psm1'), $src) -TimeoutSec 120 -BusyText 'Module werden geprüft ...' -State @{ Path = $Path; Name = $name; Items = $items; Owner = $Owner; FromPackage = [bool]$FromPackage } -OnComplete {
        param($result, $st)
        $why = @{}
        foreach ($ln in @("$result" -split "`n" | Where-Object { $_ -match '\|' })) { $p = $ln -split '\|', 2; $why[$p[0]] = $p[1] }
        $done = @(); $fail = @()
        foreach ($i in @($st.Items)) {
            $key = "$($i.Owner)/$($i.Repo)"
            if ($i.Error) { $fail += "$(if ($i.Owner) { $key } else { 'Eintrag' }): $($i.Error)"; continue }
            if (-not $why.ContainsKey($key)) { $fail += "$($key): nicht geprüft ($result)"; continue }
            if ($why[$key]) { $fail += "$($key): $($why[$key])"; continue }
            try { $ss = New-Object Security.SecureString; foreach ($ch in "$($i.Token)".ToCharArray()) { $ss.AppendChar($ch) }; $ss.MakeReadOnly()
                Save-HUAddonToken $i.Owner $i.Repo $ss; $done += $key }
            catch { $fail += "$($key): nicht gespeichert ($($_.Exception.Message))" }
        }
        foreach ($d in $done) { Write-HULogOK "Modul hinterlegt: $d" }
        foreach ($f in $fail) { Write-HULogWarn "Modul-Datei: $f" }
        $txt = ''
        if ($done.Count) { $txt += "Eingetragen:`n  $($done -join "`n  ")`n`n" }
        if ($fail.Count) {
            $txt += "Nicht eingetragen:`n  $($fail -join "`n  ")`n`nDie Datei $($st.Name) bleibt erhalten."
        } else {
            try { Remove-Item -LiteralPath $st.Path -Force -ErrorAction Stop; $txt += "Die Datei $($st.Name) wurde gelöscht."; Write-HULogInfo "Modul-Datei $($st.Name) gelöscht" }
            catch { $txt += "Die Datei $($st.Name) konnte nicht gelöscht werden, bitte selbst löschen:`n$($_.Exception.Message)"; Write-HULogWarn "Modul-Datei $($st.Name) nicht gelöscht: $($_.Exception.Message)" }
        }
        if ($st.FromPackage) {
            # Paket einer Aussenstelle: kurze Meldung fuer Laien, keine Repository-Namen (Details im Protokoll)
            if ($fail.Count) { Show-HUMessage "Die Updates für die Module konnten noch nicht eingerichtet werden (keine Internetverbindung?).`n`nHU-Amt versucht es beim nächsten Start wieder. Arbeiten kannst du trotzdem ganz normal." -Icon Info -Title 'Module' }
            else { try { Show-HUToast -Title 'Module' -Text 'Updates für die Module sind eingerichtet.' -Icon Success } catch { } }
        } else { Show-HUMessage $txt.Trim() -Icon $(if ($fail.Count) { 'Warning' } else { 'Info' }) -Title 'Module' }
        if ($done.Count) { Start-HUAddonUpdate -Loud:(-not $st.FromPackage) }
    }
}

# Hinweis nach dem Laden einer neuen Fassung: jetzt neu starten?
function Show-HUAddonRestartPrompt([string[]]$Names) {
    $txt = "Modul geladen:`n`n  $($Names -join "`n  ")`n`nEs wird nach einem Neustart von $($script:AppTitle) aktiv. Jetzt neu starten?"
    if (-not (Confirm-HU $txt 'Module')) { Write-HULogInfo 'Modul wird beim nächsten Start geladen.'; return }
    Restart-HUApp
}

# App beenden und neu starten: die neue Instanz startet sofort, zeigt "wird neu gestartet" und wartet,
# bis diese beendet ist (Main.ps1 -WaitForPid).
# -EnterTestMode / -LeaveTestMode: Testmodus starten bzw. beenden; sonst bleibt ein laufender Testmodus erhalten.
function Restart-HUApp([switch]$EnterTestMode, [switch]$LeaveTestMode, [switch]$EnterDemo, [switch]$LeaveDemo) {
    try {
        $extra = "-WaitForPid $PID"
        $demo = (Test-HUDemoMode)
        if ($EnterDemo -or ($demo -and -not $LeaveDemo)) { $extra += ' -Demo' }
        elseif ($EnterTestMode -or ((Test-HUTestMode) -and -not $demo -and -not $LeaveTestMode)) { $extra += ' -TestMode' }
        # Neustart-Fenster im gespeicherten Farbschema (auch direkt nach dem Umstellen)
        try { if ("$($script:Settings.ui.theme)" -eq 'Hell') { $extra += ' -Hell' } } catch { }
        $si = Get-HUAppStartInfo $extra
        Start-Process -FilePath $si.Exe -WorkingDirectory $script:AppRoot -WindowStyle Hidden -ArgumentList $si.Args | Out-Null
        Write-HULogOK "$($script:AppTitle) wird neu gestartet ..."
        $script:Window.Close()
    } catch { Write-HULogError "Neustart fehlgeschlagen: $($_.Exception.Message)" }
}
