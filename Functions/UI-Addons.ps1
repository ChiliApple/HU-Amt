#Requires -Version 5.1
<#
.SYNOPSIS
    Module in der Oberflaeche: Reiter einbauen, Dateien hineinziehen, beim Beenden stoppen, Neustart.
    Dazu gemeinsame Bausteine fuer Module (pulsierender gruener Punkt beim Laden).
.NOTES
    Dot-Source aus Main.ps1 (nach Core-Addons.ps1). Zielmaschine: der PC, auf dem die App laeuft.
#>

$script:LoadedAddons = @()
$script:AddonByTab = @{}

# Reiter eines (bereits gepruefen) Moduls in das Hauptfenster einbauen und die benannten Elemente registrieren
function Add-HUAddonTab($Info) {
    $x = [IO.File]::ReadAllText((Join-Path $Info.Dir $Info.Xaml), [Text.Encoding]::UTF8)
    $theme = Get-HUXaml 'Theme'
    $mt = [regex]::Match($theme, '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>')
    $x = $x.Replace('<!--HU:THEME-->', $mt.Groups[1].Value)
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
    return @(Merge-HUAddonList $lists)
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
        $t = Resolve-HUDropTarget (Get-HUDropFiles $e)
        $e.Effects = $(if ($t) { [System.Windows.DragDropEffects]::Copy } else { [System.Windows.DragDropEffects]::None })
        $e.Handled = $true
    }
    $w.Add_PreviewDragEnter($over)
    $w.Add_PreviewDragOver($over)
    $w.Add_PreviewDrop({
            param($s, $e)
            $files = @(Get-HUDropFiles $e)
            if (-not $files.Count) { return }
            $e.Handled = $true
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

# Hinweis nach dem Laden einer neuen Fassung: jetzt neu starten?
function Show-HUAddonRestartPrompt([string[]]$Names) {
    $txt = "Modul geladen:`n`n  $($Names -join "`n  ")`n`nEs wird nach einem Neustart von $($script:AppTitle) aktiv. Jetzt neu starten?"
    if (-not (Confirm-HU $txt 'Module')) { Write-HULogInfo 'Modul wird beim nächsten Start geladen.'; return }
    Restart-HUApp
}

# App beenden und neu starten: die neue Instanz startet sofort, zeigt "wird neu gestartet" und wartet,
# bis diese beendet ist (Main.ps1 -WaitForPid).
function Restart-HUApp {
    try {
        $si = Get-HUAppStartInfo "-WaitForPid $PID"
        Start-Process -FilePath $si.Exe -WorkingDirectory $script:AppRoot -WindowStyle Hidden -ArgumentList $si.Args | Out-Null
        Write-HULogOK "$($script:AppTitle) wird neu gestartet ..."
        $script:Window.Close()
    } catch { Write-HULogError "Neustart fehlgeschlagen: $($_.Exception.Message)" }
}
