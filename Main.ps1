#Requires -Version 5.1
<#
.SYNOPSIS
    HU-Amt - Werkzeuge fuer die Gemeindeverwaltung (WPF, ein Reiter je Modul).
.DESCRIPTION
    Startpunkt. Liest Config\app.json (Name, Repo, Signatur), laedt Core-Module (Core\*.psm1),
    Oberflaeche (XAML\*.xaml) und Funktionen (Functions\*.ps1).
      Config\      app.json, version.json (im Repo); lokale Dateien wie update.json nicht im Repo
      Core\        Logging
      Functions\   Oberflaeche: Dialoge, Fensterzustand, Startbildschirm, Starter und Verknuepfungen
      XAML\        Fenster und gemeinsames Theme
      Data\        wird vom Update nie angefasst; Persoenliches in Data\Benutzer\<Name>\
    Start: HU-Amt.exe (Desktop-Verknuepfung) oder Start-HU-Amt.cmd.
.NOTES
    Zielmaschine: der PC, auf dem HU-Amt laeuft (Windows 10/11, Windows PowerShell 5.1 oder PowerShell 7).
#>
param([switch]$ShowConsole, [string]$SmokeTest = '', [int]$WaitForPid = 0)

# Konsolenfenster ausblenden (nur die WPF-Oberflaeche ist sichtbar)
if (-not $ShowConsole) {
    try {
        Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("User32.dll")] public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
' -ErrorAction Stop
        $consolePtr = [Console.Window]::GetConsoleWindow()
        if ($consolePtr -ne [IntPtr]::Zero) { [void][Console.Window]::ShowWindow($consolePtr, 0) }
    } catch { }
}

# Version 1: nicht gesetzte Variablen sind ein Fehler (Tippfehler); fehlende Eigenschaften liefern $null
Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'
$script:AppRoot = $PSScriptRoot

# Zustandsvariablen (StrictMode: vor dem ersten Lesen setzen)
$script:Window = $null; $script:Controls = @{}; $script:AppIcon = $null; $script:UiScale = 1.0
$script:StartupDone = $false; $script:AbRestored = $false; $script:LastDialog = $null
$script:SmokeTestFile = $SmokeTest; $script:SmokeTimer = $null
$script:Settings = $null; $script:App = $null

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

function Show-HUFatal([string]$Text) {
    if ($SmokeTest) { "FEHLER $Text" | Set-Content -LiteralPath $SmokeTest -Encoding UTF8; exit 1 }
    if (Get-Command Close-HUSplash -ErrorAction SilentlyContinue) { Close-HUSplash }
    $t = 'HU-Amt'; try { if ($script:App) { $t = "$($script:App.title)" } } catch { }
    [void][System.Windows.MessageBox]::Show($Text, "$t - Fehler", 'OK', 'Error')
    exit 1
}

# ============================================================================
# 0. APP-MANIFEST (Name, Repo, Signatur - einzige Stelle fuer den Namen)
# ============================================================================
try {
    $script:App = [IO.File]::ReadAllText((Join-Path $script:AppRoot 'Config\app.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
    if (-not "$($script:App.id)" -or -not "$($script:App.title)") { throw 'id oder title fehlt' }
} catch { Show-HUFatal "Config\app.json fehlt oder ist nicht lesbar:`n$($_.Exception.Message)`n`nPull.ps1 ausführen, um die Dateien zu laden." }
$script:AppTitle = "$($script:App.title)"

# Neustart aus der App: die alte Instanz schliesst noch - kleines Fenster zeigen und warten (hoechstens 90 s)
if ($WaitForPid -gt 0 -and -not $SmokeTest) {
    $old = $null; try { $old = Get-Process -Id $WaitForPid -ErrorAction Stop } catch { }
    if ($old) {
        $rw = $null
        try {
            $tx = [System.Security.SecurityElement]::Escape("$($script:AppTitle) wird neu gestartet ...")
            $rw = [System.Windows.Markup.XamlReader]::Parse("<Window xmlns=`"http://schemas.microsoft.com/winfx/2006/xaml/presentation`" WindowStyle=`"None`" ResizeMode=`"NoResize`" AllowsTransparency=`"True`" Background=`"Transparent`" Width=`"360`" Height=`"90`" WindowStartupLocation=`"CenterScreen`" Topmost=`"True`" ShowInTaskbar=`"False`"><Border Background=`"#FF1E1E1E`" CornerRadius=`"10`" BorderBrush=`"#FFB9A88A`" BorderThickness=`"2`"><TextBlock Text=`"$tx`" Foreground=`"#FFEDEDED`" FontSize=`"15`" HorizontalAlignment=`"Center`" VerticalAlignment=`"Center`"/></Border></Window>")
            $rw.Show()
        } catch { $rw = $null }
        $until = (Get-Date).AddSeconds(90)
        while (-not $old.HasExited -and (Get-Date) -lt $until) {
            if ($rw) {
                $fr = New-Object System.Windows.Threading.DispatcherFrame
                [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [System.Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }, $fr)
                [System.Windows.Threading.Dispatcher]::PushFrame($fr)
            }
            [void]$old.WaitForExit(200)
        }
        if ($rw) { try { $rw.Close() } catch { } }
    }
}

# nur eine Instanz je Programmordner und Benutzer
$script:InstanceMutex = $null
if (-not $SmokeTest) {
    try {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $hid = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes("$($script:AppRoot)".ToLowerInvariant()))) -replace '-', '').Substring(0, 16)
        $sha.Dispose()
        $script:InstanceMutex = New-Object System.Threading.Mutex($false, "Local\$($script:App.id)_$hid")
        $got = $false
        try { $got = $script:InstanceMutex.WaitOne(8000) } catch [System.Threading.AbandonedMutexException] { $got = $true }
        if (-not $got) {
            [void][System.Windows.MessageBox]::Show("$($script:AppTitle) ist bereits geöffnet.`n`nBitte das offene Fenster verwenden.", $script:AppTitle, 'OK', 'Information')
            exit 0
        }
    } catch { $script:InstanceMutex = $null }
}

# Update laeuft gerade oder wurde abgebrochen (Pull.ps1 schreibt das Journal waehrend des Ersetzens)
if (Test-Path -LiteralPath (Join-Path $script:AppRoot 'Config\pull-journal.json')) {
    Show-HUFatal "Ein Update läuft gerade oder wurde abgebrochen.`n`nBitte warten, bis das Update fertig ist, oder im Programmordner Pull.ps1 erneut ausführen."
}

# ============================================================================
# DATENORDNER: Gemeinsames in Data\<Modul-Id>\, Persoenliches in Data\Benutzer\<Name>\
# ============================================================================
$script:DataRoot = Join-Path $script:AppRoot 'Data'
$script:UserName = "$env:USERNAME"
$script:UserRoot = Join-Path (Join-Path $script:DataRoot 'Benutzer') $script:UserName
try {
    foreach ($d in $script:DataRoot, $script:UserRoot) { if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) } }
    $probe = Join-Path $script:UserRoot ('.schreibtest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [IO.File]::WriteAllText($probe, 'x'); Remove-Item -LiteralPath $probe -Force
} catch {
    Show-HUFatal "Der Datenordner ist nicht beschreibbar:`n$($script:UserRoot)`n`nBitte Schreibrechte auf den Ordner Data geben lassen.`n`n$($_.Exception.Message)"
}

# ============================================================================
# STARTBILDSCHIRM - im eigenen Thread, erscheint sofort, waehrend Module und Oberflaeche laden
# ============================================================================
$script:Splash = $null; $script:SplashShown = Get-Date; $script:SplashStatus = $null; $script:SplashSync = $null
function Set-HUSplashStatus([string]$Text) {
    if ($script:SplashSync) { $script:SplashSync.Status = $Text }
}
# -Animate: zum Abschied die Startanimation (Functions\UI-Splash.ps1), sonst sofort schliessen
function Close-HUSplash([switch]$Animate) {
    if ($script:SplashSync) { if ($Animate) { $script:SplashSync.Go = $true; $script:SplashSync.MainReady = $true } else { $script:SplashSync.Close = $true }; $script:SplashSync = $null }
}
$script:Version = '0.0.0'
try { $script:Version = "$(([IO.File]::ReadAllText((Join-Path $script:AppRoot 'Config\version.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json).version)" } catch { }
if (-not $SmokeTest) {
    $spf = Join-Path $script:AppRoot 'Functions\UI-Splash.ps1'
    if (Test-Path -LiteralPath $spf) { try { . $spf; $script:SplashSync = Start-HUSplashThread $script:AppRoot $script:Version } catch { $script:SplashSync = $null } }
}

# ============================================================================
# 1. CORE-MODULE, FUNKTIONEN
# ============================================================================
Set-HUSplashStatus 'Module werden geladen ...'
foreach ($mod in @('HU.Logging')) {
    $modPath = Join-Path $script:AppRoot "Core\$mod.psm1"
    if (-not (Test-Path -LiteralPath $modPath)) { Show-HUFatal "Core-Modul fehlt: $modPath`n`nPull.ps1 ausführen, um die Dateien zu laden." }
    Import-Module $modPath -Force -DisableNameChecking
}
foreach ($f in @('UI-Common', 'UI-State', 'UI-Shell', 'UI-Splash')) {
    $fp = Join-Path $script:AppRoot "Functions\$f.ps1"
    if (-not (Test-Path -LiteralPath $fp)) { Show-HUFatal "Datei fehlt: $fp`n`nPull.ps1 ausführen, um die Dateien zu laden." }
    . $fp
}

# ============================================================================
# 2. EINSTELLUNGEN (je Benutzer, ohne Datei gelten die Vorgaben)
# ============================================================================
Set-HUSplashStatus 'Einstellungen ...'
$script:SettingsPath = Join-Path $script:UserRoot 'settings.json'
$script:FirstRun = -not (Test-Path -LiteralPath (Join-Path $script:UserRoot 'ui-state.json'))
$script:Settings = [pscustomobject]@{
    ui      = [pscustomobject]@{ splashAnimation = $true; maxLogLines = 1000 }
    logging = [pscustomobject]@{ logLevel = 'INFO'; retentionDays = 30 }
}
$saved = Read-HUJsonFile $script:SettingsPath
if ($saved) {
    foreach ($sec in 'ui', 'logging') {
        if (-not $saved.PSObject.Properties[$sec]) { continue }
        foreach ($p in $saved.$sec.PSObject.Properties) {
            if ($script:Settings.$sec.PSObject.Properties[$p.Name]) { $script:Settings.$sec.($p.Name) = $p.Value }
        }
    }
}

$iconPath = Join-Path $script:AppRoot 'Assets\icon.ico'
if (Test-Path -LiteralPath $iconPath) { try { $script:AppIcon = [System.Windows.Media.Imaging.BitmapFrame]::Create([System.Uri]::new($iconPath, [System.UriKind]::Absolute)) } catch { } }
# eigene Taskleisten-Gruppe (HU-Symbol statt PowerShell) - vor dem ersten Fenster
[void](Initialize-HUTaskbar)

# ============================================================================
# 3. HAUPTFENSTER
# ============================================================================
Set-HUSplashStatus "Oberfl$([char]0xE4)che wird aufgebaut ..."
if ($script:SplashSync) { $script:SplashSync.Go = $true }
try {
    $d = New-HUWindow 'MainWindow'
} catch { Show-HUFatal "Oberfläche konnte nicht geladen werden (XAML\MainWindow.xaml):`n$($_.Exception.Message)" }
$script:Window = $d.Window
# Fehler in einem Klick-Handler beendet sonst das ganze Programm -> protokollieren und weiterlaufen
$script:Window.Dispatcher.Add_UnhandledException({
        param($s, $e)
        try { Write-HULogError "Unerwarteter Fehler (Programm läuft weiter): $($e.Exception.Message)" } catch { }
        $e.Handled = $true
    })
$script:Controls = $d.C
$c = $script:Controls
if ($script:AppIcon) { $script:Window.Icon = $script:AppIcon; $c['imgLogo'].Source = $script:AppIcon }
# grosses Logo in der Mitte: groesstes Bild aus icon.ico (das erste ist klein und wuerde unscharf hochgezogen)
try { $big = Get-HUSplashIcon $script:AppRoot; if ($big) { $c['imgStartLogo'].Source = $big } elseif ($script:AppIcon) { $c['imgStartLogo'].Source = $script:AppIcon } } catch { }
$script:Window.Add_SourceInitialized({ Set-HUWindowTaskbar $script:Window })
$c['txtAppTitle'].Text = $script:AppTitle
$c['txtAppSubtitle'].Text = "$($script:App.subtitle)"
$c['txtStartTitle'].Text = "Willkommen bei $($script:AppTitle)"

$script:LogDir = Join-Path $script:UserRoot 'Logs'
$logFile = Join-Path $script:LogDir "$($script:App.id)_{date}.log"
$minLevel = "$($script:Settings.logging.logLevel)".ToUpper()
if ($minLevel -notin 'DEBUG', 'INFO', 'OK', 'WARN', 'ERROR') { $minLevel = 'INFO' }
Initialize-Logging -LogFilePath $logFile -RichTextBox $c['rtbLog'] -MinLevel $minLevel -MaxGuiLines ([int]$script:Settings.ui.maxLogLines)
# Tagesprotokolle aelter als logging.retentionDays (Standard 30) beim Start loeschen
try {
    $keep = 30; try { if ([int]$script:Settings.logging.retentionDays -gt 0) { $keep = [int]$script:Settings.logging.retentionDays } } catch { }
    $limit = (Get-Date).AddDays(-$keep)
    Get-ChildItem -LiteralPath $script:LogDir -Filter "$($script:App.id)_*.log" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $limit } | Remove-Item -Force -ErrorAction SilentlyContinue
} catch { }

# Kopfzeile: Version und Titel
function Update-VersionDisplay {
    $t = $script:Controls['txtVersion']
    $t.Text = "v$($script:Version)  |  Testversion"
    $t.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFA726')
    $t.ToolTip = 'Testversion: noch nicht für den Alltag freigegeben.'
    $script:Window.Title = "$($script:AppTitle) v$($script:Version)"
}

# ============================================================================
# 4. EREIGNISSE
# ============================================================================
$soon = { Show-HUMessage 'Diese Funktion kommt in einer der nächsten Testversionen.' -Icon Info }
$c['btnSettings'].Add_Click($soon)
$c['btnSupport'].Add_Click($soon)
$c['btnHelp'].Add_Click($soon)
$c['btnUpdate'].Add_Click($soon)
$c['btnOpenLogFolder'].Add_Click({ Open-HUPath $script:LogDir })
try { Set-HUOutputTools $script:Window } catch { }

# Tastatur: F1 Hilfe, Strg+0 Groesse 100 %
$script:Window.Add_PreviewKeyDown({
    param($s, $e)
    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    if ("$($e.Key)" -eq 'F1') { $e.Handled = $true; Show-HUMessage 'Die Hilfe kommt in einer der nächsten Testversionen.' -Icon Info; return }
    if ($ctrl -and "$($e.Key)" -in 'D0', 'NumPad0') { $e.Handled = $true; Set-HUUiScale 1.0; return }
})
# Strg + Mausrad = Oberflaeche skalieren
$script:Window.Add_PreviewMouseWheel({
    param($s, $e)
    if (-not ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) { return }
    Set-HUUiScale ($script:UiScale + $(if ($e.Delta -gt 0) { 0.1 } else { -0.1 }))
    $e.Handled = $true
})

$script:Window.Add_ContentRendered({
    if ($script:StartupDone) { return }
    $script:StartupDone = $true
    Close-HUSplash -Animate
    try { [void]$script:Window.Activate() } catch { }
    if ($script:SmokeTestFile) {
        # automatischer Test (CI): Fenster laeuft, Ergebnis schreiben, beenden
        $script:SmokeTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:SmokeTimer.Interval = [TimeSpan]::FromSeconds(1)
        $script:SmokeTimer.Add_Tick({
            $script:SmokeTimer.Stop()
            try {
                $steps = @()
                if ($script:Controls['tabMain'].Items.Count -lt 1) { throw 'keine Reiter' }
                $steps += 'Reiter'
                if ("$($script:Controls['txtAppTitle'].Text)" -ne $script:AppTitle) { throw 'Titel nicht aus app.json' }
                $steps += 'Titel'
                if (-not ('HUTaskbar' -as [type])) { throw 'Taskleiste' }
                $steps += 'Taskleiste'
                Save-HUWindowState
                if (-not (Test-Path -LiteralPath (Join-Path $script:UserRoot 'ui-state.json'))) { throw 'Fensterzustand nicht gespeichert' }
                $steps += 'Fensterzustand'
                "OK Version=$($script:Version) App=$($script:App.id) PS=$($PSVersionTable.PSVersion) Schritte=$($steps -join ',')" | Set-Content -LiteralPath $script:SmokeTestFile -Encoding UTF8
            } catch {
                "FEHLER nach [$($steps -join ',')]: $($_.Exception.Message)" | Set-Content -LiteralPath $script:SmokeTestFile -Encoding UTF8
            }
            $script:Window.Close()
        })
        $script:SmokeTimer.Start()
        return
    }
    # Starter HU-Amt.exe (eigenes Logo, ohne Konsolenfenster) bereitstellen, falls er fehlt
    try { [void](New-HULauncher) } catch { }
    if ($script:FirstRun) {
        # neuer Benutzer: einmal fragen, ob ein Symbol auf den Desktop soll
        $asked = $false; try { $asked = [bool](Get-HUStateValue 'shortcutAsked' $false) } catch { }
        if (-not $asked -and -not (Test-HUShortcut)) {
            try { Set-HUStateValue 'shortcutAsked' $true; Save-HUUIState } catch { }
            if (Confirm-HU "Ein Symbol für $($script:AppTitle) auf den Desktop legen?`n`nDamit startest du die App künftig mit Doppelklick.") {
                try { $l = New-HUShortcut -Location Desktop; Write-HULogOK "Verknüpfung erstellt: $l" } catch { Show-HUMessage "Verknüpfung fehlgeschlagen:`n$($_.Exception.Message)" -Icon Warning }
            }
        }
    }
})

$script:Window.Add_Closing({
    param($s, $e)
    # nach dem Wiederherstellen einer Sicherung nichts aus dem Speicher zurueckschreiben
    if (-not $script:AbRestored) { Save-HUWindowState }
    try { Close-HUToasts } catch { }
    Write-HULog -Message "=== $($script:AppTitle) beendet ===" -Level 'INFO'
})

# ============================================================================
# 5. START
# ============================================================================
Write-HULogInfo "$($script:AppTitle) v$($script:Version) startet (PowerShell $($PSVersionTable.PSVersion), Benutzer $($script:UserName)) ..."
Update-VersionDisplay
Restore-HUWindowState
Select-HUStartTab
Write-HULogOK 'Bereit.'

try { [void]$script:Window.ShowDialog() }
finally { if ($script:InstanceMutex) { try { $script:InstanceMutex.ReleaseMutex() } catch { }; $script:InstanceMutex.Dispose() } }
