#Requires -Version 5.1
<#
.SYNOPSIS
    HU-Amt - Werkzeuge fuer die Gemeindeverwaltung (WPF, ein Reiter je Modul).
.DESCRIPTION
    Startpunkt. Liest Config\app.json (Name, Repo, Signatur), laedt Core-Module (Core\*.psm1),
    Oberflaeche (XAML\*.xaml) und Funktionen (Functions\*.ps1).
      Config\      app.json, version.json (im Repo); lokale Dateien wie update.json nicht im Repo
      Core\        Logging, Datenordner und Sperren (HU.Paths), Sicherung (HU.AppBackup), Datei-Austausch (HU.Exchange), Outlook und Word (HU.Office), Kontakte (HU.Contacts),
                   Datenablage/Zahlen/Verlauf/Anhaenge (HU.Data), PDF (HU.Pdf), Excel/CSV (HU.Import), Werkzeug-Aufrufe (HU.Tools)
      Functions\   Oberflaeche: Dialoge, Fensterzustand, Startbildschirm, Starter und Verknuepfungen, Module, Update-Bibliothek
      Modules\     mitgelieferte Module; Addons\ private (signiert); AddonsDev\ Entwicklung - je Modul ein Reiter
      XAML\        Fenster und gemeinsames Theme
      Data\        wird vom Update nie angefasst; Persoenliches in Data\Benutzer\<Name>\
    Start: HU-Amt.exe (Desktop-Verknuepfung) oder Start-HU-Amt.cmd.
.NOTES
    Zielmaschine: der PC, auf dem HU-Amt laeuft (Windows 10/11, Windows PowerShell 5.1 oder PowerShell 7).
#>
# -Hell: nur fuer das kleine Neustart-Fenster (die Einstellungen sind dort noch nicht geladen), setzt Restart-HUApp
param([switch]$ShowConsole, [string]$SmokeTest = '', [int]$WaitForPid = 0, [switch]$TestMode, [switch]$Demo, [switch]$Hell)

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
$script:SmokeTestFile = $SmokeTest; $script:SmokeTimer = $null; $script:SmokeSteps = @(); $script:SmokeWatch = $null; $script:OnContentRendered = $null; $script:SmokeState = $null; $script:SmokeRun = $null; $script:SmokeRan = $false
$script:Settings = $null; $script:App = $null; $script:Paths = $null; $script:RunLock = $null

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
            $rbg = $(if ($Hell) { '#FFF2EFE9' } else { '#FF1E1E1E' }); $rfg = $(if ($Hell) { '#FF2E2A24' } else { '#FFEDEDED' })
            $rw = [System.Windows.Markup.XamlReader]::Parse("<Window xmlns=`"http://schemas.microsoft.com/winfx/2006/xaml/presentation`" WindowStyle=`"None`" ResizeMode=`"NoResize`" AllowsTransparency=`"True`" Background=`"Transparent`" Width=`"360`" Height=`"90`" WindowStartupLocation=`"CenterScreen`" Topmost=`"True`" ShowInTaskbar=`"False`"><Border Background=`"$rbg`" CornerRadius=`"10`" BorderBrush=`"#FFB9A88A`" BorderThickness=`"2`"><TextBlock Text=`"$tx`" Foreground=`"$rfg`" FontSize=`"15`" HorizontalAlignment=`"Center`" VerticalAlignment=`"Center`"/></Border></Window>")
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
# DATENORDNER (Core\HU.Paths): Gemeinsames in Data\<Modul-Id>\, Persoenliches in Data\Benutzer\<Name>\,
# laufende Apps in Data\Laufend\ - gilt am Einzel-PC und im gemeinsamen Netzordner gleich
# ============================================================================
$pathsMod = Join-Path $script:AppRoot 'Core\HU.Paths.psm1'
if (-not (Test-Path -LiteralPath $pathsMod)) { Show-HUFatal "Core-Modul fehlt: $pathsMod`n`nPull.ps1 ausführen, um die Dateien zu laden." }
Import-Module $pathsMod -Force -DisableNameChecking
try {
    $script:Paths = Initialize-HUDataPaths -AppRoot $script:AppRoot
    Assert-HUWritable $script:Paths.UserRoot
} catch {
    $__loc = $null; try { $__loc = Get-HUDataLocation -AppRoot $script:AppRoot } catch { }
    if ($__loc -and $__loc.Configured) {
        Show-HUFatal "Der Datenordner ist nicht erreichbar oder nicht beschreibbar:`n$($__loc.Path)`n`nNetzwerk bzw. VPN prüfen. Eingestellt ist er in Config\datenort.json (Einstellungen › Verwalter).`n`n$($_.Exception.Message)"
    }
    Show-HUFatal "Der Datenordner ist nicht beschreibbar:`n$(Join-Path $script:AppRoot 'Data')`n`nBitte Schreibrechte auf den Ordner Data geben lassen.`n`n$($_.Exception.Message)"
}
$script:DataRoot = $script:Paths.DataRoot
$script:UserRoot = $script:Paths.UserRoot
$script:UserName = $script:Paths.UserName
# angemeldet als laufend (Regel 2) - nicht im Starttest, damit die CI keine Sperren hinterlaesst
$script:RunLockError = ''
if (-not $SmokeTest) { try { $script:RunLock = Open-HURunLock -RunDir $script:Paths.RunDir } catch { $script:RunLock = $null; $script:RunLockError = "$($_.Exception.Message)" } }

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
foreach ($mod in @('HU.Logging', 'HU.AppBackup', 'HU.Exchange', 'HU.Office', 'HU.Contacts', 'HU.Data', 'HU.Pdf', 'HU.Qr', 'HU.Import', 'HU.Tools')) {
    $modPath = Join-Path $script:AppRoot "Core\$mod.psm1"
    if (-not (Test-Path -LiteralPath $modPath)) { Show-HUFatal "Core-Modul fehlt: $modPath`n`nPull.ps1 ausführen, um die Dateien zu laden." }
    Import-Module $modPath -Force -DisableNameChecking
}
foreach ($f in @('Core-Async', 'Core-Update', 'Core-Addons', 'Core-Outpost', 'Core-TestMode', 'UI-Theme', 'UI-Common', 'UI-Controls', 'UI-Jobs', 'UI-State', 'UI-Shell', 'UI-Splash', 'UI-Update', 'UI-Addons', 'UI-AppBackup', 'UI-Support', 'UI-Settings', 'UI-Outpost', 'UI-Help', 'UI-TestAssistant', 'UI-Mail', 'UI-Contacts')) {
    $fp = Join-Path $script:AppRoot "Functions\$f.ps1"
    if (-not (Test-Path -LiteralPath $fp)) { Show-HUFatal "Datei fehlt: $fp`n`nPull.ps1 ausführen, um die Dateien zu laden." }
    . $fp
}

# ============================================================================
# 1b. TESTMODUS (nur Verwalter, nur diese Sitzung): Datenstamm Data\_Test\, Outlook/Word/Mail gesperrt
# ============================================================================
Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue   # nie vom Vorgaenger-Prozess erben
Remove-Item -LiteralPath Env:\HUAMT_DEMO -ErrorAction SilentlyContinue
# Achtung: $script:TestMode waere der Parameter selbst (Skript-Parameter sind Skriptvariablen), darum TestModeActive
$script:TestModeRequested = [bool]($TestMode -or $Demo); $script:TestModeActive = $false; $script:TestResetDone = $false
$script:DemoActive = $false
# Demo (alle Benutzer): eigener Bestand je Benutzer unter _Demo, bei jedem Start geleert und mit Beispieldaten neu gefuellt
$script:RealUserRoot = $script:UserRoot   # echter Benutzerordner (Demo: Farbschema von dort)
if ($Demo) {
    try {
        $env:HUAMT_DEMO = '1'
        Initialize-HUDemoData
        $script:Paths = Initialize-HUDataPaths -AppRoot $script:AppRoot -Demo
        Assert-HUWritable $script:Paths.UserRoot
        $script:DataRoot = $script:Paths.DataRoot
        $script:UserRoot = $script:Paths.UserRoot
        $script:StateFilePath = Join-Path $script:UserRoot 'ui-state.json'
        $env:HUAMT_TESTMODE = '1'
        $script:TestModeActive = $true; $script:DemoActive = $true
    } catch { Remove-Item -LiteralPath Env:\HUAMT_DEMO -ErrorAction SilentlyContinue; Show-HUFatal "Die Demo konnte nicht starten:`n$($_.Exception.Message)" }
}
# im Starttest (CI) ohne Verwalter-Pruefung, damit auch der Testmodus automatisch geprueft wird
if ($TestMode -and -not $script:DemoActive) {
    if ($SmokeTest -or (Test-HUVerwalter)) {
        try {
            $script:TestResetDone = Initialize-HUTestData
            $script:Paths = Initialize-HUDataPaths -AppRoot $script:AppRoot -TestMode
            Assert-HUWritable $script:Paths.UserRoot
            $script:DataRoot = $script:Paths.DataRoot
            $script:UserRoot = $script:Paths.UserRoot
            # beim Laden der Funktionen schon aus dem echten Benutzerordner gebildet: neu setzen
            $script:StateFilePath = Join-Path $script:UserRoot 'ui-state.json'
            $env:HUAMT_TESTMODE = '1'
            $script:TestModeActive = $true
        } catch { Show-HUFatal "Der Testmodus konnte nicht starten:`n$($_.Exception.Message)" }
    }
}

# ============================================================================
# 2. EINSTELLUNGEN (je Benutzer, ohne Datei gelten die Vorgaben)
# ============================================================================
Set-HUSplashStatus 'Einstellungen ...'
$script:SettingsPath = Join-Path $script:UserRoot 'settings.json'
$script:FirstRun = -not (Test-Path -LiteralPath (Join-Path $script:UserRoot 'ui-state.json'))
$script:Settings = [pscustomobject]@{
    ui      = [pscustomobject]@{ startTab = 'Last'; splashAnimation = $true; maxLogLines = 1000; checkUpdatesOnStart = $true; theme = 'Dunkel' }
    logging = [pscustomobject]@{ logLevel = 'INFO'; retentionDays = 30 }
}
# Demo: der Bestand ist leer, das eigene Farbschema gilt trotzdem
if ($script:DemoActive) { try { $real = Read-HUJsonFile (Join-Path $script:RealUserRoot 'settings.json'); if ($real -and $real.ui -and $real.ui.PSObject.Properties['theme']) { $script:Settings.ui.theme = "$($real.ui.theme)" } } catch { } }
$saved = Read-HUJsonFile $script:SettingsPath
if ($saved) {
    foreach ($sec in 'ui', 'logging') {
        if (-not $saved.PSObject.Properties[$sec]) { continue }
        foreach ($p in $saved.$sec.PSObject.Properties) {
            if ($script:Settings.$sec.PSObject.Properties[$p.Name]) { $script:Settings.$sec.($p.Name) = $p.Value }
        }
    }
}

# Art der Installation (Zentrale / Aussenstelle) und abgeschaltete Module: Config\installation.json (Functions\Core-Outpost.ps1)
$script:Installation = Get-HUInstallationConfig $script:ConfigDir

# Farbschema (persoenlich, wirkt ab dem Start): Hell setzt die dunklen Farben beim Laden um (Functions\UI-Theme.ps1)
Set-HUTheme "$($script:Settings.ui.theme)"
if ($SmokeTest -and "$env:HUAMT_SMOKE_THEME" -eq 'Hell') { Set-HUTheme 'Hell' }   # Starttest im hellen Schema (CI)
if (Test-HULightTheme) { try { Set-HULogColors @{ OK = '#2E7D32'; WARN = '#C25E00'; ERROR = '#C62828'; INFO = '#1565C0'; DEBUG = '#6E675C' } '#2E2A24' } catch { } }

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
$c['txtStartTitle'].Text = "Willkommen bei $($script:AppTitle)"

$script:LogDir = Join-Path $script:UserRoot 'Logs'
$logFile = Join-Path $script:LogDir "$($script:App.id)_{date}.log"
$minLevel = "$($script:Settings.logging.logLevel)".ToUpper()
if ($minLevel -notin 'DEBUG', 'INFO', 'OK', 'WARN', 'ERROR') { $minLevel = 'INFO' }
Initialize-Logging -LogFilePath $logFile -RichTextBox $c['rtbLog'] -MinLevel $minLevel -MaxGuiLines ([int]$script:Settings.ui.maxLogLines)
Write-HULogInfo "$($script:AppTitle) v$($script:Version) startet (PowerShell $($PSVersionTable.PSVersion), Benutzer $($script:UserName)) ..."
# ohne Anmeldung in Data\Laufend sehen Update und Datenordner-Wechsel diesen PC nicht (Pruefung 090, KERN-U1)
if ($script:RunLockError) { Write-HULogWarn "Anmeldung als laufende App fehlgeschlagen, Update und Datenordner-Wechsel sehen diesen PC nicht: $($script:RunLockError)" }
if ($script:DemoActive) {
    Write-HULogWarn "DEMO: Beispieldaten in $($script:DataRoot). Outlook, Word und Mails sind gesperrt, die automatische Sicherung ruht."
    $c['bdrTestMode'].Visibility = 'Visible'
    $c['bdrTestMode'].Background = Get-HUBrush '#1565C0'
    $c['bdrTestMode'].ToolTip = 'Demo: eigene Beispieldaten nur für dich, echte Daten bleiben unberührt'
    $c['txtTestMode'].Text = "$([char]::ConvertFromUtf32(0x1F3AC)) DEMO – Beispieldaten zum Ausprobieren, nichts davon ist echt. Kein Outlook, kein Word, keine Mails."
    $c['txtTestMode'].ToolTip = 'Beim nächsten Demo-Start ist alles wieder wie neu. Die automatische Sicherung ruht.'
    $c['btnTestReset'].Visibility = 'Collapsed'
    $c['btnTestEnd'].Content = 'Demo beenden'
    $c['btnTestEnd'].ToolTip = 'Neu starten und wieder mit den echten Daten arbeiten'
    $c['btnTestEnd'].Add_Click({ Switch-HUDemoMode })
} elseif ($script:TestModeActive) {
    Write-HULogWarn "TESTMODUS: Daten in $($script:DataRoot). Outlook, Word und Mails sind gesperrt, die automatische Sicherung ruht."
    if ($script:TestResetDone) { Write-HULogOK 'Testbestand wurde zurückgesetzt.' }
    $c['bdrTestMode'].Visibility = 'Visible'
    $c['btnTestReset'].Add_Click({ Request-HUTestReset })
    $c['btnTestEnd'].Add_Click({ Switch-HUTestMode })
}
# Protokolle aelter als logging.retentionDays (Standard 30) beim Start loeschen: die der App und die der Module (alle *.log in Logs)
try {
    $keep = 30; try { if ([int]$script:Settings.logging.retentionDays -gt 0) { $keep = [int]$script:Settings.logging.retentionDays } } catch { }
    $limit = (Get-Date).AddDays(-$keep)
    Get-ChildItem -LiteralPath $script:LogDir -Filter '*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $limit } | Remove-Item -Force -ErrorAction SilentlyContinue
} catch { }

Initialize-AsyncPool -PoolSize 4

# Kopfzeile: Version und Update-Kanal. Stabil = gruen, Test = orange, Entwicklungsstand (Branch) = rot.
function Update-VersionDisplay {
    $cfg = $null; try { $cfg = Update-UpdateConfig } catch { }
    $txt = 'Kanal Stabil'; $col = '#4CAF50'
    if ($cfg -and $cfg.UseBranch) { $txt = 'Entwicklungsstand'; $col = '#EF5350' }
    elseif ($cfg -and "$($cfg.Channel)" -eq 'Test') { $txt = 'Kanal Test'; $col = '#FFA726' }
    $t = $script:Controls['txtVersion']
    $t.Text = "v$($script:Version)  |  $txt"
    $t.Foreground = Get-HUBrush $col 'Fg'
    $t.ToolTip = $(if ($cfg -and $cfg.UseBranch) { "Updates kommen direkt aus dem Repo (Branch $($cfg.Branch)), ohne freigegebene Version" } elseif ($txt -eq 'Kanal Test') { 'Updates: auch Testversionen.' } else { 'Updates: nur freigegebene Versionen.' })
    $ost = ''; try { $ost = Get-HUOutpostLabel $script:Installation } catch { }
    if ($ost) { $t.Text = "$($t.Text)  |  $ost" }
    $script:Window.Title = "$(if ($script:DemoActive) { 'DEMO - ' } elseif ($script:TestModeActive) { 'TEST - ' })$($script:AppTitle)$(if ($ost) { " - $ost" }) v$($script:Version)"
}

# ============================================================================
# 4. EREIGNISSE
# ============================================================================
$c['btnSettings'].Add_Click({ Open-HUSettings })
$c['btnSupport'].Add_Click({ Show-HUSupport })
Register-HUAbMainStatus
Register-HUHelpHandlers
Register-HUUpdateHandlers
$c['btnOpenLogFolder'].Add_Click({ Open-HUPath $script:LogDir })
# Protokoll unten: Hoehe per Ziehgriff (gemerkt), nie hoeher als das Fenster minus Platz fuer die Reiter
try { $h = [double](Get-HUStateValue 'logHeight' 170); if ($h -ge 80 -and $h -le 900) { $c['pnlLog'].Height = $h } } catch { }
$c['thbLog'].Add_DragDelta({
        param($s, $e)
        $max = [Math]::Max(80.0, $script:Window.ActualHeight - 260)
        $h = $script:Controls['pnlLog'].Height - $e.VerticalChange
        $script:Controls['pnlLog'].Height = [Math]::Max(80.0, [Math]::Min($max, $h))
    })
$c['thbLog'].Add_DragCompleted({ Set-HUStateValue 'logHeight' ([int]$script:Controls['pnlLog'].Height); Save-HUUIState })
try { Set-HUOutputTools $script:Window } catch { }

# ============================================================================
# MODULE: je Modul ein Reiter (Funktionen per Dot-Source in den Programm-Bereich, dann Reiter einbauen und starten)
# ============================================================================
Set-HUSplashStatus 'Module werden eingebaut ...'
foreach ($__a in @(Get-HUAddonStartList)) {
    try {
        # Testmodus: Beispieldaten des Moduls (Testdaten\) in den leeren Testbestand legen, bevor das Modul liest
        if ($script:TestModeActive) {
            try { $__n = Initialize-HUAddonTestData $__a; if ($__n) { Write-HULogInfo "$(if ($script:DemoActive) { 'Demo' } else { 'Testmodus' }): $__n Beispieldateien für $($__a.Title) angelegt" } }
            catch { Write-HULogWarn "Testmodus: Beispieldaten für $($__a.Id) nicht angelegt: $($_.Exception.Message)" }
        }
        . (Join-Path $__a.Dir $__a.Script)
        Add-HUAddonTab $__a
        & $__a.Init
        # verschiebbare Bereiche des Reiters: Kern merkt sie sich (nach dem Start des Moduls, damit die gemerkte Groesse gilt)
        try { Register-HUSplitterState (@($script:Controls['tabMain'].Items)[-1]) "Reiter.$($__a.Id)" } catch { Write-HULogWarn "Bereiche $($__a.Id): $($_.Exception.Message)" }
        $script:LoadedAddons += $__a
        Write-HULogOK "Modul geladen: $($__a.Title) $($__a.Version)$(if ($__a.Source -ne 'Programm') { " ($($__a.Source))" })"
    } catch { Write-HULogWarn "Modul $($__a.Id) nicht geladen: $($_.Exception.Message)" }
}
Remove-Variable __a -ErrorAction SilentlyContinue
# Dateien ins Fenster ziehen: gehen an das Modul im aktiven Reiter
Register-HUFileDrop
# Postfach zwischen Modulen: Zahl wartender Dateien am Reiter
if (-not $SmokeTest) { Start-HUAddonInboxWatch }

# Tastatur: F1 Hilfe, Strg+0 Groesse 100 %
$script:Window.Add_PreviewKeyDown({
    param($s, $e)
    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    if ("$($e.Key)" -eq 'F1') { $e.Handled = $true; Show-HUHelp; return }
    if ($ctrl -and "$($e.Key)" -in 'D0', 'NumPad0') { $e.Handled = $true; Set-HUUiScale 1.0; return }
})
# Strg + Mausrad = Oberflaeche skalieren
$script:Window.Add_PreviewMouseWheel({
    param($s, $e)
    if (-not ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) { return }
    Set-HUUiScale ($script:UiScale + $(if ($e.Delta -gt 0) { 0.1 } else { -0.1 }))
    $e.Handled = $true
})

$script:OnContentRendered = {
    if ($script:StartupDone) { return }
    $script:StartupDone = $true
    Close-HUSplash -Animate
    try { [void]$script:Window.Activate() } catch { }
    if ($script:SmokeTestFile) {
        # automatischer Test (CI): Fenster laeuft, Ergebnis schreiben, beenden
        $script:SmokeTimer = [System.Windows.Threading.DispatcherTimer]::new([System.Windows.Threading.DispatcherPriority]::Normal)
        $script:SmokeTimer.Interval = [TimeSpan]::FromSeconds(1)
        $script:SmokeTimer.Add_Tick({ $script:SmokeTimer.Stop(); $script:SmokeState.Steps = '(Zeitgeber lief)'; & $script:SmokeRun })
        $script:SmokeRun = {
            if ($script:SmokeRan) { return }
            $script:SmokeRan = $true
            try {
                $script:SmokeSteps = @()
                if ($script:Controls['tabMain'].Items.Count -lt 1) { throw 'keine Reiter' }
                $script:SmokeSteps += 'Reiter'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                if (Test-HULightTheme) { $script:SmokeSteps += 'Hell'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',') }
                # Bilder fuer die Sichtpruefung (nur mit HUAMT_SMOKE_SHOTS = Ordner): Hauptfenster mit jedem Reiter
                if ($env:HUAMT_SMOKE_SHOTS) {
                    $selShot = $script:Controls['tabMain'].SelectedItem
                    foreach ($it in @($script:Controls['tabMain'].Items)) {
                        $script:Controls['tabMain'].SelectedItem = $it
                        Save-HUWindowShot $script:Window (Join-Path $env:HUAMT_SMOKE_SHOTS "Reiter-$($it.Header).png")
                    }
                    $script:Controls['tabMain'].SelectedItem = $selShot
                }
                if ("$($script:Controls['txtAppTitle'].Text)" -ne $script:AppTitle) { throw 'Titel nicht aus app.json' }
                $script:SmokeSteps += 'Titel'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                if (-not ('HUTaskbar' -as [type])) { throw 'Taskleiste' }
                $script:SmokeSteps += 'Taskleiste'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                if ($script:TestModeRequested) {
                    if (-not $script:TestModeActive -or -not (Test-HUTestMode)) { throw 'Testmodus nicht aktiv' }
                    $sep = [IO.Path]::DirectorySeparatorChar
                    if ($script:DemoActive) {
                        if (-not (Test-HUDemoMode)) { throw 'Demo nicht aktiv' }
                        if ($script:DataRoot -notlike "*Data$($sep)_Demo$($sep)*") { throw "Demo: Datenstamm $($script:DataRoot)" }
                        if ("$($script:Controls['btnTestReset'].Visibility)" -ne 'Collapsed') { throw 'Demo: Zurücksetzen sichtbar' }
                    } elseif (-not $script:DataRoot.EndsWith('Data' + $sep + '_Test')) { throw "Testmodus: Datenstamm $($script:DataRoot)" }
                    if ("$($script:Controls['bdrTestMode'].Visibility)" -ne 'Visible') { throw 'Testmodus: Leiste nicht sichtbar' }
                    if (-not (Test-HUOfficeTestMode)) { throw 'Testmodus: Office nicht gesperrt' }
                    if ((Show-HUMail -Subject 'x').Way -ne 'testmodus') { throw 'Testmodus: Mail nicht gesperrt' }
                    $script:SmokeSteps += $(if ($script:DemoActive) { 'Demo' } else { 'Testmodus' }); $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                }
                $tabs = @($script:Controls['tabMain'].Items).Count
                if ($tabs -ne [Math]::Max(1, @($script:LoadedAddons).Count)) { throw "Reiter: $tabs, Module: $(@($script:LoadedAddons).Count)" }
                $script:SmokeSteps += "Module=$(@($script:LoadedAddons).Count)"; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                # Hineinziehen: bei offenem Reiter eines Moduls muss jede seiner Endungen dort ankommen (Ziel ohne echtes Ziehen bestimmen).
                # Teilen sich mehrere Module eine Endung (z. B. .pdf), entscheidet in der App der offene Reiter.
                $selVorher = $script:Controls['tabMain'].SelectedItem
                foreach ($a in @($script:LoadedAddons | Where-Object { Test-HUAddonCanDrop $_ })) {
                    $tab = $null
                    foreach ($it in @($script:Controls['tabMain'].Items)) { if ($script:AddonByTab.ContainsKey("$($it.Name)") -and $script:AddonByTab["$($it.Name)"].Id -eq $a.Id) { $tab = $it; break } }
                    if ($tab) { $script:Controls['tabMain'].SelectedItem = $tab }
                    foreach ($x in @($a.Drop)) {
                        $tg = Resolve-HUDropTarget @("C:\Test\Datei$x")
                        if (-not $tg -or $tg.Info.Id -ne $a.Id) { throw "Hineinziehen: $x kommt bei offenem Reiter nicht bei $($a.Id) an" }
                    }
                    $script:SmokeSteps += "Ziehen-$($a.Id)"; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                }
                if ($selVorher) { $script:Controls['tabMain'].SelectedItem = $selVorher }
                if (Resolve-HUDropTarget @('C:\Test\Datei.unbekannt')) { throw 'Hineinziehen: unbekannte Endung angenommen' }
                # Sicherung: Quellen nur aus Data, ohne Logs, Token, Sperren
                $src = @(Get-HUAppBackupSources -DataRoot $script:DataRoot)
                if (@($src | Where-Object { $_.Name -match '/(Logs|Token)/|^data/(Laufend|Docs)/|\.lock$' }).Count) { throw 'Sicherung: verbotene Dateien in der Liste' }
                $script:SmokeSteps += "Sicherung=$($src.Count)"; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                # Support: Info ohne Benutzername, ZIP entsteht
                if ((ConvertTo-HURedacted ((Get-HUSupportInfo) -join "`n") (Get-HURedactMap)) -match [regex]::Escape($env:USERNAME)) { throw 'Support: Benutzername in info.txt' }
                $zf = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtSmoke-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
                $z = New-HUSupportZip -Include @{ Logs = $true; Settings = $true } -Folder $zf
                if (-not (Test-Path -LiteralPath $z)) { throw 'Support-ZIP fehlt' }
                Remove-Item -LiteralPath $zf -Recurse -Force
                $script:SmokeSteps += 'Support'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                # Fenster wirklich aufbauen: ein zweiter Timer schliesst das jeweils offene Fenster nach kurzer Zeit
                $closer = [System.Windows.Threading.DispatcherTimer]::new([System.Windows.Threading.DispatcherPriority]::Normal)
                $closer.Interval = [TimeSpan]::FromMilliseconds(1200)
                $closer.Add_Tick({
                        if ($script:LastDialog -and $script:LastDialog.IsVisible) {
                            if ($env:HUAMT_SMOKE_SHOTS) { Save-HUWindowShot $script:LastDialog (Join-Path $env:HUAMT_SMOKE_SHOTS "Fenster-$($script:LastDialog.Title).png") }
                            $script:LastDialog.Close()
                        }
                    })
                $closer.Start()
                try {
                    Show-HUSupport; $script:SmokeSteps += 'Fenster-Support'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                    Open-HUSettings -Tab 'Sicherung'; $script:SmokeSteps += 'Fenster-Sicherung'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                    [void](Show-HUAbPasswordDialog -New); $script:SmokeSteps += 'Fenster-Kennwort'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                    Show-HUAbout; $script:SmokeSteps += 'Fenster-Info'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                    Open-HUSettings; $script:SmokeSteps += 'Fenster-Einstellungen'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                    Show-HUAddonDialog; $script:SmokeSteps += 'Fenster-Module'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                    if (-not (Get-HUHelpPage $null)) { throw 'Grundhilfe fehlt' }
                    $script:SmokeSteps += 'Hilfe'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                } finally { $closer.Stop() }
                Save-HUWindowState
                if (-not (Test-Path -LiteralPath (Join-Path $script:UserRoot 'ui-state.json'))) { throw 'Fensterzustand nicht gespeichert' }
                $script:SmokeSteps += 'Fensterzustand'; $script:SmokeState.Steps = ($script:SmokeSteps -join ',')
                $how = if ($script:SmokeState.Kicked) { " Start=Dispatcher(Zeitgeber kam nicht, Thread=$($script:SmokeState.ThreadState))" } else { '' }
                "OK Version=$($script:Version) App=$($script:App.id) PS=$($PSVersionTable.PSVersion) Schritte=$($script:SmokeSteps -join ',')$how" | Set-Content -LiteralPath $script:SmokeTestFile -Encoding UTF8
            } catch {
                "FEHLER nach [$($script:SmokeSteps -join ',')]: $($_.Exception.Message)" | Set-Content -LiteralPath $script:SmokeTestFile -Encoding UTF8
            }
            $script:Window.Close()
        }
        $script:SmokeTimer.Start()
        $script:SmokeState.Steps = '(Zeitgeber gestartet)'
        return
    }
    # Erststart ohne Verwalter: einmal fragen, ob dieser Benutzer Verwalter wird
    if (-not $script:TestModeActive) { try { [void](Request-HUFirstVerwalter) } catch { Write-HULogWarn "Verwalter: $($_.Exception.Message)" } }
    # Erster Start aus einem Paket fuer eine Aussenstelle: Programm pruefen, beigelegten Zugang fuer Modul-Updates einlesen
    if (-not $script:TestModeActive) { try { Invoke-HUPackageFirstStart } catch { Write-HULogWarn "Paket: $($_.Exception.Message)" } }
    # Rechte der Verwalter-Liste: nur ins Protokoll (keine Meldung fuer Benutzer); Details sieht der Verwalter in den Einstellungen
    if (-not $script:TestModeActive) {
        # im Hintergrund mit Zeitlimit, damit eine langsame Freigabe den Start nicht blockiert
        try {
            Start-HUBackgroundJob -Owner 'Start' -LogPrefix 'Rechte' -TimeoutSec 30 -Arguments @($script:AppRoot, $script:DataRoot) -Script {
                param($AppRoot, $DataRoot)
                Import-Module (Join-Path $AppRoot 'Core\HU.Paths.psm1') -DisableNameChecking -ErrorAction Stop
                Test-HUVerwalterAcl -DataRoot $DataRoot -AppRoot $AppRoot
            } -Done { param($Res, $Err, $Ctx) if ($Res -and $Res.Status -eq 'Warnung') { Write-HULogWarn "Rechte: $($Res.Text) $(@($Res.Funde) -join '; ')" } }
        } catch { }
    }
    # Sicherung: Stand in der Kopfzeile, Automatik (beim Start und stuendlich, wenn faellig) - nicht im Testmodus
    if (-not $script:TestModeActive) { try { Update-HUAbMainStatus; Start-HUAppBackupScheduler } catch { Write-HULogWarn "Sicherung: $($_.Exception.Message)" } }
    # Update-Pruefung im Hintergrund (Knopf wird gold), Module im Hintergrund aktualisieren
    if ($script:Settings.ui.checkUpdatesOnStart -ne $false) { try { Invoke-UpdateCheck -Quiet; Start-HUAddonUpdate } catch { Write-HULogWarn "Update-Prüfung: $($_.Exception.Message)" } }
    # Starter HU-Amt.exe (eigenes Logo, ohne Konsolenfenster) bereitstellen, falls er fehlt
    try { [void](New-HULauncher) } catch { }
    if ($script:FirstRun -and -not $script:DemoActive) {   # die Demo ist bei jedem Start neu: dort nicht jedes Mal fragen
        # neuer Benutzer: einmal fragen, ob ein Symbol auf den Desktop soll, und zeigen, wie die App an die Taskleiste kommt
        $asked = $false; try { $asked = [bool](Get-HUStateValue 'shortcutAsked' $false) } catch { }
        if (-not $asked) {
            try { Set-HUStateValue 'shortcutAsked' $true; Save-HUUIState } catch { }
            if (-not (Test-HUShortcut) -and (Confirm-HU "Ein Symbol für $($script:AppTitle) auf den Desktop legen?`n`nDamit startest du die App künftig mit Doppelklick.")) {
                try { $l = New-HUShortcut -Location Desktop; Write-HULogOK "Verknüpfung erstellt: $l" } catch { Show-HUMessage "Verknüpfung fehlgeschlagen:`n$($_.Exception.Message)" -Icon Warning }
            }
            $pinned = $true; try { $pinned = Test-HUTaskbarPinned } catch { }
            if (-not $pinned -and (Confirm-HU "$($script:AppTitle) auch an die Taskleiste anheften?`n`nWindows lässt das Anheften nur durch dich selbst zu. Bei Ja zeige ich dir, wie es geht.")) {
                Show-HUMessage "So heftest du $($script:AppTitle) an:`n`n1. Unten in der Taskleiste das Symbol von $($script:AppTitle) mit der rechten Maustaste anklicken.`n2. $([char]0x201E)An Taskleiste anheften$([char]0x201C) wählen.`n`nDanach startest du die App mit einem Klick auf das Symbol." -Icon Info -Title 'Taskleiste'
            }
        }
    }
}
$script:Window.Add_ContentRendered($script:OnContentRendered)

$script:Window.Add_Closing({
    param($s, $e)
    # Absicherung zuerst (Pruefung 091, KERN-B1): Endet der Prozess nicht innerhalb von 60 s (beobachtet 08.11.2026,
    # Ursache unklar; auch ein haengender Schritt hier im Schliessen), beendet ihn ein eigener Thread hart und schreibt
    # vorher den zuletzt erreichten Schritt ins Protokoll. Normales Beenden verzoegert er nicht.
    $script:CloseState = [hashtable]::Synchronized(@{ Step = 'Start' })
    if (-not $SmokeTest) {
        try {
            $__lf = ''; try { $__lf = "$(Get-CurrentLogFile)" } catch { }
            $__xr = [runspacefactory]::CreateRunspace(); $__xr.ThreadOptions = 'UseNewThread'; $__xr.Open()
            $__xp = [powershell]::Create(); $__xp.Runspace = $__xr
            [void]$__xp.AddScript({
                    param($St, $Log)
                    Start-Sleep -Seconds 60   # Platz fuer laufende Abgleiche der Module (Buchung wartet bis 20 s auf Outlook)
                    if ($Log) { try { [IO.File]::AppendAllText($Log, "[$((Get-Date).ToString('yyyy-MM-dd'))] [$((Get-Date).ToString('HH:mm:ss'))] [WARN] Beenden hart abgeschlossen nach 60 s, letzter Schritt: $($St.Step)`r`n", (New-Object System.Text.UTF8Encoding $true)) } catch { } }
                    [Environment]::Exit(0)
                }).AddArgument($script:CloseState).AddArgument($__lf).BeginInvoke()
        } catch { }
    }
    # nach dem Wiederherstellen einer Sicherung nichts aus dem Speicher zurueckschreiben
    $script:CloseState.Step = 'Fensterzustand'; if (-not $script:AbRestored) { Save-HUWindowState }
    $script:CloseState.Step = 'Hinweise'; try { Close-HUToasts } catch { }
    try { if ($script:AbTimer) { $script:AbTimer.Stop() } } catch { }
    $script:CloseState.Step = 'Module beenden'; try { Stop-HUAddons } catch { }
    $script:CloseState.Step = 'Hintergrund-Aufträge'; try { Close-AsyncPool } catch { }
    $script:CloseState.Step = 'Sperren freigeben'
    Write-HULog -Message "=== $($script:AppTitle) beendet ===" -Level 'INFO'
    # Sperre und Einzelinstanz sofort freigeben (nicht erst nach ShowDialog): Bleibt der Prozess danach haengen,
    # blockiert er weder den Neustart ("bereits geoeffnet") noch das Update ("jemand hat die App offen").
    try { Close-HURunLock $script:RunLock; $script:RunLock = $null } catch { }
    if ($script:InstanceMutex) { try { $script:InstanceMutex.ReleaseMutex() } catch { }; try { $script:InstanceMutex.Dispose() } catch { }; $script:InstanceMutex = $null }
    $script:CloseState.Step = 'Fenster schließen (alles freigegeben)'
})

# ============================================================================
# 5. START
# ============================================================================
Update-VersionDisplay
# Verwalter-Ansicht: Protokoll und Verwaltung nur fuer Verwalter
Update-HUAdminView
Write-HULogInfo "Ansicht: $(if (Test-HUVerwalter) { 'Verwalter' } else { 'Bedienung' })"
Restore-HUWindowState
Select-HUStartTab
try {
    $others = @(Get-HURunningInstances -RunDir $script:Paths.RunDir -Self $script:RunLock)
    if ($others.Count) { Write-HULogInfo "Ebenfalls geöffnet: $(($others | ForEach-Object { "$($_.UserName) ($($_.ComputerName))" }) -join ', ')" }
} catch { }
if (-not $SmokeTest -and -not $script:RunLock) { Write-HULogWarn 'Anmeldung in Data\Laufend nicht möglich - ein Update erkennt diese Sitzung nicht.' }
Write-HULogOK 'Bereit.'

if ($SmokeTest) {
    # Starttest ohne sichtbaren Bildschirm (CI): kommt ContentRendered nicht, nach 10 s trotzdem starten.
    # Zeitgeber mit Prioritaet Normal: ohne Bildschirm kommen Zeitgeber mit der Vorgabe Background u. U. nie dran.
    $script:SmokeState = [hashtable]::Synchronized(@{ Loaded = $false; Rendered = $false; Fallback = $false; Steps = ''; File = $script:SmokeTestFile; Log = "$(Get-CurrentLogFile)"; Kicked = $false; ThreadState = '' })
    # Ausweg, falls Zeitgeber ohne Bildschirm nie feuern: der Waechter-Thread stoesst den Test direkt ueber den Dispatcher an
    $script:SmokeState.Disp = $script:Window.Dispatcher
    $script:SmokeState.Kick = [Action]{ if (-not $script:SmokeRan) { $script:SmokeState.Kicked = $true; if ($script:SmokeTimer) { $script:SmokeTimer.Stop() }; & $script:SmokeRun } }
    $script:Window.Add_Loaded({ $script:SmokeState.Loaded = $true })
    $script:Window.Add_ContentRendered({ $script:SmokeState.Rendered = $true })
    $fb = [System.Windows.Threading.DispatcherTimer]::new([System.Windows.Threading.DispatcherPriority]::Normal)
    $fb.Interval = [TimeSpan]::FromSeconds(10)
    $fb.Add_Tick({
            $this.Stop()
            if (-not $script:StartupDone) { $script:SmokeState.Fallback = $true; Write-HULogWarn "Starttest: ContentRendered kam nicht (Loaded=$($script:SmokeState.Loaded)) - Start über Zeitgeber."; & $script:OnContentRendered }
        })
    $fb.Start()
    # zweiter Waechter in eigenem Thread: greift auch, wenn die Oberflaeche keine Zeitgeber mehr ausfuehrt
    $wr = [runspacefactory]::CreateRunspace(); $wr.Open()
    $wp = [powershell]::Create(); $wp.Runspace = $wr
    [void]$wp.AddScript({
            param($St)
            Start-Sleep -Seconds 40
            if ($St.Steps -like '(Zeitgeber gestartet)*' -and -not (Test-Path -LiteralPath $St.File)) {
                # Testzeitgeber hat nach 40 s nicht gefeuert: Zustand des UI-Threads merken, Test ueber den Dispatcher anstossen
                try { $St.ThreadState = "$($St.Disp.Thread.ThreadState)" } catch { $St.ThreadState = '?' }
                try { [void]$St.Disp.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Normal, $St.Kick) } catch { $St.ThreadState += " Anstoss: $($_.Exception.Message)" }
            }
            Start-Sleep -Seconds 160
            if (-not (Test-Path -LiteralPath $St.File)) {
                # letzte Protokollzeilen dazu: zeigen, was die Oberflaeche zuletzt getan hat
                $tail = ''; try { if ($St.Log -and (Test-Path -LiteralPath $St.Log)) { $tail = (Get-Content -LiteralPath $St.Log -Tail 15 -Encoding UTF8) -join ' | ' } } catch { }
                "FEHLER Oberfläche blockiert (kein Zeitgeber lief): Loaded=$($St.Loaded) ContentRendered=$($St.Rendered) Ausweichstart=$($St.Fallback) Anstoss=$($St.Kicked) Thread=$($St.ThreadState) fertig bis: [$($St.Steps)] Protokoll: $tail" | Set-Content -LiteralPath $St.File -Encoding UTF8
                [Environment]::Exit(3)
            }
        }).AddArgument($script:SmokeState)
    [void]$wp.BeginInvoke()
    # Waechter: haengt der Test (z. B. an einem offenen Fenster), Ergebnis mit dem letzten Schritt schreiben und beenden
    $script:SmokeWatch = [System.Windows.Threading.DispatcherTimer]::new([System.Windows.Threading.DispatcherPriority]::Normal)
    $script:SmokeWatch.Interval = [TimeSpan]::FromSeconds(150)
    $script:SmokeWatch.Add_Tick({
            $this.Stop()
            $open = ''; try { if ($script:LastDialog -and $script:LastDialog.IsVisible) { $open = " offenes Fenster: $($script:LastDialog.Title)" } } catch { }
            "FEHLER Zeitüberschreitung nach [$($script:SmokeSteps -join ',')] Start=$($script:StartupDone) Anstoss=$($script:SmokeState.Kicked) Thread=$($script:SmokeState.ThreadState) Stand=$($script:SmokeState.Steps)$open" | Set-Content -LiteralPath $script:SmokeTestFile -Encoding UTF8
            [Environment]::Exit(2)
        })
    $script:SmokeWatch.Start()
}

if ($SmokeTest -and $env:HUAMT_SMOKE_SHOTS) {
    # Bilder ohne sichtbares Fenster: ausserhalb des Bildschirms, feste Groesse (Dialoge oeffnen mittig ueber dem Hauptfenster)
    $script:Window.WindowState = 'Normal'; $script:Window.WindowStartupLocation = 'Manual'
    $script:Window.Left = -32000; $script:Window.Top = -32000; $script:Window.Width = 1400; $script:Window.Height = 900
}
try { [void]$script:Window.ShowDialog() }
finally {
    Close-HURunLock $script:RunLock
    if ($script:InstanceMutex) { try { $script:InstanceMutex.ReleaseMutex() } catch { }; $script:InstanceMutex.Dispose() }
}
