#Requires -Version 5.1
# Pester 5: Werkzeug-Aufrufe und Dateinamen (Core\HU.Tools.psm1) sowie die Vorlagen fuer Modul-Repos (.github\modul-vorlage).
# Nur kuenstliche Daten in Testordnern unter TEMP. Prozesse nur mit Windows-Bordmitteln (cmd.exe, findstr.exe, ping.exe), ohne Fenster.
# Die Vorlagen laufen nur gegen eine kuenstliche Test-App unter TEMP, nie gegen eine echte Installation.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:ModPath = Join-Path $script:Root 'Core\HU.Tools.psm1'
    $script:Vorlage = Join-Path $script:Root '.github\modul-vorlage'
    Import-Module $script:ModPath -Force -DisableNameChecking
    $script:Cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
    $script:Utf8Bom = New-Object System.Text.UTF8Encoding $true
    $script:Utf8 = New-Object System.Text.UTF8Encoding $false

    # Zerlegt eine Befehlszeile wie die C-Laufzeit von Windows (shell32 CommandLineToArgvW)
    if (-not ('HUToolsTest.Argv' -as [type])) {
        Add-Type -Namespace HUToolsTest -Name Argv -MemberDefinition @'
[DllImport("shell32.dll", SetLastError = true)]
static extern IntPtr CommandLineToArgvW([MarshalAs(UnmanagedType.LPWStr)] string lpCmdLine, out int pNumArgs);
[DllImport("kernel32.dll")]
static extern IntPtr LocalFree(IntPtr hMem);
public static string[] Split(string commandLine) {
    int n;
    IntPtr p = CommandLineToArgvW(commandLine, out n);
    if (p == IntPtr.Zero) { throw new System.ComponentModel.Win32Exception(); }
    try {
        string[] r = new string[n];
        for (int i = 0; i < n; i++) { r[i] = Marshal.PtrToStringUni(Marshal.ReadIntPtr(p, i * IntPtr.Size)); }
        return r;
    } finally { LocalFree(p); }
}
'@
    }

    function New-TestDir([string]$Prefix = 'HUAmtTools') {
        $d = Join-Path ([IO.Path]::GetTempPath()) ($Prefix + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $d)
        return $d
    }
    function Write-TestText([string]$Path, [string]$Text, $Encoding) {
        $dir = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        [IO.File]::WriteAllText($Path, $Text, $Encoding)
    }
    function Test-Bom([string]$Path) {
        $b = [IO.File]::ReadAllBytes($Path)
        return ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
    }
    # kuenstliche Test-App: Main.ps1 schreibt wie der echte Starttest "OK ... Module=<n> Schritte=Ziehen-<id>"
    function New-FakeApp([string]$Dir, [string]$Version = '0.1.5') {
        $main = @'
param([switch]$ShowConsole, [string]$SmokeTest = '', [int]$WaitForPid = 0, [switch]$TestMode)
$dev = Join-Path $PSScriptRoot 'AddonsDev'
$mods = @(Get-ChildItem -LiteralPath $dev -Directory -ErrorAction SilentlyContinue | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'addon.json') })
$ids = @($mods | ForEach-Object { 'Ziehen-' + (([IO.File]::ReadAllText((Join-Path $_.FullName 'addon.json'))) | ConvertFrom-Json).id })
$data = Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Data')
$txt = "OK Version=0 PS=$($PSVersionTable.PSVersion.Major) Module=$($mods.Count) Data=$data Schritte=$($ids -join ',')"
if ($SmokeTest) { $txt | Set-Content -LiteralPath $SmokeTest -Encoding UTF8 }
'@
        Write-TestText (Join-Path $Dir 'Main.ps1') $main $script:Utf8Bom
        Write-TestText (Join-Path $Dir 'Config\version.json') ('{ "version": "' + $Version + '" }') $script:Utf8Bom
        Write-TestText (Join-Path $Dir 'Data\daten.json') '{ "kuenstlich": true }' $script:Utf8Bom
        Write-TestText (Join-Path $Dir 'AddonsDev\Anderes\addon.json') '{ "id": "Anderes" }' $script:Utf8Bom
    }
    # kuenstliches Modul-Repo mit den Vorlagen in .github\dev (wie ein Modul sie uebernimmt)
    function New-FakeRepo([string]$Dir, [string]$Id = 'Testmodul', [string]$MinApp = '0.1.9') {
        Write-TestText (Join-Path $Dir 'addon.json') ('{ "id": "' + $Id + '", "minApp": "' + $MinApp + '" }') $script:Utf8Bom
        Write-TestText (Join-Path $Dir 'Addon.ps1') "# kuenstlich`nWrite-Output 'x'`n" $script:Utf8
        Write-TestText (Join-Path $Dir 'Teil\Tab.xaml') "<Grid/>`n" $script:Utf8
        Write-TestText (Join-Path $Dir '.git\HEAD') "ref: refs/heads/test`n" $script:Utf8
        $dev = Join-Path $Dir '.github\dev'
        [void](New-Item -ItemType Directory -Path $dev -Force)
        foreach ($n in 'Format-Dateien.ps1', 'Deploy-Dev.ps1', 'Test-Start.ps1') { Copy-Item -LiteralPath (Join-Path $script:Vorlage $n) -Destination $dev }
        return $dev
    }
}

Describe 'Argumente maskieren' {
    It 'laesst einfache Argumente unveraendert und setzt sonst Anfuehrungszeichen' {
        ConvertTo-HUProcessArgument 'abc' | Should -Be 'abc'
        ConvertTo-HUProcessArgument 'C:\a b\x.pdf' | Should -Be '"C:\a b\x.pdf"'
        ConvertTo-HUProcessArgument 'C:\a b\' | Should -Be '"C:\a b\\"'
        ConvertTo-HUProcessArgument 'say "hi"' | Should -Be '"say \"hi\""'
        ConvertTo-HUProcessArgument '' | Should -Be '""'
        ConvertTo-HUProcessArgument $null | Should -Be '""'
    }
    It 'kommt nach CommandLineToArgvW unveraendert zurueck' {
        $list = @('einfach', 'mit Leerzeichen', 'C:\Ordner mit\', 'C:\Ordner mit\\', 'say "hi"', 'x\"y', 'a\\"b c', '', "tab`there",
            'Umlaute äöü ß €', '\\server\freigabe x\datei.pdf', '-sOutputFile=C:\a b\c.pdf', '"', '\', 'ende\\')
        $line = 'prog.exe ' + (Join-HUProcessArguments $list)
        $back = [HUToolsTest.Argv]::Split($line)
        $back.Count | Should -Be ($list.Count + 1)
        for ($i = 0; $i -lt $list.Count; $i++) { $back[$i + 1] | Should -BeExactly $list[$i] }
    }
    It 'verbindet nichts zu einem leeren Text' {
        Join-HUProcessArguments @() | Should -Be ''
    }
}

Describe 'Prozess starten' {
    It 'liefert Exit-Code und Ausgabe' {
        $r = Invoke-HUNativeProcess -Exe $script:Cmd -Arguments @('/c', 'echo hallo& exit 3') -TimeoutSec 30
        $r.ExitCode | Should -Be 3
        $r.Out.Trim() | Should -Be 'hallo'
        $r.TimedOut | Should -BeFalse
    }
    It 'liest die Fehlerausgabe getrennt' {
        $r = Invoke-HUNativeProcess -Exe $script:Cmd -Arguments @('/c', 'echo fehler 1>&2') -TimeoutSec 30
        $r.Err.Trim() | Should -Be 'fehler'
        $r.Out.Trim() | Should -Be ''
    }
    It 'setzt Umgebung, PATHVORNE und Arbeitsordner' {
        $d = New-TestDir
        try {
            $r = Invoke-HUNativeProcess -Exe $script:Cmd -Arguments @('/c', 'echo %HUTOOLSTEST%') -Env @{ HUTOOLSTEST = 'wert1' } -TimeoutSec 30
            $r.Out.Trim() | Should -Be 'wert1'
            $r = Invoke-HUNativeProcess -Exe $script:Cmd -Arguments @('/c', 'echo %PATH%') -Env @{ PATHVORNE = 'C:\HUToolsVorne' } -TimeoutSec 30
            $r.Out.Trim() | Should -BeLike 'C:\HUToolsVorne;*'
            $r = Invoke-HUNativeProcess -Exe $script:Cmd -Arguments @('/c', 'cd') -WorkDir $d -TimeoutSec 30
            $r.Out.Trim().TrimEnd('\') | Should -Be $d.TrimEnd('\')
            $env:HUTOOLSTEST | Should -BeNullOrEmpty   # eigene Umgebung bleibt unveraendert
        } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
    It 'schreibt Eingabe als UTF-8 und schliesst sie' {
        $find = Join-Path $env:SystemRoot 'System32\findstr.exe'
        $r = Invoke-HUNativeProcess -Exe $find -Arguments @('^') -StdIn "zeile eins`r`nGrüße €`r`n" -TimeoutSec 30
        $r.ExitCode | Should -Be 0
        $r.Out | Should -Match 'zeile eins'
        $r.Out | Should -Match 'Grüße €'
    }
    It 'haengt ohne Eingabe nicht (Eingabe bekommt Dateiende)' {
        $find = Join-Path $env:SystemRoot 'System32\findstr.exe'
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $r = Invoke-HUNativeProcess -Exe $find -Arguments @('^') -TimeoutSec 20
        $r.TimedOut | Should -BeFalse
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 15
    }
    It 'beendet bei Zeitueberschreitung den ganzen Prozessbaum' {
        $state = @{ Pid = 0 }
        $r = Invoke-HUNativeProcess -Exe $script:Cmd -Arguments @('/c', 'ping -n 47 127.0.0.1 >nul') -TimeoutSec 1 -State $state
        $r.TimedOut | Should -BeTrue
        $r.ExitCode | Should -Be -1
        $r.Err | Should -BeLike 'Zeitüberschreitung nach 1 s'
        $state.Pid | Should -Be 0
        Start-Sleep -Milliseconds 500
        @(Get-CimInstance Win32_Process -Filter "Name='ping.exe'" | Where-Object { "$($_.CommandLine)" -like '*-n 47 127.0.0.1*' }).Count | Should -Be 0
    }
    It 'laesst sich ueber State.Pid aus einem anderen Thread abbrechen' {
        $state = [hashtable]::Synchronized(@{ Pid = 0 })
        $ps = [powershell]::Create()
        try {
            [void]$ps.AddScript('param($m, $exe, $st) Import-Module $m -DisableNameChecking; Invoke-HUNativeProcess -Exe $exe -Arguments @("/c", "ping -n 53 127.0.0.1 >nul") -TimeoutSec 60 -State $st').
            AddArgument($script:ModPath).AddArgument($script:Cmd).AddArgument($state)
            $h = $ps.BeginInvoke()
            $sw = [Diagnostics.Stopwatch]::StartNew()
            while (-not $state.Pid -and $sw.Elapsed.TotalSeconds -lt 20) { Start-Sleep -Milliseconds 50 }
            $state.Pid | Should -BeGreaterThan 0
            Stop-HUProcessTree ([int]$state.Pid)
            $res = @($ps.EndInvoke($h))
            $sw.Elapsed.TotalSeconds | Should -BeLessThan 30
            $res[0].TimedOut | Should -BeFalse
            $res[0].ExitCode | Should -Not -Be 0
            $state.Pid | Should -Be 0
        } finally { $ps.Dispose() }
    }
    It 'wirft, wenn das Programm fehlt' {
        { Invoke-HUNativeProcess -Exe (Join-Path ([IO.Path]::GetTempPath()) 'gibt-es-nicht-4711.exe') -TimeoutSec 5 } | Should -Throw
    }
    It 'Stop-HUProcessTree ignoriert 0 und unbekannte Prozesse' {
        { Stop-HUProcessTree 0 } | Should -Not -Throw
        { Stop-HUProcessTree 999999 } | Should -Not -Throw
    }
}

Describe 'Dateinamen' {
    It 'ersetzt ungueltige Zeichen und fasst zusammen' {
        ConvertTo-HUSafeFileName 'Antrag: Meier/Huber?' | Should -Be 'Antrag_Meier_Huber'
        ConvertTo-HUSafeFileName "Zeile`r`neins`tzwei" | Should -Be 'Zeile_eins_zwei'
        ConvertTo-HUSafeFileName '  Rechnung   März  2026 ' | Should -Be 'Rechnung März 2026'
        ConvertTo-HUSafeFileName ' - Bescheid - ' | Should -Be 'Bescheid'
        ConvertTo-HUSafeFileName 'Äpfel & Öl, Größe' | Should -Be 'Äpfel & Öl, Größe'
    }
    It 'einfache Regel mit -KeepSeparators' {
        ConvertTo-HUSafeFileName 'Antrag: Meier/Huber?' -KeepSeparators | Should -Be 'Antrag_ Meier_Huber'
        ConvertTo-HUSafeFileName ' - Bescheid - ' -KeepSeparators | Should -Be '- Bescheid -'
        ConvertTo-HUSafeFileName 'a  b' -KeepSeparators | Should -Be 'a b'
    }
    It 'leerer Name, Ersatzname, reservierte Namen, Laenge' {
        ConvertTo-HUSafeFileName '   ' | Should -Be 'Dokument'
        ConvertTo-HUSafeFileName '???' -Fallback 'Scan:1' | Should -Be 'Scan_1'
        ConvertTo-HUSafeFileName 'CON' | Should -Be '_CON'
        ConvertTo-HUSafeFileName 'lpt1' | Should -Be '_lpt1'
        ConvertTo-HUSafeFileName 'CONTRA' | Should -Be 'CONTRA'
        (ConvertTo-HUSafeFileName ('x' * 200)).Length | Should -Be 150
        (ConvertTo-HUSafeFileName ('x' * 200) -MaxLength 120).Length | Should -Be 120
        ConvertTo-HUSafeFileName (('x' * 149) + ' .y') | Should -Be ('x' * 149)
    }
    It 'findet freie Datei- und Ordnerpfade' {
        $d = New-TestDir
        try {
            Split-Path -Leaf (Get-HUFreeFilePath $d 'X') | Should -Be 'X.pdf'
            Set-Content -LiteralPath (Join-Path $d 'X.pdf') -Value 'a'
            Split-Path -Leaf (Get-HUFreeFilePath $d 'X') | Should -Be 'X_2.pdf'
            Set-Content -LiteralPath (Join-Path $d 'X_2.pdf') -Value 'a'
            Split-Path -Leaf (Get-HUFreeFilePath $d 'X' '.pdf') | Should -Be 'X_3.pdf'
            Split-Path -Leaf (Get-HUFreeFilePath $d 'X' 'txt') | Should -Be 'X.txt'
            Set-Content -LiteralPath (Join-Path $d 'Ohne') -Value 'a'
            Split-Path -Leaf (Get-HUFreeFilePath $d 'Ohne' '') | Should -Be 'Ohne_2'
            Split-Path -Leaf (Get-HUFreeDirectoryPath $d 'Teil') | Should -Be 'Teil'
            [void](New-Item -ItemType Directory -Path (Join-Path $d 'Teil'))
            Split-Path -Leaf (Get-HUFreeDirectoryPath $d 'Teil') | Should -Be 'Teil_2'
            Split-Path -Parent (Get-HUFreeFilePath $d 'X') | Should -Be $d
        } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
}

Describe 'Vorlage Format-Dateien' {
    It 'bringt Textdateien auf BOM/CRLF und .yml/.py auf LF ohne BOM' {
        $d = New-TestDir 'HUAmtVorlage'
        try {
            $repo = Join-Path $d 'Repo'
            $dev = New-FakeRepo $repo
            Write-TestText (Join-Path $repo '.github\workflows\release.yml') "a: 1`r`nb: 2`r`n" $script:Utf8Bom
            Write-TestText (Join-Path $repo 'werkzeug.py') "print(1)`r`n" $script:Utf8
            Write-TestText (Join-Path $repo '.gitignore') "Data/`n" $script:Utf8
            [IO.File]::WriteAllBytes((Join-Path $repo 'bild.bin'), [byte[]](1, 10, 2))
            & (Join-Path $dev 'Format-Dateien.ps1') 6>$null
            $a = Join-Path $repo 'Addon.ps1'
            Test-Bom $a | Should -BeTrue
            [IO.File]::ReadAllText($a) | Should -Be "# kuenstlich`r`nWrite-Output 'x'`r`n"
            $y = Join-Path $repo '.github\workflows\release.yml'
            Test-Bom $y | Should -BeFalse
            [IO.File]::ReadAllText($y) | Should -Be "a: 1`nb: 2`n"
            [IO.File]::ReadAllText((Join-Path $repo 'werkzeug.py')) | Should -Be "print(1)`n"
            Test-Bom (Join-Path $repo '.gitignore') | Should -BeFalse
            [IO.File]::ReadAllText((Join-Path $repo '.gitignore')) | Should -Be "Data/`r`n"
            [IO.File]::ReadAllBytes((Join-Path $repo 'bild.bin')) -join ',' | Should -Be '1,10,2'
            [IO.File]::ReadAllText((Join-Path $repo '.git\HEAD')) | Should -Be "ref: refs/heads/test`n"
        } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
    It 'beachtet formatAusschluss und ohneFormat aus modul-dev.json' {
        $d = New-TestDir 'HUAmtVorlage'
        try {
            $repo = Join-Path $d 'Repo'
            $dev = New-FakeRepo $repo
            New-Item -ItemType Directory -Force (Join-Path $repo 'web') | Out-Null
            $h = Join-Path $repo 'web\index.html'
            Write-TestText $h "<p>x</p>`n" $script:Utf8
            Write-TestText (Join-Path $dev 'modul-dev.json') '{ "formatAusschluss": "^web\\\\" }' $script:Utf8Bom
            & (Join-Path $dev 'Format-Dateien.ps1') 6>$null
            Test-Bom $h | Should -BeFalse
            [IO.File]::ReadAllText($h) | Should -Be "<p>x</p>`n"
            Test-Bom (Join-Path $repo 'Addon.ps1') | Should -BeTrue
            Write-TestText (Join-Path $repo 'neu.ps1') "x`n" $script:Utf8
            Write-TestText (Join-Path $dev 'modul-dev.json') '{ "ohneFormat": true }' $script:Utf8Bom
            & (Join-Path $dev 'Format-Dateien.ps1') 6>$null
            Test-Bom (Join-Path $repo 'neu.ps1') | Should -BeFalse
        } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
}

Describe 'Vorlage Deploy-Dev' {
    BeforeEach {
        $script:D = New-TestDir 'HUAmtVorlage'
        $script:App = Join-Path $script:D 'App'
        $script:Repo = Join-Path $script:D 'Repo'
        New-FakeApp $script:App
        $script:Dev = New-FakeRepo $script:Repo
        $script:Deploy = Join-Path $script:Dev 'Deploy-Dev.ps1'
        $script:Ziel = Join-Path $script:App 'AddonsDev\Testmodul'
    }
    AfterEach { Remove-Item -LiteralPath $script:D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'kopiert das Modul ohne Punkt-Ordner und formatiert die Quelle' {
        & $script:Deploy -App $script:App 6>$null
        Test-Path -LiteralPath (Join-Path $script:Ziel 'addon.json') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Ziel 'Teil\Tab.xaml') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Ziel '.github') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:Ziel '.git') | Should -BeFalse
        Test-Bom (Join-Path $script:Repo 'Addon.ps1') | Should -BeTrue
        [IO.File]::ReadAllText((Join-Path $script:Repo 'Teil\Tab.xaml')) | Should -Be "<Grid/>`r`n"
        # andere Module und Daten der Test-App bleiben
        Test-Path -LiteralPath (Join-Path $script:App 'AddonsDev\Anderes\addon.json') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:App 'Data\daten.json') | Should -BeTrue
        ([IO.File]::ReadAllText((Join-Path $script:Ziel 'addon.json')) | ConvertFrom-Json).minApp | Should -Be '0.1.9'
    }
    It 'ersetzt ein altes Ziel vollstaendig' {
        Write-TestText (Join-Path $script:Ziel 'alt.ps1') '# alt' $script:Utf8Bom
        & $script:Deploy -App $script:App 6>$null
        Test-Path -LiteralPath (Join-Path $script:Ziel 'alt.ps1') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:Ziel 'Addon.ps1') | Should -BeTrue
    }
    It 'bricht bei Dateien ab, die nicht ins Modul gehoeren, und laesst das Ziel stehen' {
        Write-TestText (Join-Path $script:Ziel 'alt.ps1') '# alt' $script:Utf8Bom
        Write-TestText (Join-Path $script:Repo 'Teil\liste.csv') 'a;b' $script:Utf8Bom
        { & $script:Deploy -App $script:App 6>$null } | Should -Throw '*Teil\liste.csv*'
        Test-Path -LiteralPath (Join-Path $script:Ziel 'alt.ps1') | Should -BeTrue
        Remove-Item -LiteralPath (Join-Path $script:Repo 'Teil\liste.csv')
        Write-TestText (Join-Path $script:Repo 'einstellungen.json') '{}' $script:Utf8Bom
        { & $script:Deploy -App $script:App 6>$null } | Should -Throw '*einstellungen.json*'
        Remove-Item -LiteralPath (Join-Path $script:Repo 'einstellungen.json')
        [void](New-Item -ItemType Directory -Path (Join-Path $script:Repo '__pycache__'))
        { & $script:Deploy -App $script:App 6>$null } | Should -Throw '*__pycache__*'
    }
    It 'nimmt Einstellungen aus modul-dev.json, Parameter gehen vor' {
        Write-TestText (Join-Path $script:Repo 'Testseite.pdf') '%PDF-1.4' $script:Utf8
        Write-TestText (Join-Path $script:Repo 'web\lib\fremd.html') "<p>x</p>`n" $script:Utf8
        $cfg = @{ erlaubteNamen = @('Testseite.pdf'); formatAusschluss = '^web\\lib\\'; minAppAnpassen = $true } | ConvertTo-Json
        Write-TestText (Join-Path $script:Dev 'modul-dev.json') $cfg $script:Utf8Bom
        & $script:Deploy -App $script:App 6>$null 3>$null
        Test-Path -LiteralPath (Join-Path $script:Ziel 'Testseite.pdf') | Should -BeTrue
        [IO.File]::ReadAllText((Join-Path $script:Repo 'web\lib\fremd.html')) | Should -Be "<p>x</p>`n"
        ([IO.File]::ReadAllText((Join-Path $script:Ziel 'addon.json')) | ConvertFrom-Json).minApp | Should -Be '0.1.5'
        ([IO.File]::ReadAllText((Join-Path $script:Repo 'addon.json')) | ConvertFrom-Json).minApp | Should -Be '0.1.9'
        { & $script:Deploy -App $script:App -ErlaubteNamen @() 6>$null 3>$null } | Should -Throw '*Testseite.pdf*'
    }
    It 'nimmt -ModulId und -Quelle ohne addon.json-Id' {
        & $script:Deploy -App $script:App -ModulId 'Anderer_Name' -Quelle $script:Repo 6>$null
        Test-Path -LiteralPath (Join-Path $script:App 'AddonsDev\Anderer_Name\Addon.ps1') | Should -BeTrue
        { & $script:Deploy -App $script:App -ModulId '..\raus' 6>$null } | Should -Throw '*Modul-Id*'
    }
    It 'bricht ab, wenn keine Test-App im Ziel liegt' {
        { & $script:Deploy -App (Join-Path $script:D 'Leer') 6>$null } | Should -Throw '*Main.ps1*'
    }
    It 'bricht ab, wenn die Test-App laeuft' {
        Write-TestText (Join-Path $script:Ziel 'alt.ps1') '# alt' $script:Utf8Bom
        $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $main = Join-Path $script:App 'Main.ps1'
        # harmloser Prozess, dessen Befehlszeile wie eine laufende App aussieht (Main.ps1 wird nicht ausgefuehrt)
        $p = Start-Process -FilePath $ps -ArgumentList @('-NoProfile', '-NonInteractive', '-Command', "Start-Sleep -Seconds 60 # `"$main`"") -WindowStyle Hidden -PassThru
        try {
            Start-Sleep -Milliseconds 500
            { & $script:Deploy -App $script:App 6>$null } | Should -Throw '*Test-App l*'
            Test-Path -LiteralPath (Join-Path $script:Ziel 'alt.ps1') | Should -BeTrue
        } finally { Stop-HUProcessTree $p.Id }
    }
    It 'bricht ab, wenn eine Datei im Ziel gesperrt ist' {
        $f = Join-Path $script:Ziel 'Sperre.dll'
        Write-TestText $f 'x' $script:Utf8
        $fs = [IO.File]::Open($f, 'Open', 'Read', 'Read')
        try {
            { & $script:Deploy -App $script:App 6>$null } | Should -Throw '*gesperrt*'
            Test-Path -LiteralPath $f | Should -BeTrue
        } finally { $fs.Dispose() }
    }
}

Describe 'Vorlage Test-Start' {
    BeforeAll {
        $script:D2 = New-TestDir 'HUAmtVorlage'
        $script:App2 = Join-Path $script:D2 'App'
        $script:Repo2 = Join-Path $script:D2 'Repo'
        New-FakeApp $script:App2
        $script:Dev2 = New-FakeRepo $script:Repo2
        $script:Start = Join-Path $script:Dev2 'Test-Start.ps1'
        $script:AppVorher = @(Get-ChildItem -LiteralPath $script:App2 -Recurse -Force | ForEach-Object { $_.FullName.Substring($script:App2.Length) }) -join '|'
    }
    AfterAll { Remove-Item -LiteralPath $script:D2 -Recurse -Force -ErrorAction SilentlyContinue }

    It 'startet nur dieses Modul ohne Data in einer Kopie (PS 5.1 und 7)' {
        $out = & $script:Start -App $script:App2 -Erwartet 'Ziehen-Testmodul', 'Data=False' 6>&1 3>$null | Out-String
        $out | Should -Match 'PS 5\.1: OK'
        if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) { $out | Should -Match 'PS 7: OK' }
        $out | Should -Not -Match 'Ziehen-Anderes'
    }
    It 'meldet fehlende Muster als Fehler' {
        { & $script:Start -App $script:App2 -PowerShell 5 -Erwartet 'Ziehen-GibtEsNicht' 6>$null } | Should -Throw '*fehlgeschlagen*'
    }
    It 'nimmt erwartete Muster aus modul-dev.json' {
        Write-TestText (Join-Path $script:Dev2 'modul-dev.json') '{ "erwartet": ["Ziehen-Fehlt"] }' $script:Utf8Bom
        try {
            { & $script:Start -App $script:App2 -PowerShell 5 6>$null } | Should -Throw '*fehlgeschlagen*'
            { & $script:Start -App $script:App2 -PowerShell 5 -Erwartet @() 6>$null } | Should -Not -Throw
        } finally { Remove-Item -LiteralPath (Join-Path $script:Dev2 'modul-dev.json') }
    }
    It 'veraendert die Test-App nicht und raeumt den Temp-Ordner auf' {
        $nachher = @(Get-ChildItem -LiteralPath $script:App2 -Recurse -Force | ForEach-Object { $_.FullName.Substring($script:App2.Length) }) -join '|'
        $nachher | Should -Be $script:AppVorher
        @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'HUAmtStart-Testmodul-*').Count | Should -Be 0
    }
}
