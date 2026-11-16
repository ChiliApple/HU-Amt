#Requires -Version 5.1
# Pester 5: Sicherung (Core\HU.AppBackup.psm1) und Support-Text. Nur kuenstliche Daten unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.AppBackup.psm1') -Force -DisableNameChecking
    $script:Pw = 'Testkennwort-123'

    function New-TestData([string]$Base) {
        $files = @{
            'Verwalter.txt'                           = 'anna'
            'sicherung.json'                          = '{}'
            'Beispiel\verlauf.txt'                    = 'a'
            'Beispiel\_Kopien\Liste_20261103-101010.xlsx' = 'k'
            'Benutzer\anna\ui-state.json'             = '{}'
            'Benutzer\anna\sicherung-kennwort.dat'    = 'geheim'
            'Benutzer\anna\Logs\HU-Amt_2026-11-03.log' = 'log'
            'Benutzer\anna\Token\GitHub.xml'          = 'tok'
            'Laufend\anna_PC-1.lock'                  = 'x'
            'Docs\Beispiel\index.html'                = 'h'
            'Wiederherstellung\2026-11-01_101010\x'   = 'v'
            'Beispiel\Liste.xlsx.lock'                = 'l'
            '_Test\Beispiel\verlauf.txt'              = 't'
        }
        foreach ($k in $files.Keys) {
            $p = Join-Path $Base $k
            [void](New-Item -ItemType Directory -Path (Split-Path $p -Parent) -Force)
            [IO.File]::WriteAllText($p, $files[$k])
        }
    }
}

Describe 'Was gesichert wird' {
    BeforeAll {
        $script:Data = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtAb-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-TestData $script:Data
        $script:Names = @(Get-HUAppBackupSources -DataRoot $script:Data | ForEach-Object { $_.Name } | Sort-Object)
    }
    AfterAll { Remove-Item -LiteralPath $script:Data -Recurse -Force -ErrorAction SilentlyContinue }

    It 'nimmt Moduldaten, Kopien, Benutzer-Einstellungen und Verwalter mit' {
        $script:Names | Should -Contain 'data/Verwalter.txt'
        $script:Names | Should -Contain 'data/sicherung.json'
        $script:Names | Should -Contain 'data/Beispiel/verlauf.txt'
        $script:Names | Should -Contain 'data/Beispiel/_Kopien/Liste_20261103-101010.xlsx'
        $script:Names | Should -Contain 'data/Benutzer/anna/ui-state.json'
    }
    It 'laesst Logs, Token, Sperren, Hilfe, Vorstand, Testbestand und das gespeicherte Kennwort weg' {
        @($script:Names | Where-Object { $_ -match '/Logs/|/Token/|^data/Laufend/|^data/Docs/|^data/Wiederherstellung/|^data/_Test/|\.lock$|sicherung-kennwort\.dat$' }).Count | Should -Be 0
        $script:Names.Count | Should -Be 5
    }
    It 'Bereiche: je Modul, Benutzer, Verwalter, Allgemein' {
        Get-HUAppBackupArea 'data/Beispiel/verlauf.txt' | Should -Be 'Daten: Beispiel'
        Get-HUAppBackupArea 'data/Benutzer/anna/ui-state.json' | Should -Be 'Benutzer (Fenster und Einstellungen)'
        Get-HUAppBackupArea 'data/Verwalter.txt' | Should -Be 'Verwalter'
        Get-HUAppBackupArea 'data/sicherung.json' | Should -Be 'Allgemein'
        Test-HUAppBackupCodeName 'data/Verwalter.txt' | Should -BeTrue
    }
    It 'Ziele: nur erlaubte Namen' {
        Resolve-HUAppBackupTarget 'data/Beispiel/a.txt' 'C:\App' 'C:\App\Data' | Should -Be 'C:\App\Data\Beispiel\a.txt'
        Resolve-HUAppBackupTarget 'data/../Main.ps1' 'C:\App' 'C:\App\Data' | Should -Be ''
        Resolve-HUAppBackupTarget 'data/Laufend/a.lock' 'C:\App' 'C:\App\Data' | Should -Be ''
        Resolve-HUAppBackupTarget 'data/Benutzer/anna/Token/GitHub.xml' 'C:\App' 'C:\App\Data' | Should -Be ''
        Resolve-HUAppBackupTarget 'config/update.json' 'C:\App' 'C:\App\Data' | Should -Be ''
        Resolve-HUAppBackupTarget 'data/C:/x' 'C:\App' 'C:\App\Data' | Should -Be ''
    }
}

Describe 'Sichern und Wiederherstellen' {
    BeforeAll {
        $script:Base = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtAbRt-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $script:Data = Join-Path $script:Base 'App\Data'
        $script:Out = Join-Path $script:Base 'Sicherungen'
        New-TestData $script:Data
        $src = @(Get-HUAppBackupSources -DataRoot $script:Data)
        $script:Res = New-HUAppBackup -Sources $src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1' -Keep 2
    }
    AfterAll { Remove-Item -LiteralPath $script:Base -Recurse -Force -ErrorAction SilentlyContinue }

    It 'schreibt eine geprüfte Datei mit festem Namensmuster' {
        Test-Path $script:Res.Path | Should -BeTrue
        Split-Path $script:Res.Path -Leaf | Should -Match '^HU-Sicherung_PC-1_\d{4}-\d{2}-\d{2}_\d{6}\.husicherung$'
        $script:Res.Files | Should -Be 5
    }
    It 'falsches Kennwort: nichts entschluesselt' {
        $o = Open-HUAppBackup -Path $script:Res.Path -Password 'Falsches-Kennwort-1' -AppVersion '0.0.1' -AppRoot (Join-Path $script:Base 'App') -DataRoot $script:Data
        @($o.Errors).Count | Should -BeGreaterThan 0
        $o.Entries.Count | Should -Be 0
    }
    It 'Sicherung aus neuerer Version wird abgelehnt' {
        $o = Open-HUAppBackup -Path $script:Res.Path -Password $script:Pw -AppVersion '0.0.0' -AppRoot (Join-Path $script:Base 'App') -DataRoot $script:Data
        (@($o.Errors) -join ' ') | Should -Match 'neuer|aktualisieren|Version'
    }
    It 'stellt gewaehlte Bereiche wieder her und legt den vorherigen Stand ab' {
        [IO.File]::WriteAllText((Join-Path $script:Data 'Beispiel\verlauf.txt'), 'geaendert')
        $o = Open-HUAppBackup -Path $script:Res.Path -Password $script:Pw -AppVersion '0.0.1' -AppRoot (Join-Path $script:Base 'App') -DataRoot $script:Data
        @($o.Errors).Count | Should -Be 0
        $w = Restore-HUAppBackup -Opened $o -Areas @('Daten: Beispiel') -AppRoot (Join-Path $script:Base 'App') -DataRoot $script:Data
        $w.Written | Should -Be 2
        [IO.File]::ReadAllText((Join-Path $script:Data 'Beispiel\verlauf.txt')) | Should -Be 'a'
        [IO.File]::ReadAllText((Join-Path $w.Before 'data\Beispiel\verlauf.txt')) | Should -Be 'geaendert'
        [IO.File]::ReadAllText((Join-Path $script:Data 'Verwalter.txt')) | Should -Be 'anna'
    }
    It 'zu kurzes Kennwort wird abgelehnt' {
        { New-HUAppBackup -Sources @() -Password 'kurz' -Folder $script:Out -AppVersion '0.0.1' } | Should -Throw
    }
}

Describe 'Sichern: verschwundene Dateien und Grenzen (V6)' {
    BeforeEach {
        $script:Base = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtAbV6-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $script:Data = Join-Path $script:Base 'App\Data'
        $script:Out = Join-Path $script:Base 'Sicherungen'
        New-TestData $script:Data
        $script:Src = @(Get-HUAppBackupSources -DataRoot $script:Data)
        $script:LimitsBefore = & (Get-Module HU.AppBackup) { $script:HUAbLimits.Clone() }
    }
    AfterEach {
        $lb = $script:LimitsBefore
        & (Get-Module HU.AppBackup) { param($l) $script:HUAbLimits = $l } $lb
        Remove-Item -LiteralPath $script:Base -Recurse -Force -ErrorAction SilentlyContinue
    }
    It 'Datei verschwindet nach dem Sammeln: Sicherung entsteht trotzdem, die Datei wird genannt' {
        Remove-Item -LiteralPath (Join-Path $script:Data 'Beispiel\verlauf.txt') -Force
        $r = New-HUAppBackup -Sources $script:Src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1'
        Test-Path $r.Path | Should -BeTrue
        $r.Files | Should -Be 4
        @($r.Skipped) | Should -Be @('data/Beispiel/verlauf.txt')
        $o = Open-HUAppBackup -Path $r.Path -Password $script:Pw -AppVersion '0.0.1' -AppRoot (Join-Path $script:Base 'App') -DataRoot $script:Data
        @($o.Errors).Count | Should -Be 0
    }
    It 'gesperrte Datei wird nicht still uebergangen: Sicherung bricht ab' {
        $p = Join-Path $script:Data 'Beispiel\verlauf.txt'
        $h = [IO.File]::Open($p, 'Open', 'ReadWrite', 'None')
        try { { New-HUAppBackup -Sources $script:Src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1' } | Should -Throw }
        finally { $h.Dispose() }
        @(Get-ChildItem -LiteralPath $script:Out -Filter '*.husicherung' -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
    It 'mehr Dateien als das Wiederherstellen erlaubt: keine Sicherung, klare Meldung' {
        & (Get-Module HU.AppBackup) { $script:HUAbLimits = @{ Files = 3; EntryBytes = 512MB; TotalBytes = 2GB } }
        { New-HUAppBackup -Sources $script:Src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1' } | Should -Throw '*Wiederherstellen*'
        @(Get-ChildItem -LiteralPath $script:Out -Filter '*.husicherung*' -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
    It 'Datei groesser als erlaubt: keine Sicherung, Meldung nennt die Datei' {
        & (Get-Module HU.AppBackup) { $script:HUAbLimits = @{ Files = 200000; EntryBytes = 1; TotalBytes = 2GB } }
        { New-HUAppBackup -Sources $script:Src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1' } | Should -Throw '*data/*'
    }
    It 'zusammen zu gross: keine Sicherung' {
        & (Get-Module HU.AppBackup) { $script:HUAbLimits = @{ Files = 200000; EntryBytes = 512MB; TotalBytes = 6 } }
        { New-HUAppBackup -Sources $script:Src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1' } | Should -Throw '*Wiederherstellen*'
    }
    It 'ab 80 Prozent einer Grenze: Sicherung entsteht, mit Warnung' {
        & (Get-Module HU.AppBackup) { $script:HUAbLimits = @{ Files = 7; EntryBytes = 512MB; TotalBytes = 2GB } }
        $r = New-HUAppBackup -Sources $script:Src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1'
        Test-Path $r.Path | Should -BeTrue
        "$($r.Warning)" | Should -Match 'Dateien'
    }
    It 'weit unter den Grenzen: keine Warnung, nichts uebersprungen' {
        $r = New-HUAppBackup -Sources $script:Src -Password $script:Pw -Folder $script:Out -AppVersion '0.0.1' -Pc 'PC-1'
        "$($r.Warning)" | Should -Be ''
        @($r.Skipped).Count | Should -Be 0
    }
}

Describe 'Automatik' {
    It 'taeglich: faellig am naechsten Kalendertag, nicht am selben' {
        $now = [datetime]'2026-11-03T08:00:00'
        Test-HUAppBackupDue -Mode 'taeglich' -Last '2026-11-03T01:00:00' -Now $now | Should -BeFalse
        Test-HUAppBackupDue -Mode 'taeglich' -Last '2026-11-02T23:00:00' -Now $now | Should -BeTrue
        Test-HUAppBackupDue -Mode 'woechentlich' -Last '2026-10-30T08:00:00' -Now $now | Should -BeFalse
        Test-HUAppBackupDue -Mode 'aus' -Last '' -Now $now | Should -BeFalse
        Test-HUAppBackupDue -Mode 'taeglich' -Last '' -Now $now | Should -BeTrue
    }
    It 'Warnung: nicht eingerichtet, aus, nie gelaufen, ueberfaellig' {
        $now = [datetime]'2026-11-10T08:00:00'
        Get-HUAppBackupWarning -Mode 'taeglich' -Path '' -Last '' -Now $now | Should -Match 'nicht eingerichtet'
        Get-HUAppBackupWarning -Mode 'aus' -Path 'X:\S' -Last '' -Now $now | Should -Match 'aus'
        Get-HUAppBackupWarning -Mode 'taeglich' -Path 'X:\S' -Last '' -Now $now | Should -Match 'noch keine'
        Get-HUAppBackupWarning -Mode 'taeglich' -Path 'X:\S' -Last '2026-11-09T23:00:00' -Now $now | Should -Be ''
        Get-HUAppBackupWarning -Mode 'taeglich' -Path 'X:\S' -Last '2026-11-08T08:00:00' -Now $now | Should -Match '2 Tage'
        Get-HUAppBackupWarning -Mode 'woechentlich' -Path 'X:\S' -Last '2026-11-02T08:00:00' -Now $now | Should -Be ''
        Get-HUAppBackupWarning -Mode 'woechentlich' -Path 'X:\S' -Last '2026-11-01T08:00:00' -Now $now | Should -Match '9 Tage'
    }
}

Describe 'Support' {
    BeforeAll {
        $script:AppRoot = 'C:\Tools\HU-Amt'
        . (Join-Path $script:Root 'Functions\UI-Support.ps1')
    }
    It 'macht Programmordner, Benutzer, PC, Adressen und IDs unkenntlich' {
        $t = ConvertTo-HURedacted "C:\Tools\HU-Amt\Main.ps1 C:\Users\$env:USERNAME\Downloads $env:USERNAME $env:COMPUTERNAME max@example.org 203.0.113.7 12345678-1234-1234-1234-123456789abc" (Get-HURedactMap)
        $t | Should -Not -Match ([regex]::Escape($env:USERNAME))
        $t | Should -Not -Match ([regex]::Escape($env:COMPUTERNAME))
        $t | Should -Match '<Programmordner>'
        $t | Should -Match '<anonym>'
        $t | Should -Match '<mail>'
        $t | Should -Match '<ip>'
        $t | Should -Not -Match '123456789abc'
    }
}

Describe 'Support: UNC-Pfade (Pruefung 086, KERN-U3)' {
    BeforeAll {
        $script:AppRoot = '\\srv-probe\apps\HU-Amt'
        . (Join-Path $script:Root 'Functions\UI-Support.ps1')
    }
    It 'ersetzt Server und Freigabe in <In>' -TestCases @(
        @{ In = 'Fehler in \\fs01\Gemeinde\Belege\2026\a.pdf'; Out = 'Fehler in \\<server>\<freigabe>\Belege\2026\a.pdf' }
        @{ In = 'Pfad "\\203.0.113.9\daten$\x.json"'; Out = 'Pfad "\\<server>\<freigabe>\x.json"' }
        @{ In = '{"Ordner":"\\\\fs01\\Gemeinde\\Eingang"}'; Out = '{"Ordner":"\\\\<server>\\<freigabe>\\Eingang"}' }
        @{ In = '\\fs01\Gemeinde'; Out = '\\<server>\<freigabe>' }
    ) {
        param($In, $Out)
        ConvertTo-HURedacted $In | Should -Be $Out
    }
    It 'laesst lokale Pfade, \\?\ und \\.\ stehen' -TestCases @(
        @{ In = 'C:\Tools\x.log' }
        @{ In = '{"Pfad":"C:\\Users\\Public\\x"}' }
        @{ In = '\\?\C:\lang\pfad' }
        @{ In = '\\.\pipe\probe' }
    ) {
        param($In)
        ConvertTo-HURedacted $In | Should -Be $In
    }
    It 'Programmordner auf einer Freigabe bleibt <Programmordner>' {
        ConvertTo-HURedacted '\\srv-probe\apps\HU-Amt\Main.ps1' (Get-HURedactMap) | Should -Be '<Programmordner>\Main.ps1'
    }
}

Describe 'Jetzt sichern: Sperre nicht moeglich (ab 1.0.8)' {
    BeforeAll {
        Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking
        $script:Base = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtAbL-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $script:AppRoot = Join-Path $script:Base 'App'
        $script:DataRoot = Join-Path $script:Base 'Daten'
        $script:UserRoot = Join-Path $script:DataRoot 'Benutzer\anna'
        $script:Paths = $null; $script:UserName = 'anna'; $script:Version = '1.0.8'; $script:Window = $null
        foreach ($d in $script:AppRoot, $script:UserRoot) { [void](New-Item -ItemType Directory -Path $d -Force) }
        $out = (Join-Path $script:Base 'Sicherung').Replace('\', '\\')   # JSON
        [IO.File]::WriteAllText((Join-Path $script:DataRoot 'sicherung.json'), "{ ""Folder"": ""$out"", ""Mode"": ""taeglich"", ""Keep"": 5 }")
        $script:Msgs = New-Object System.Collections.Generic.List[string]
        $script:Toasts = New-Object System.Collections.Generic.List[string]
        $script:Log = New-Object System.Collections.Generic.List[string]
        function Read-HUJsonFile([string]$Path) { if (Test-Path -LiteralPath $Path) { return ([IO.File]::ReadAllText($Path) | ConvertFrom-Json) } return $null }
        function Show-HUMessage { param([string]$Text, [string]$Title, [string]$Icon = 'Info', $Owner = $null) $script:Msgs.Add("$Icon|$Text") }
        function Show-HUToast { param([string]$Title, [string]$Text, [string]$Icon = 'Info', [scriptblock]$OnClick) $script:Toasts.Add("$Icon|$Title|$Text") }
        function Write-HULogInfo { param($m) $script:Log.Add("INFO|$m") }
        function Write-HULogWarn { param($m) $script:Log.Add("WARN|$m") }
        function Write-HULogOK { param($m) $script:Log.Add("OK|$m") }
        function Write-HULogError { param($m) $script:Log.Add("ERROR|$m") }
        function Test-HUJobRunning { param($Name) return $false }
        function Test-HUVerwalter { return $true }
        . (Join-Path $script:Root 'Functions\UI-AppBackup.ps1')
        $script:Dir = $script:DataRoot
        function New-FailLock([string]$Reason, [switch]$NoReason) {
            $o = [ordered]@{ Ok = $false; Path = (Join-Path $script:DataRoot 'sicherung.json'); LockPath = ''; Stream = $null; UserName = ''; ComputerName = ''; Since = $null }
            if ($Reason -eq 'busy') { $o.UserName = 'berta'; $o.ComputerName = 'PC-2' }
            if (-not $NoReason) { $o['Reason'] = $Reason }
            return [pscustomobject]$o
        }
    }
    AfterAll { Remove-Item -LiteralPath $script:Base -Recurse -Force -ErrorAction SilentlyContinue }
    BeforeEach {
        $script:Msgs.Clear(); $script:Toasts.Clear(); $script:Log.Clear(); $script:AbLock = $null; $script:AbDlg = $null
        Mock Get-HUAppBackupPassword { return 'Testkennwort-123' }
    }

    It 'Entscheidung ohne Fenster: rights und unreachable sind Fehler, busy und ohne Reason nicht' {
        $r = Get-HUAbLockFailure (New-FailLock 'rights')
        $r.IsError | Should -BeTrue
        $r.Text | Should -Not -BeLike '*Es läuft gerade eine Sicherung*'
        $r.Text | Should -BeLike 'Die Sicherung kann nicht starten.*Kein Schreibrecht im Ordner*'
        $r.Info | Should -BeLike 'Kein Schreibrecht im Ordner*'
        (Get-HUAbLockFailure (New-FailLock 'unreachable')).IsError | Should -BeTrue
        foreach ($b in (New-FailLock 'busy'), (New-FailLock 'busy' -NoReason)) {
            $q = Get-HUAbLockFailure $b
            $q.IsError | Should -BeFalse
            $q.Text | Should -BeExactly "Es läuft gerade eine Sicherung.`n`nGerade in Bearbeitung von berta (PC-2). Du kannst die Daten nur ansehen."
        }
    }
    It 'kein Schreibrecht, Klick: Meldung ohne "Es läuft gerade eine Sicherung"' {
        Mock Lock-HUDataFile { New-FailLock 'rights' }
        Start-HUAppBackupNow
        $script:Msgs.Count | Should -Be 1
        $script:Msgs[0] | Should -Not -BeLike '*Es läuft gerade eine Sicherung*'
        $script:Msgs[0] | Should -BeExactly "Warning|Die Sicherung kann nicht starten.`n`nKein Schreibrecht im Ordner $($script:Dir). Bitte an den Verwalter wenden."
        $script:AbLock | Should -BeNullOrEmpty
    }
    It 'kein Schreibrecht, automatisch: Protokoll und Hinweis wie bei einer fehlgeschlagenen Sicherung' {
        Mock Lock-HUDataFile { New-FailLock 'rights' }
        Start-HUAppBackupNow -Auto
        $script:Msgs.Count | Should -Be 0
        $script:Log | Should -Contain "WARN|App-Sicherung (automatisch): Kein Schreibrecht im Ordner $($script:Dir). Bitte an den Verwalter wenden."
        $script:Toasts.Count | Should -Be 1
        $script:Toasts[0] | Should -BeLike "Error|App-Sicherung fehlgeschlagen|Kein Schreibrecht im Ordner*Klick: Einstellungen > Sicherung"
    }
    It 'Ordner nicht erreichbar, automatisch: ebenso ein Fehler' {
        Mock Lock-HUDataFile { New-FailLock 'unreachable' }
        Start-HUAppBackupNow -Auto
        @($script:Log | Where-Object { $_ -like 'WARN|App-Sicherung (automatisch): Der Ordner * ist nicht erreichbar.*' }).Count | Should -Be 1
        $script:Toasts.Count | Should -Be 1
    }
    It 'ein anderer sichert gerade (busy und ohne Reason): Text wie bisher, automatisch still' {
        foreach ($nr in $false, $true) {
            $script:Msgs.Clear(); $script:Toasts.Clear(); $script:Log.Clear()
            $script:NoR = $nr
            Mock Lock-HUDataFile { New-FailLock 'busy' -NoReason:$script:NoR }
            Start-HUAppBackupNow
            $script:Msgs.Count | Should -Be 1
            $script:Msgs[0] | Should -BeExactly "Info|Es läuft gerade eine Sicherung.`n`nGerade in Bearbeitung von berta (PC-2). Du kannst die Daten nur ansehen."
            Start-HUAppBackupNow -Auto
            $script:Msgs.Count | Should -Be 1
            $script:Toasts.Count | Should -Be 0
            @($script:Log | Where-Object { $_ -like 'WARN*' -or $_ -like 'ERROR*' }).Count | Should -Be 0
        }
    }
}
