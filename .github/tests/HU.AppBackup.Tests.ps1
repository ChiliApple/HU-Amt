#Requires -Version 5.1
# Pester 5: Sicherung (Core\HU.AppBackup.psm1) und Support-Text. Nur kuenstliche Daten unter TEMP.

BeforeAll {
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

Describe 'Automatik' {
    It 'taeglich: faellig am naechsten Kalendertag, nicht am selben' {
        $now = [datetime]'2026-11-03T08:00:00'
        Test-HUAppBackupDue -Mode 'taeglich' -Last '2026-11-03T01:00:00' -Now $now | Should -BeFalse
        Test-HUAppBackupDue -Mode 'taeglich' -Last '2026-11-02T23:00:00' -Now $now | Should -BeTrue
        Test-HUAppBackupDue -Mode 'woechentlich' -Last '2026-10-30T08:00:00' -Now $now | Should -BeFalse
        Test-HUAppBackupDue -Mode 'aus' -Last '' -Now $now | Should -BeFalse
        Test-HUAppBackupDue -Mode 'taeglich' -Last '' -Now $now | Should -BeTrue
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
