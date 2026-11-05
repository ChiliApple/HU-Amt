#Requires -Version 5.1
# Pester 5: Datenordner und Sperren (Core\HU.Paths.psm1). Nur kuenstliche Daten in einem Testordner unter TEMP.

BeforeAll {
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking
}

Describe 'Datenordner' {
    BeforeEach {
        $script:App = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:App)
    }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'legt Data, Benutzer und Laufend an' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'Test.Person' -PcName 'PC-1'
        $p.DataRoot | Should -Be (Join-Path $script:App 'Data')
        $p.UserRoot | Should -Be (Join-Path $script:App 'Data\Benutzer\Test.Person')
        Test-Path $p.UserRoot | Should -BeTrue
        Test-Path $p.RunDir | Should -BeTrue
    }
    It 'macht aus heiklen Zeichen im Benutzernamen einen sicheren Ordnernamen' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'a\b:c' -PcName 'PC-1'
        Split-Path $p.UserRoot -Leaf | Should -Be 'a_b_c'
        Split-Path (Split-Path $p.UserRoot -Parent) -Leaf | Should -Be 'Benutzer'
    }
    It 'Modulordner: gueltige Id wird angelegt' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'u' -PcName 'pc'
        $m = Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Beispiel'
        Test-Path $m | Should -BeTrue
    }
    It 'Modulordner: reservierte und ungueltige Ids werden abgelehnt' {
        $p = Initialize-HUDataPaths -AppRoot $script:App -UserName 'u' -PcName 'pc'
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Benutzer' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Postfach' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId 'Wiederherstellung' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId '..\x' } | Should -Throw
        { Get-HUModuleDataPath -DataRoot $p.DataRoot -ModuleId '' } | Should -Throw
    }
    It 'Schreibprobe hinterlaesst keine Datei' {
        Assert-HUWritable $script:App
        @(Get-ChildItem -LiteralPath $script:App -Force).Count | Should -Be 0
    }
}

Describe 'Regel 2: laufende App' {
    BeforeEach {
        $script:Run = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtRun-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Run)
    }
    AfterEach { Remove-Item -LiteralPath $script:Run -Recurse -Force -ErrorAction SilentlyContinue }

    It 'meldet sich an und haelt die Datei offen' {
        $l = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        try {
            Test-Path $l.Path | Should -BeTrue
            Split-Path $l.Path -Leaf | Should -Be 'anna_PC-1.lock'
            { [IO.File]::Delete($l.Path) } | Should -Throw
        } finally { Close-HURunLock $l }
        Test-Path $l.Path | Should -BeFalse
    }
    It 'zweite Anmeldung mit gleichem Namen und PC liefert nichts' {
        $l = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        try { Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1' | Should -BeNullOrEmpty }
        finally { Close-HURunLock $l }
    }
    It 'zaehlt fremde Sitzungen, die eigene nicht' {
        $me = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        $other = Open-HURunLock -RunDir $script:Run -UserName 'bernd' -PcName 'PC-2'
        try {
            $r = @(Get-HURunningInstances -RunDir $script:Run -Self $me)
            $r.Count | Should -Be 1
            $r[0].UserName | Should -Be 'bernd'
            $r[0].ComputerName | Should -Be 'PC-2'
            $r[0].Since | Should -Not -BeNullOrEmpty
        } finally { Close-HURunLock $other; Close-HURunLock $me }
    }
    It 'entfernt verwaiste Lock-Dateien (Absturz)' {
        Set-Content -LiteralPath (Join-Path $script:Run 'alt_PC-9.lock') -Value 'alt|PC-9|2026-01-01T08:00:00|1'
        @(Get-HURunningInstances -RunDir $script:Run).Count | Should -Be 0
        Test-Path (Join-Path $script:Run 'alt_PC-9.lock') | Should -BeFalse
    }
    It 'Anmeldung ueberschreibt eine verwaiste eigene Datei' {
        Set-Content -LiteralPath (Join-Path $script:Run 'anna_PC-1.lock') -Value 'alt'
        $l = Open-HURunLock -RunDir $script:Run -UserName 'anna' -PcName 'PC-1'
        try { $l | Should -Not -BeNullOrEmpty } finally { Close-HURunLock $l }
    }
}

Describe 'Regel 3: gemeinsame Datei bearbeiten' {
    BeforeEach {
        $script:Dir = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtLock-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Dir)
        $script:File = Join-Path $script:Dir 'Liste.xlsx'
        Set-Content -LiteralPath $script:File -Value 'x'
    }
    AfterEach { Remove-Item -LiteralPath $script:Dir -Recurse -Force -ErrorAction SilentlyContinue }

    It 'der Erste bekommt die Sperre, der Zweite sieht, wer bearbeitet' {
        $a = Lock-HUDataFile -Path $script:File -UserName 'anna' -PcName 'PC-1'
        try {
            $a.Ok | Should -BeTrue
            $b = Lock-HUDataFile -Path $script:File -UserName 'bernd' -PcName 'PC-2'
            $b.Ok | Should -BeFalse
            $b.UserName | Should -Be 'anna'
            $b.ComputerName | Should -Be 'PC-1'
            Get-HULockInfoText $b | Should -Match 'anna \(PC-1\)'
            Get-HULockInfoText $b | Should -Match 'nur ansehen'
        } finally { Unlock-HUDataFile $a }
        $c = Lock-HUDataFile -Path $script:File -UserName 'bernd' -PcName 'PC-2'
        try { $c.Ok | Should -BeTrue } finally { Unlock-HUDataFile $c }
        Test-Path "$($script:File).lock" | Should -BeFalse
    }
    It 'eine verwaiste Sperre haelt niemanden auf' {
        Set-Content -LiteralPath "$($script:File).lock" -Value 'alt|PC-9|2026-01-01T08:00:00|1'
        $l = Lock-HUDataFile -Path $script:File -UserName 'anna' -PcName 'PC-1'
        try { $l.Ok | Should -BeTrue } finally { Unlock-HUDataFile $l }
    }
    It 'Unlock einer fremden Sperre tut nichts' {
        $a = Lock-HUDataFile -Path $script:File -UserName 'anna' -PcName 'PC-1'
        try {
            $b = Lock-HUDataFile -Path $script:File -UserName 'bernd' -PcName 'PC-2'
            Unlock-HUDataFile $b
            Test-Path "$($script:File).lock" | Should -BeTrue
        } finally { Unlock-HUDataFile $a }
    }
    It 'Kopie mit Zeitstempel, nur die letzten behalten' {
        1..4 | ForEach-Object { [void](Backup-HUDataFile -Path $script:File -Keep 2) }
        $k = @(Get-ChildItem -LiteralPath (Join-Path $script:Dir '_Kopien') -File)
        $k.Count | Should -Be 2
        $k[0].Name | Should -Match '^Liste_\d{8}-\d{6}(_\d+)?\.xlsx$'
    }
    It 'Aufraeumen der Kopien laesst fremde Dateien mit gleichem Anfang stehen' {
        $kd = Join-Path $script:Dir '_Kopien'
        [void](New-Item -ItemType Directory -Path $kd -Force)
        foreach ($n in 'Liste_alt.xlsx', 'Liste_Notizen.xlsx', 'Liste_20200101-000000.xlsx') { Set-Content -LiteralPath (Join-Path $kd $n) -Value 'y' }
        1..3 | ForEach-Object { [void](Backup-HUDataFile -Path $script:File -Keep 1) }
        Test-Path (Join-Path $kd 'Liste_alt.xlsx') | Should -BeTrue
        Test-Path (Join-Path $kd 'Liste_Notizen.xlsx') | Should -BeTrue
        Test-Path (Join-Path $kd 'Liste_20200101-000000.xlsx') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $kd -File | Where-Object { $_.Name -match '^Liste_\d{8}-\d{6}(_\d+)?\.xlsx$' }).Count | Should -Be 1
    }
    It 'Kopie einer fehlenden Datei liefert leer' {
        Backup-HUDataFile -Path (Join-Path $script:Dir 'fehlt.xlsx') | Should -Be ''
    }
}
