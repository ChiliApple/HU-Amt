#Requires -Version 5.1
# Pester 5: gemeinsame Daten-Bausteine (Core\HU.Data.psm1). Nur kuenstliche Daten in einem Testordner unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:DataModule = Join-Path $script:Root 'Core\HU.Data.psm1'
    Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking
    Import-Module $script:DataModule -Force -DisableNameChecking
    $script:Euro = [string][char]0x20AC
    function New-TestDir {
        $d = Join-Path ([IO.Path]::GetTempPath()) ('HUDataTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $d)
        return $d
    }
    function Set-TestFile([string]$Path, [string]$Text) {
        $dir = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        [IO.File]::WriteAllText($Path, $Text, (New-Object System.Text.UTF8Encoding $true))
    }
}

Describe 'Laden' {
    It 'exportiert die Bausteine und laeuft im strengen Modus' {
        foreach ($f in 'Read-HUJson', 'Save-HUDataFile', 'Wait-HUDataFileLock', 'Get-HUDataRecords', 'Get-HUNextNumber', 'ConvertTo-HUDecimal',
            'Format-HUEuro', 'ConvertFrom-HUDateText', 'New-HUHistoryEntry', 'Add-HUAttachment', 'Remove-HUAttachment', 'Read-HUSettings',
            'Save-HUSettings', 'Test-HUPermittedUser', 'Get-HUUserDisplayName') {
            Get-Command $f -Module HU.Data -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }
        Get-Command Get-HUDataFileBase -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
    }
    It 'laeuft im nackten Hintergrund-Runspace (laedt HU.Paths und HU.Exchange selbst)' {
        $dir = New-TestDir
        try {
            $rs = [runspacefactory]::CreateRunspace()
            $rs.Open()
            $ps = [powershell]::Create()
            $ps.Runspace = $rs
            [void]$ps.AddScript({
                    param($Module, $Dir)
                    Set-StrictMode -Version 1.0
                    Import-Module $Module -DisableNameChecking
                    $p = Join-Path $Dir 'a.json'
                    $s = Save-HUDataFile -Path $p -Stamp '' -UserName 'u1' -PcName 'pc1' -Build { '{"x":1,"zeit":"2026-11-07T15:51:39"}' }
                    $r = Read-HUJson $p
                    [pscustomobject]@{ Ok = $s.Ok; X = $r.Data.x; Zeit = $r.Data.zeit; Global = [bool](Get-Command Lock-HUDataFile -ErrorAction SilentlyContinue | Where-Object { -not $_.Module -or $_.Module.Name -ne 'HU.Paths' }) }
                }).AddArgument($script:DataModule).AddArgument($dir)
            $out = @($ps.Invoke())
            $ps.Streams.Error.Count | Should -Be 0
            $ps.Dispose(); $rs.Dispose()
            $out[0].Ok | Should -BeTrue
            $out[0].X | Should -Be 1
            "$($out[0].Zeit)" | Should -Not -BeNullOrEmpty
        } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'JSON und Dateien' {
    BeforeEach { $script:D = New-TestDir }
    AfterEach { Remove-Item -LiteralPath $script:D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Read-HUJson: fehlt, leer, kaputt, gut' {
        $r = Read-HUJson (Join-Path $script:D 'fehlt.json')
        $r.Data | Should -BeNullOrEmpty; $r.Error | Should -Be ''
        Set-TestFile (Join-Path $script:D 'leer.json') '   '
        (Read-HUJson (Join-Path $script:D 'leer.json')).Error | Should -Be 'Datei ist leer'
        Set-TestFile (Join-Path $script:D 'kaputt.json') '{ "a": '
        (Read-HUJson (Join-Path $script:D 'kaputt.json')).Error | Should -Not -BeNullOrEmpty
        Set-TestFile (Join-Path $script:D 'gut.json') '{ "name": "Ä Test", "n": 3 }'
        $g = Read-HUJson (Join-Path $script:D 'gut.json')
        $g.Error | Should -Be ''; $g.Data.name | Should -Be 'Ä Test'; $g.Data.n | Should -Be 3
    }
    It 'Datumstexte bleiben Texte (5.1 immer, 7 ab 7.5)' {
        $hasKind = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')
        if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $hasKind) { Set-ItResult -Skipped -Because 'ConvertFrom-Json ohne -DateKind (PowerShell 7.0 bis 7.4)'; return }
        $o = ConvertFrom-HUJsonText '{ "zeit": "2026-11-07T15:51:39" }'
        $o.zeit | Should -BeOfType [string]
        $o.zeit | Should -Be '2026-11-07T15:51:39'
    }
    It 'Dateistempel: Ticks|Laenge, fehlt = leer' {
        Get-HUFileStamp (Join-Path $script:D 'nix.json') | Should -Be ''
        $p = Join-Path $script:D 's.json'
        Set-TestFile $p '{}'
        $i = Get-Item -LiteralPath $p
        Get-HUFileStamp $p | Should -Be ('{0}|{1}' -f $i.LastWriteTimeUtc.Ticks, $i.Length)
    }
    It 'Get-HUProp: Objekt, Hashtable, ordered, $null = Standard' {
        $o = [pscustomobject]@{ a = 1; b = $null }
        Get-HUProp $o 'a' 9 | Should -Be 1
        Get-HUProp $o 'b' 9 | Should -Be 9
        Get-HUProp $o 'c' 9 | Should -Be 9
        Get-HUProp @{ a = 2 } 'a' | Should -Be 2
        Get-HUProp ([ordered]@{ a = 3 }) 'a' | Should -Be 3
        Get-HUProp $null 'a' 'x' | Should -Be 'x'
    }
    It 'Zeit als Text: Jetzt und DateTime' {
        Get-HUNowText ([datetime]::new(2026, 11, 7, 8, 5, 3)) | Should -Be '2026-11-07T08:05:03'
        ConvertTo-HUTimeText ([datetime]::new(2026, 1, 2, 3, 4, 5)) | Should -Be '2026-01-02T03:04:05'
        ConvertTo-HUTimeText 'frei' | Should -Be 'frei'
        ConvertTo-HUTimeText $null | Should -Be ''
    }
}

Describe 'Sperre und Speichern' {
    BeforeEach { $script:D = New-TestDir; $script:P = Join-Path $script:D 'sub\daten.json' }
    AfterEach { Remove-Item -LiteralPath $script:D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'neue Datei mit Stand "" (darf noch nicht existieren), zweites Anlegen scheitert' {
        $r = Save-HUDataFile -Path $script:P -Stamp '' -UserName 'u1' -PcName 'pc1' -Build { '{"v":1}' }
        $r.Ok | Should -BeTrue
        $r.Stamp | Should -Be (Get-HUFileStamp $script:P)
        Test-Path (Join-Path $script:D 'sub\_Kopien') | Should -BeFalse
        $r2 = Save-HUDataFile -Path $script:P -Stamp '' -UserName 'u1' -PcName 'pc1' -Changed 'gibt es schon' -Build { '{"v":2}' }
        $r2.Ok | Should -BeFalse; $r2.Message | Should -Be 'gibt es schon'
        (Read-HUJson $script:P).Data.v | Should -Be 1
        Test-Path "$script:P.lock" | Should -BeFalse
    }
    It 'passender Stand: Kopie und Ersetzen; falscher Stand: nichts geschrieben' {
        $r = Save-HUDataFile -Path $script:P -Build { '{"v":1}' }
        $r2 = Save-HUDataFile -Path $script:P -Stamp $r.Stamp -Build { '{"v":2}' }
        $r2.Ok | Should -BeTrue
        @(Get-ChildItem (Join-Path $script:D 'sub\_Kopien') -File).Count | Should -Be 1
        $r3 = Save-HUDataFile -Path $script:P -Stamp $r.Stamp -Build { '{"v":3}' }
        $r3.Ok | Should -BeFalse
        $r3.Message | Should -Match 'anderen PC'
        (Read-HUJson $script:P).Data.v | Should -Be 2
    }
    It '-NoBackup legt keine Kopie an; Objekt wird JSON; Text mit Umlauten bleibt UTF-8 mit BOM' {
        [void](Save-HUDataFile -Path $script:P -Build { '{"v":1}' })
        $obj = [ordered]@{ name = 'Größe'; liste = @(1, 2) }
        $r = Save-HUDataFile -Path $script:P -NoBackup -Build { $obj }
        $r.Ok | Should -BeTrue
        Test-Path (Join-Path $script:D 'sub\_Kopien') | Should -BeFalse
        $b = [IO.File]::ReadAllBytes($script:P)
        ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) | Should -BeTrue
        $d = (Read-HUJson $script:P).Data
        $d.name | Should -Be 'Größe'; @($d.liste).Count | Should -Be 2
    }
    It '-Build wirft: Meldung, nichts geschrieben; -Build $null: Skipped' {
        [void](Save-HUDataFile -Path $script:P -Build { '{"v":1}' })
        $r = Save-HUDataFile -Path $script:P -Build { throw 'Prüfung fehlgeschlagen' }
        $r.Ok | Should -BeFalse; $r.Message | Should -Be 'Prüfung fehlgeschlagen'
        $s = Save-HUDataFile -Path $script:P -Build { $null }
        $s.Ok | Should -BeTrue; $s.Skipped | Should -BeTrue
        $s.Stamp | Should -Be (Get-HUFileStamp $script:P)
        (Read-HUJson $script:P).Data.v | Should -Be 1
        Test-Path "$script:P.lock" | Should -BeFalse
    }
    It '-Build sieht Variablen des Aufrufers und liest unter der Sperre neu' {
        [void](Save-HUDataFile -Path $script:P -Build { '{"n":1}' })
        $plus = 5
        $r = Save-HUDataFile -Path $script:P -Build { $cur = (Read-HUJson $script:P).Data; [ordered]@{ n = $cur.n + $plus } }
        $r.Ok | Should -BeTrue
        (Read-HUJson $script:P).Data.n | Should -Be 6
    }
    It 'fremde Sperre: wartet -WaitMs, dann Meldung mit Name' {
        [void](New-Item -ItemType Directory -Path (Split-Path $script:P) -Force)
        $other = Lock-HUDataFile -Path $script:P -UserName 'maria' -PcName 'PC-12'
        try {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $r = Save-HUDataFile -Path $script:P -UserName 'u1' -PcName 'pc1' -WaitMs 400 -Build { '{}' }
            $sw.Stop()
            $r.Ok | Should -BeFalse
            $r.Message | Should -Match 'maria'
            $sw.ElapsedMilliseconds | Should -BeGreaterOrEqual 350
            Test-Path $script:P | Should -BeFalse
        } finally { Unlock-HUDataFile $other }
    }
    It 'gehaltene Sperre (-Lock) wird genutzt und nicht freigegeben' {
        $l = Wait-HUDataFileLock -Path $script:P -UserName 'u1' -PcName 'pc1'
        $l.Ok | Should -BeTrue
        try {
            $r = Save-HUDataFile -Path $script:P -Lock $l -Build { '{"v":9}' }
            $r.Ok | Should -BeTrue
            Test-Path "$script:P.lock" | Should -BeTrue
        } finally { Unlock-HUDataFile $l }
        Test-Path "$script:P.lock" | Should -BeFalse
    }
    It 'Wait-HUDataFileLock: frei sofort, belegt nach Ablauf mit Ok=$false' {
        $a = Wait-HUDataFileLock -Path $script:P -UserName 'a' -PcName 'pa' -WaitMs 200
        try {
            $a.Ok | Should -BeTrue
            $b = Wait-HUDataFileLock -Path $script:P -UserName 'b' -PcName 'pb' -WaitMs 200
            $b.Ok | Should -BeFalse
            $b.UserName | Should -Be 'a'
        } finally { Unlock-HUDataFile $a }
    }
    It 'Write-HUDataFile -New wirft bei vorhandener Datei' {
        Write-HUDataFile -Path $script:P -Text '{"v":1}' -New
        { Write-HUDataFile -Path $script:P -Text '{"v":2}' -New } | Should -Throw
        Write-HUDataFile -Path $script:P -Text '{"v":3}'
        (Read-HUJson $script:P).Data.v | Should -Be 3
    }
}

Describe 'Kennungen und Nummernkreis' {
    BeforeEach { $script:D = New-TestDir }
    AfterEach { Remove-Item -LiteralPath $script:D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Datum-Zeit-PC-Kennung' {
        $id = New-HUDataId -PcName 'pc-1 ä' -Now ([datetime]::new(2026, 11, 7, 14, 30, 0))
        $id | Should -Match '^20261107-143000-PC1-[0-9a-f]{8}$'
        Test-HUDataId $id | Should -BeTrue
        Test-HUDataId (New-HUDataId -PcName '' ) | Should -BeTrue
        Test-HUDataId ('20261107-143000-' + ('A' * 16) + '-0123abcd') | Should -BeFalse
        Test-HUDataId 'x' | Should -BeFalse
    }
    It 'kurze Kennung mit Praefix' {
        $id = New-HUShortId 'f'
        $id | Should -Match '^f-[0-9a-f]{8}$'
        Test-HUShortId $id -Prefix 'f' | Should -BeTrue
        Test-HUShortId $id -Prefix 'm' | Should -BeFalse
        Test-HUShortId $id | Should -BeTrue
        Test-HUShortId 'r-abc' -Prefix 'r' | Should -BeFalse
        Test-HUShortId 'r-abc' -Prefix 'r' -Loose | Should -BeTrue
        Test-HUShortId 'F-12345678' | Should -BeFalse
        { New-HUShortId 'A b' } | Should -Throw
    }
    It 'Nummernkreis: fortlaufend je Schluessel, ueberspringt Vergebene' {
        $z = Join-Path $script:D 'zaehler.json'
        Get-HUNextNumber -Path $z -Key '2026' | Should -Be '2026-0001'
        Get-HUNextNumber -Path $z -Key '2026' | Should -Be '2026-0002'
        Get-HUNextNumber -Path $z -Key '2027' | Should -Be '2027-0001'
        $vergeben = @('2026-0003', '2026-0004')
        Get-HUNextNumber -Path $z -Key '2026' -Exists { param($n) $vergeben -contains $n } | Should -Be '2026-0005'
        Get-HUNextNumber -Path $z -Key 'R' -Format '{0}{1:D3}' | Should -Be 'R001'
        $d = (Read-HUJson $z).Data
        $d.'2026' | Should -Be 5; $d.'2027' | Should -Be 1
        Test-Path "$z.lock" | Should -BeFalse
    }
    It 'Nummernkreis: kaputter Zaehler wirft, statt neu zu beginnen' {
        $z = Join-Path $script:D 'zaehler.json'
        Set-TestFile $z '{ kaputt'
        { Get-HUNextNumber -Path $z } | Should -Throw '*Zähler nicht lesbar*'
    }
}

Describe 'Datensaetze mit Zwischenspeicher' {
    BeforeEach {
        $script:D = New-TestDir
        $script:F = Join-Path $script:D 'fristen'
        Set-TestFile (Join-Path $script:F 'f-00000001.json') '{ "id": "f-00000001", "titel": "Eins" }'
        Set-TestFile (Join-Path $script:F 'f-00000002.json') '{ "id": "f-00000002", "titel": "Zwei" }'
        Set-TestFile (Join-Path $script:F '.~f-00000003.json.abc.tmp') 'halb'
        Set-TestFile (Join-Path $script:F 'notiz.json') '{ "x": 1 }'
        $script:Conv = { param($Data, $File) if ("$($Data.id).json" -ne $File.Name) { throw "enthält eine andere Kennung ($($Data.id))" }; [pscustomobject]@{ id = "$($Data.id)"; titel = "$($Data.titel)" } }
    }
    AfterEach { Remove-Item -LiteralPath $script:D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'liest, nutzt den Zwischenspeicher und liest nur Geaendertes neu' {
        $c = @{}
        $r = Get-HUDataRecords -Folder $script:F -Cache $c -Convert $script:Conv -NamePattern '^f-[a-z0-9]{8}$'
        @($r.Items).Count | Should -Be 2
        $r.Gelesen | Should -Be 2; $r.Neu | Should -Be 2
        @($r.Items | ForEach-Object titel) -join ',' | Should -Be 'Eins,Zwei'
        $r2 = Get-HUDataRecords -Folder $script:F -Cache $c -Convert $script:Conv -NamePattern '^f-[a-z0-9]{8}$'
        $r2.Gelesen | Should -Be 0
        Set-TestFile (Join-Path $script:F 'f-00000002.json') '{ "id": "f-00000002", "titel": "Zwei geändert" }'
        (Get-Item (Join-Path $script:F 'f-00000002.json')).LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddMinutes(1)
        $r3 = Get-HUDataRecords -Folder $script:F -Cache $c -Convert $script:Conv -NamePattern '^f-[a-z0-9]{8}$'
        $r3.Gelesen | Should -Be 1
        @($r3.Items | Where-Object id -eq 'f-00000002')[0].titel | Should -Be 'Zwei geändert'
    }
    It 'verschwundene Datei faellt raus, kaputte behaelt den alten Stand mit Fehler' {
        $c = @{}
        [void](Get-HUDataRecords -Folder $script:F -Cache $c -Convert $script:Conv -NamePattern '^f-')
        Remove-Item -LiteralPath (Join-Path $script:F 'f-00000001.json')
        Set-TestFile (Join-Path $script:F 'f-00000002.json') '{ kaputt und länger als vorher'
        $r = Get-HUDataRecords -Folder $script:F -Cache $c -Convert $script:Conv -NamePattern '^f-'
        $r.Weg | Should -Be 1
        @($r.Items).Count | Should -Be 1
        $r.Items[0].titel | Should -Be 'Zwei'
        @($r.Errors).Count | Should -Be 1
        $c.Count | Should -Be 1
    }
    It 'Convert wirft oder liefert $null: Fehler, Datensatz fehlt' {
        Set-TestFile (Join-Path $script:F 'f-00000009.json') '{ "id": "f-99999999" }'
        $r = Get-HUDataRecords -Folder $script:F -Convert $script:Conv -NamePattern '^f-'
        @($r.Items).Count | Should -Be 2
        ($r.Errors -join ' ') | Should -Match 'andere Kennung'
        $r2 = Get-HUDataRecords -Folder $script:F -Convert { param($d, $f) $null } -NamePattern '^f-'
        @($r2.Items).Count | Should -Be 0
        @($r2.Errors).Count | Should -Be 3
    }
    It 'mehrere Ordner, ohne Convert, -Immutable liest nie neu, fehlender Ordner ist kein Fehler' {
        $g = Join-Path $script:D 'ausgaben'
        Set-TestFile (Join-Path $g 'a1.json') '{ "id": "a1" }'
        $c = @{}
        $r = Get-HUDataRecords -Folder @($script:F, $g, (Join-Path $script:D 'fehlt')) -Cache $c -Immutable
        @($r.Items).Count | Should -Be 4
        @($r.Errors).Count | Should -Be 0
        Set-TestFile (Join-Path $g 'a1.json') '{ "id": "a1", "neu": true, "mehr": "Text" }'
        $r2 = Get-HUDataRecords -Folder @($script:F, $g) -Cache $c -Immutable
        $r2.Gelesen | Should -Be 0
        $r.Cache | Should -Be $c
    }
    It 'Zwischenspeicher eines anderen Ordners bleibt erhalten' {
        $g = Join-Path $script:D 'jahr2027'
        Set-TestFile (Join-Path $g 'b.json') '{ "id": "b" }'
        $c = @{}
        [void](Get-HUDataRecords -Folder $g -Cache $c)
        [void](Get-HUDataRecords -Folder $script:F -Cache $c)
        $c.Count | Should -Be 4   # b.json + f-1 + f-2 + notiz.json
    }
}

Describe 'Zahlen und Euro' {
    It 'ConvertTo-HUDecimal liest oesterreichische und technische Schreibweisen' {
        ConvertTo-HUDecimal "1.234,56 $script:Euro" | Should -Be ([decimal]1234.56)
        ConvertTo-HUDecimal '12,5' | Should -Be ([decimal]12.5)
        ConvertTo-HUDecimal '1234.56' | Should -Be ([decimal]1234.56)
        ConvertTo-HUDecimal "$script:Euro -3,10" | Should -Be ([decimal]-3.10)
        ConvertTo-HUDecimal ('1' + [char]0x00A0 + '234,00') | Should -Be ([decimal]1234)
        ConvertTo-HUDecimal 7 | Should -Be ([decimal]7)
        ConvertTo-HUDecimal ([double]1.5) | Should -Be ([decimal]1.5)
        ConvertTo-HUDecimal '' | Should -Be ([decimal]0)
        ConvertTo-HUDecimal 'abc' -Default 9 | Should -Be ([decimal]9)
        ConvertTo-HUDecimal $null -Default 2 | Should -Be ([decimal]2)
        (ConvertTo-HUDecimal '5') | Should -BeOfType [decimal]
    }
    It 'Test-HUNumberText' {
        Test-HUNumberText '' | Should -BeTrue
        Test-HUNumberText '1.234,5' | Should -BeTrue
        Test-HUNumberText 'zwölf' | Should -BeFalse
    }
    It 'Anzeige einheitlich mit Tausenderpunkt (gleich unter 5.1 und 7)' {
        Format-HUEuro 1234.5 | Should -Be "1.234,50 $script:Euro"
        Format-HUEuro -1234567.891 | Should -Be "-1.234.567,89 $script:Euro"
        Format-HUEuro 1234.5 -NoGrouping | Should -Be "1234,50 $script:Euro"
        Format-HUEuro 0 | Should -Be "0,00 $script:Euro"
        Format-HUNumber 1234.5 | Should -Be '1.234,5'
        Format-HUNumber 1234.5 -Fixed | Should -Be '1.234,50'
        Format-HUNumber 1234.567 -Digits 0 | Should -Be '1.235'
        Format-HUNumber 1234.5 -NoGrouping | Should -Be '1234,5'
        Format-HUKm 1234.56 | Should -Be '1.234,6 km'
        Format-HUKm 12 | Should -Be '12 km'
        Format-HUKm 1234.5 -NoGrouping | Should -Be '1234,5 km'
    }
    It 'kaufmaennisch runden' {
        Get-HURound 2.345 | Should -Be ([decimal]2.35)
        Get-HURound 2.5 -Digits 0 | Should -Be ([decimal]3)
        Get-HURound -2.5 -Digits 0 | Should -Be ([decimal]-3)
        Get-HURound 0.125 | Should -Be ([decimal]0.13)
    }
}

Describe 'Datum und Uhrzeit' {
    It 'liest Eingaben, behaelt die Uhrzeit, -DateOnly schneidet ab' {
        ConvertFrom-HUDateText '6.11.2026' | Should -Be ([datetime]::new(2026, 11, 6))
        ConvertFrom-HUDateText '06.11.26' | Should -Be ([datetime]::new(2026, 11, 6))
        ConvertFrom-HUDateText '2026-11-06' | Should -Be ([datetime]::new(2026, 11, 6))
        ConvertFrom-HUDateText '2026-11-06T14:30' | Should -Be ([datetime]::new(2026, 11, 6, 14, 30, 0))
        ConvertFrom-HUDateText '6.11.2026 7:05' | Should -Be ([datetime]::new(2026, 11, 6, 7, 5, 0))
        ConvertFrom-HUDateText '2026-11-06T14:30:15' -DateOnly | Should -Be ([datetime]::new(2026, 11, 6))
        ConvertFrom-HUDateText '2026-11-06T14:30:15.1234567' | Should -Be ([datetime]::new(2026, 11, 6, 14, 30, 15).AddTicks(1234567))
        $dt = [datetime]::new(2026, 3, 4, 5, 6, 7)
        ConvertFrom-HUDateText $dt | Should -Be $dt
        ConvertFrom-HUDateText $dt -DateOnly | Should -Be ([datetime]::new(2026, 3, 4))
    }
    It 'Excel-Seriennummer, ungueltige Eingaben' {
        $serial = [datetime]::new(2026, 11, 6).ToOADate()
        ConvertFrom-HUDateText "$serial" | Should -Be ([datetime]::new(2026, 11, 6))
        ConvertFrom-HUDateText ("{0}.5" -f $serial) | Should -Be ([datetime]::new(2026, 11, 6, 12, 0, 0))
        ConvertFrom-HUDateText '31.02.2026' | Should -BeNullOrEmpty
        ConvertFrom-HUDateText '2026' | Should -BeNullOrEmpty
        ConvertFrom-HUDateText '' | Should -BeNullOrEmpty
        ConvertFrom-HUDateText 'morgen' | Should -BeNullOrEmpty
    }
    It 'Text fuer Dateien und Anzeige' {
        ConvertTo-HUDateText '6.11.2026' | Should -Be '2026-11-06'
        ConvertTo-HUDateText ([datetime]::new(2026, 1, 2, 3, 4, 5)) | Should -Be '2026-01-02'
        ConvertTo-HUDateText 'x' | Should -Be ''
        Format-HUDate '2026-11-06' | Should -Be '06.11.2026'
        Format-HUDate 'frei' | Should -Be 'frei'
        Format-HUDate 'frei' -EmptyIfInvalid | Should -Be ''
        Format-HUTime '2026-11-06T14:30:00' | Should -Be '06.11.2026 14:30'
        Format-HUTime ([datetime]::new(2026, 11, 6, 9, 1, 0)) | Should -Be '06.11.2026 09:01'
        Format-HUTime 'frei' | Should -Be 'frei'
        Format-HUTime $null | Should -Be ''
    }
    It 'Uhrzeit lesen' {
        ConvertFrom-HUClockText '7:30' | Should -Be ([TimeSpan]::new(7, 30, 0))
        ConvertFrom-HUClockText '07:30' | Should -Be ([TimeSpan]::new(7, 30, 0))
        ConvertFrom-HUClockText '730' | Should -Be ([TimeSpan]::new(7, 30, 0))
        ConvertFrom-HUClockText '1945' | Should -Be ([TimeSpan]::new(19, 45, 0))
        ConvertFrom-HUClockText '7.30' | Should -Be ([TimeSpan]::new(7, 30, 0))
        ConvertFrom-HUClockText '7,30' | Should -Be ([TimeSpan]::new(7, 30, 0))
        ConvertFrom-HUClockText '7' | Should -Be ([TimeSpan]::new(7, 0, 0))
        ConvertFrom-HUClockText '24:00' | Should -BeNullOrEmpty
        ConvertFrom-HUClockText '7:60' | Should -BeNullOrEmpty
        ConvertFrom-HUClockText '' | Should -BeNullOrEmpty
        ConvertFrom-HUClockText 'acht' | Should -BeNullOrEmpty
    }
}

Describe 'Verlauf' {
    It 'Eintrag in fester Reihenfolge, mit Feldern und Zusatzwerten' {
        $f = @([pscustomobject]@{ feld = 'Titel'; alt = 'A'; neu = 'B' })
        $e = New-HUHistoryEntry -Action 'geaendert' -Text 'Titel' -Fields $f -Extra @{ faellig = '2026-12-01' } -UserName 'u1' -PcName 'pc1' -Now ([datetime]::new(2026, 11, 7, 10, 0, 0))
        @($e.PSObject.Properties.Name) -join ',' | Should -Be 'zeit,benutzer,pc,aktion,text,felder,faellig'
        $e.zeit | Should -Be '2026-11-07T10:00:00'
        $e.benutzer | Should -Be 'u1'; $e.pc | Should -Be 'pc1'; $e.aktion | Should -Be 'geaendert'
        $e.felder[0].neu | Should -Be 'B'
        $e2 = New-HUHistoryEntry -Action 'angelegt' -UserName 'u1' -PcName 'pc1'
        $e2.PSObject.Properties['felder'] | Should -BeNullOrEmpty
    }
    It 'an Datensatz haengen: Objekt ohne und mit Verlauf, Hashtable' {
        $r = [pscustomobject]@{ id = 'x' }
        [void](Add-HUHistoryEntry -Record $r -Action 'angelegt' -UserName 'u1' -PcName 'pc1')
        @($r.verlauf).Count | Should -Be 1
        [void](Add-HUHistoryEntry -Record $r -Action 'geaendert' -Text 'Status' -UserName 'u2' -PcName 'pc2')
        @($r.verlauf).Count | Should -Be 2
        $r.verlauf[1].benutzer | Should -Be 'u2'
        $h = @{ id = 'y' }
        [void](Add-HUHistoryEntry -Record $h -Action 'angelegt' -Property 'log')
        @($h.log).Count | Should -Be 1
    }
    It 'liest alte Felder (von, art) und behaelt Zusatzwerte' {
        $old = ConvertFrom-HUJsonText '{ "zeit": "2026-11-01T08:00:00", "von": "anna", "art": "erledigt", "text": "Erledigt", "erledigtAm": "2026-11-01", "felder": [ { "feld": "F", "alt": "1", "neu": "2" } ] }'
        $e = ConvertTo-HUHistoryEntry $old
        $e.benutzer | Should -Be 'anna'
        $e.aktion | Should -Be 'erledigt'
        $e.pc | Should -Be ''
        $e.erledigtAm | Should -Be '2026-11-01'
        $e.felder[0].alt | Should -Be '1'
        $e.PSObject.Properties['von'] | Should -BeNullOrEmpty
        $new = ConvertTo-HUHistoryEntry ([pscustomobject]@{ zeit = '2026-11-02T08:00:00'; benutzer = 'bert'; pc = 'p'; aktion = 'neu'; text = '' })
        $new.benutzer | Should -Be 'bert'; $new.aktion | Should -Be 'neu'
    }
    It 'Feldvergleich alt/neu und Aenderungstext' {
        $a = [pscustomobject]@{ titel = 'Alt'; teile = @('a', 'b'); aktiv = $true; notiz = 'gleich' }
        $b = [pscustomobject]@{ titel = 'Neu'; teile = @('a'); aktiv = $false; notiz = 'gleich' }
        $c = Compare-HURecordFields -Old $a -New $b -Fields ([ordered]@{ titel = 'Titel'; teile = 'Teilräume'; aktiv = 'Aktiv'; notiz = 'Notiz' })
        @($c).Count | Should -Be 3
        $c[0].feld | Should -Be 'Titel'; $c[0].alt | Should -Be 'Alt'; $c[0].neu | Should -Be 'Neu'
        $c[1].alt | Should -Be 'a, b'
        $c[2].neu | Should -Be 'nein'
        $f = Compare-HURecordFields -Old $a -New $b -Fields @('titel') -Format { param($k, $v) "[$v]" }
        $f[0].alt | Should -Be '[Alt]'
        @(Compare-HURecordFields -Old $null -New $b -Fields @('titel'))[0].alt | Should -Be ''
        Get-HUFieldChangeText $c[0..1] | Should -Be ('Titel: Alt ' + [char]0x2192 + ' Neu; Teilräume: a, b ' + [char]0x2192 + ' a')
    }
}

Describe 'Anhaenge' {
    BeforeEach {
        $script:D = New-TestDir
        $script:Src = Join-Path $script:D 'quelle'
        [void](New-Item -ItemType Directory -Path $script:Src)
        $script:Pdf = Join-Path $script:Src 'Prüf bericht.PDF'
        [IO.File]::WriteAllBytes($script:Pdf, [byte[]](37, 80, 68, 70, 45))
        $script:Ziel = Join-Path $script:D 'anhaenge\f-00000001'
    }
    AfterEach { Remove-Item -LiteralPath $script:D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'prueft Endung, Groesse, fehlende Datei und ausgeschaltete Anhaenge' {
        (Test-HUAttachment -Path $script:Pdf -Extensions 'pdf').Ok | Should -BeTrue
        (Test-HUAttachment -Path $script:Pdf -Extensions '.jpg', '.png').Message | Should -Match 'JPG, PNG'
        (Test-HUAttachment -Path (Join-Path $script:Src 'nix.pdf')).Ok | Should -BeFalse
        (Test-HUAttachment -Path $script:Pdf -MaxMB 0).Message | Should -Be 'Anhänge sind ausgeschaltet.'
        $big = Join-Path $script:Src 'gross.pdf'
        $fs = [IO.File]::Create($big); $fs.SetLength(1MB + 1); $fs.Dispose()
        (Test-HUAttachment -Path $big -MaxMB 1).Message | Should -Match 'erlaubt sind 1 MB'
        { Add-HUAttachment -Folder $script:Ziel -Path $big -MaxMB 1 } | Should -Throw
    }
    It 'mit Zeitstempel: nie ueberschreiben, Anzeigename ohne Stempel' {
        $now = [datetime]::new(2026, 11, 7, 14, 30, 0)
        $a = Add-HUAttachment -Folder $script:Ziel -Path $script:Pdf -Extensions '.pdf' -TimePrefix -Now $now
        $b = Add-HUAttachment -Folder $script:Ziel -Path $script:Pdf -Extensions '.pdf' -TimePrefix -Now $now
        $a.Name | Should -Be '20261107-143000_Prüf_bericht.pdf'
        $b.Name | Should -Be '20261107-143000_2_Prüf_bericht.pdf'
        Test-Path -LiteralPath $b.Path | Should -BeTrue
        Get-HUAttachmentDisplayName $b.Name | Should -Be 'Prüf_bericht.pdf'
        [IO.File]::ReadAllBytes($a.Path).Length | Should -Be 5
    }
    It 'ohne Zeitstempel: Name, dann _2; auflisten mit Filter' {
        $a = Add-HUAttachment -Folder $script:Ziel -Path $script:Pdf
        $b = Add-HUAttachment -Folder $script:Ziel -Path $script:Pdf
        $a.Name | Should -Be 'Prüf_bericht.pdf'
        $b.Name | Should -Be 'Prüf_bericht_2.pdf'
        Set-TestFile (Join-Path $script:Ziel 'notiz.txt') 'x'
        @(Get-HUAttachments $script:Ziel -Extensions '.pdf').Count | Should -Be 2
        @(Get-HUAttachments $script:Ziel).Count | Should -Be 3
        @(Get-HUAttachments (Join-Path $script:D 'fehlt')).Count | Should -Be 0
    }
    It 'Geraetenamen von Windows werden entschaerft (interne Hilfe, eine Quelldatei con.pdf laesst Windows nicht zu)' {
        InModuleScope HU.Data {
            Get-HUDataFileBase 'con' | Should -Be '_con'
            Get-HUDataFileBase 'LPT1.alt' | Should -Be '_LPT1.alt'
            Get-HUDataFileBase 'a\b:c' | Should -Be 'a_b_c'
            Get-HUDataFileBase '...' | Should -Be 'unbekannt'
        }
    }
    It 'entfernen = nach _entfernt verschieben, zweites Mal mit Zeitstempel, fremder Pfad wirft' {
        $a = Add-HUAttachment -Folder $script:Ziel -Path $script:Pdf
        $now = [datetime]::new(2026, 11, 7, 15, 0, 0)
        $m = Remove-HUAttachment -Folder $script:Ziel -Name $a.Name -Now $now
        $m | Should -Be (Join-Path $script:Ziel ('_entfernt\' + $a.Name))
        Test-Path -LiteralPath $a.Path | Should -BeFalse
        @(Get-HUAttachments $script:Ziel).Count | Should -Be 0
        $a2 = Add-HUAttachment -Folder $script:Ziel -Path $script:Pdf
        $m2 = Remove-HUAttachment -Folder $script:Ziel -Name $a2.Path -Now $now
        Split-Path $m2 -Leaf | Should -Be ('20261107-150000_' + $a2.Name)
        Remove-HUAttachment -Folder $script:Ziel -Name 'gibtsnicht.pdf' | Should -Be ''
        { Remove-HUAttachment -Folder $script:Ziel -Name $script:Pdf } | Should -Throw
        $bin = Join-Path $script:D 'papierkorb'
        $a3 = Add-HUAttachment -Folder $script:Ziel -Path $script:Pdf
        Split-Path (Remove-HUAttachment -Folder $script:Ziel -Name $a3.Name -BinFolder $bin) -Parent | Should -Be $bin
    }
}

Describe 'Gemeinsame Einstellungen und Berechtigte' {
    BeforeEach { $script:D = New-TestDir; $script:S = Join-Path $script:D 'einstellungen.json' }
    AfterEach { Remove-Item -LiteralPath $script:D -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Vorgaben mischen, auch verschachtelt; unbekannte Schluessel nur mit -KeepExtra' {
        $def = [ordered]@{ tage = 30; outlook = [ordered]@{ art = 'aufgabe'; uhrzeit = '08:00' }; liste = @() }
        $data = ConvertFrom-HUJsonText '{ "tage": 7, "outlook": { "uhrzeit": "09:30" }, "fremd": 1 }'
        $m = Merge-HUSettings -Defaults $def -Data $data
        $m.tage | Should -Be 7
        $m.outlook.art | Should -Be 'aufgabe'
        $m.outlook.uhrzeit | Should -Be '09:30'
        $m.PSObject.Properties['fremd'] | Should -BeNullOrEmpty
        (Merge-HUSettings -Defaults $def -Data $data -KeepExtra).fremd | Should -Be 1
        $def.outlook.uhrzeit | Should -Be '08:00'
    }
    It 'lesen ohne Datei: Vorgaben, Stand leer; speichern setzt geaendert/geaendertVon; Stand wird geprueft' {
        $def = @{ tage = 30; berechtigte = @() }
        $r = Read-HUSettings -Path $script:S -Defaults $def
        $r.Stamp | Should -Be ''; $r.Error | Should -Be ''
        $r.Settings.tage | Should -Be 30
        $r.Settings.geaendertVon | Should -Be ''
        $r.Settings.tage = 14
        $r.Settings.berechtigte = @('anna')
        $s = Save-HUSettings -Path $script:S -Settings $r.Settings -Stamp $r.Stamp -UserName 'verwalter'
        $s.Ok | Should -BeTrue
        $r.Settings.geaendertVon | Should -Be 'verwalter'
        $r2 = Read-HUSettings -Path $script:S -Defaults $def
        $r2.Settings.tage | Should -Be 14
        @($r2.Settings.berechtigte) | Should -Be @('anna')
        $r2.Settings.geaendertVon | Should -Be 'verwalter'
        $r2.Settings.geaendert | Should -BeOfType [string]
        $r2.Settings.geaendert | Should -Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$'
        $r2.Stamp | Should -Be $s.Stamp
        $old = Save-HUSettings -Path $script:S -Settings $r.Settings -Stamp $r.Stamp -UserName 'anna'
        $old.Ok | Should -BeFalse
        $old.Message | Should -Match 'Einstellungen'
    }
    It 'Hashtable als Einstellungen' {
        $h = [ordered]@{ a = 1 }
        (Save-HUSettings -Path $script:S -Settings $h -UserName 'u').Ok | Should -BeTrue
        $h.geaendertVon | Should -Be 'u'
        (Read-HUJson $script:S).Data.a | Should -Be 1
    }
    It 'Test-HUPermittedUser: Verwalter, Name, DOMAENE\Name, sonst nein' {
        Test-HUPermittedUser @() 'anna' -IsVerwalter | Should -BeTrue
        Test-HUPermittedUser @('Anna', 'bert') 'anna' | Should -BeTrue
        Test-HUPermittedUser @('GEMEINDE\anna') 'anna' -Domain 'gemeinde' | Should -BeTrue
        Test-HUPermittedUser @('GEMEINDE\anna') 'anna' -Domain 'ANDERE' | Should -BeFalse
        Test-HUPermittedUser @('anna') 'GEMEINDE\anna' | Should -BeTrue
        Test-HUPermittedUser @('GEMEINDE\anna') 'GEMEINDE\Anna' -Domain '' | Should -BeTrue
        Test-HUPermittedUser @('bert', '', $null) 'anna' | Should -BeFalse
        Test-HUPermittedUser $null 'anna' | Should -BeFalse
        Test-HUPermittedUser @('anna') '' | Should -BeFalse
        Test-HUPermittedUser @('anna') 'anna' -IsVerwalter:$false | Should -BeTrue
    }
    It 'Anzeigename: fuer andere Benutzer der Name, fuer sich selbst nie leer' {
        Get-HUUserDisplayName 'jemand-anderer' | Should -Be 'jemand-anderer'
        Get-HUUserDisplayName | Should -Not -BeNullOrEmpty
    }
}
