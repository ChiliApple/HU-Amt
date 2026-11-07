#Requires -Version 5.1
# Pester 5: Datei-Austausch (Core\HU.Exchange.psm1). Nur kuenstliche Daten in einem Testordner unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Exchange.psm1') -Force -DisableNameChecking
}

Describe 'Profil' {
    It 'hat die Vorgaben fuer Oesterreich' {
        $p = New-HUTableProfile
        $p.Delimiter | Should -Be ';'
        $p.DecimalSeparator | Should -Be ','
        $p.DateFormat | Should -Be 'dd.MM.yyyy'
        $p.Encoding | Should -Be 'utf8bom'
    }
    It 'lehnt ungueltige Angaben ab' {
        { New-HUTableProfile -Encoding 'latin9' } | Should -Throw
        { New-HUTableProfile -Delimiter ';;' } | Should -Throw
        { New-HUTableProfile -Delimiter ',' -DecimalSeparator ',' } | Should -Throw
        { New-HUTableProfile -Columns @(@{ Name = 'A'; Type = 'geld' }) } | Should -Throw
        { New-HUTableProfile -Columns @(@{ Name = 'A' }, @{ Name = 'A' }) } | Should -Throw
    }
}

Describe 'Tabellen' {
    BeforeEach {
        $script:Dir = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtEx-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Dir)
        $script:Rows = @(
            [pscustomobject]@{ Nr = 'R-1'; Text = 'Strom; Jänner "Netz"'; Betrag = [decimal]1234.5; Datum = [datetime]'2026-01-31' }
            [pscustomobject]@{ Nr = 'R-2'; Text = "Zwei`r`nZeilen"; Betrag = [decimal]-0.25; Datum = [datetime]'2026-02-01' }
        )
        $script:Cols = @(@{ Name = 'Belegnummer'; Source = 'Nr' }, @{ Name = 'Text' }, @{ Name = 'Betrag'; Type = 'number' }, @{ Name = 'Datum'; Type = 'date' })
    }
    AfterEach { Remove-Item -LiteralPath $script:Dir -Recurse -Force -ErrorAction SilentlyContinue }

    It 'schreibt Dezimalkomma, Datum und Anfuehrungszeichen wie vorgegeben' {
        $f = Export-HUTable -Rows $script:Rows -TableProfile (New-HUTableProfile -Columns $script:Cols) -Path (Join-Path $script:Dir 'a.csv')
        $t = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
        $t | Should -Match '^Belegnummer;Text;Betrag;Datum\r\n'
        $t | Should -Match 'R-1;"Strom; Jänner ""Netz""";1234,5;31\.01\.2026\r\n'
        $t | Should -Match ';-0,25;01\.02\.2026'
        $b = [IO.File]::ReadAllBytes($f)
        ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) | Should -BeTrue
    }
    It 'Windows-1252: ohne BOM, Umlaut als ein Byte, liest gleich zurueck' {
        $p = New-HUTableProfile -Encoding 'windows-1252' -Columns $script:Cols
        $f = Export-HUTable -Rows $script:Rows -TableProfile $p -Path (Join-Path $script:Dir 'b.csv')
        $b = [IO.File]::ReadAllBytes($f)
        $b[0] | Should -Be ([byte][char]'B')
        @($b | Where-Object { $_ -eq 0xE4 }).Count | Should -Be 1
        $r = @(Import-HUTable -Path $f -TableProfile $p)
        $r.Count | Should -Be 2
        $r[0].Nr | Should -Be 'R-1'
        $r[0].Text | Should -Be 'Strom; Jänner "Netz"'
        $r[0].Betrag | Should -Be ([decimal]1234.5)
        $r[0].Datum | Should -Be ([datetime]'2026-01-31')
        $r[1].Text | Should -Be "Zwei`r`nZeilen"
    }
    It 'meldet falsche Werte mit Zeile und Spalte' {
        $f = Join-Path $script:Dir 'c.csv'
        [IO.File]::WriteAllText($f, "Belegnummer;Text;Betrag;Datum`r`nR-1;x;abc;31.01.2026`r`n", (New-Object Text.UTF8Encoding $true))
        { Import-HUTable -Path $f -TableProfile (New-HUTableProfile -Columns $script:Cols) } | Should -Throw '*Zeile 2*Betrag*'
    }
    It 'meldet fehlende Spalten' {
        $f = Join-Path $script:Dir 'd.csv'
        [IO.File]::WriteAllText($f, "Belegnummer;Text`r`nR-1;x`r`n", (New-Object Text.UTF8Encoding $true))
        { Import-HUTable -Path $f -TableProfile (New-HUTableProfile -Columns $script:Cols) } | Should -Throw '*Betrag*'
    }
    It 'ueberschreibt nie und laesst keine Temp-Datei liegen' {
        $f = Join-Path $script:Dir 'e.csv'
        [void](Export-HUTable -Rows $script:Rows -TableProfile (New-HUTableProfile) -Path $f)
        { Export-HUTable -Rows $script:Rows -TableProfile (New-HUTableProfile) -Path $f } | Should -Throw
        @(Get-ChildItem -LiteralPath $script:Dir -Force -Filter '*.tmp').Count | Should -Be 0
    }
    It 'Profil als Datei speichern und lesen' {
        $pf = Join-Path $script:Dir 'Profile\Ziel.json'
        Save-HUTableProfile -TableProfile (New-HUTableProfile -Name 'Ziel' -Delimiter '\t' -Encoding 'windows-1252' -Columns $script:Cols) -Path $pf
        $p = Read-HUTableProfile $pf
        $p.Delimiter | Should -Be "`t"
        $p.Encoding | Should -Be 'windows-1252'
        @($p.Columns).Count | Should -Be 4
        (@($p.Columns) | Where-Object Name -eq 'Belegnummer').Source | Should -Be 'Nr'
        @(Get-HUTableProfiles (Split-Path $pf -Parent)).Count | Should -Be 1
    }
}

Describe 'Uebergabeordner' {
    BeforeEach {
        $script:Data = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtExD-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Data)
    }
    AfterEach { Remove-Item -LiteralPath $script:Data -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Standard ist der Ordner Ausgang des Moduls, einstellbar' {
        Get-HUHandoverFolder $script:Data 'Beispiel' | Should -Be (Join-Path $script:Data 'Beispiel\Ausgang')
        $other = Join-Path $script:Data 'Freigabe'
        Set-HUHandoverFolder $script:Data 'Beispiel' $other
        Get-HUHandoverFolder $script:Data 'Beispiel' | Should -Be $other
        Set-HUHandoverFolder $script:Data 'Beispiel' ''
        Get-HUHandoverFolder $script:Data 'Beispiel' | Should -Be (Join-Path $script:Data 'Beispiel\Ausgang')
    }
    It 'Dateiname mit Zeitstempel, bei Gleichstand fortlaufend' {
        $now = [datetime]'2026-11-04T07:30:15'
        $rows = @([pscustomobject]@{ A = 1 })
        $a = Export-HUHandoverTable -Rows $rows -TableProfile (New-HUTableProfile) -DataRoot $script:Data -ModuleId 'Beispiel' -BaseName 'Liste' -Now $now
        $b = Export-HUHandoverTable -Rows $rows -TableProfile (New-HUTableProfile) -DataRoot $script:Data -ModuleId 'Beispiel' -BaseName 'Liste' -Now $now
        Split-Path -Leaf $a | Should -Be 'Liste_2026-11-04_073015.csv'
        Split-Path -Leaf $b | Should -Be 'Liste_2026-11-04_073015_2.csv'
    }
    It 'lehnt reservierte und ungueltige Modul-Ids ab' {
        { Get-HUHandoverFolder $script:Data 'Benutzer' } | Should -Throw
        { Get-HUHandoverFolder $script:Data '..\x' } | Should -Throw
    }
}

Describe 'Werte aus Text (Pruefung 086, KERN-B1)' {
    BeforeAll { $script:Pr = New-HUTableProfile }
    It 'Zahl <In> wird <Out>' -TestCases @(
        @{ In = '1,5'; Out = '1,5' }
        @{ In = '12,50'; Out = '12,5' }
        @{ In = '1.234,56'; Out = '1234,56' }
        @{ In = '-1.234.567,8'; Out = '-1234567,8' }
        @{ In = ' 7 '; Out = '7' }
        @{ In = '1.5'; Out = '1,5' }
        @{ In = '1234.56'; Out = '1234,56' }
        @{ In = '1.234.567'; Out = '1234567' }
    ) {
        param($In, $Out)
        InModuleScope HU.Exchange -Parameters @{ V = $In; P = $script:Pr } { param($V, $P) Format-HUExValue $V 'number' $P } | Should -Be $Out
    }
    It 'Zahl <In> ist keine Zahl' -TestCases @(@{ In = '1,234.56' }, @{ In = '1.2,5' }, @{ In = 'abc' }, @{ In = '1,5,6' }) {
        param($In)
        { InModuleScope HU.Exchange -Parameters @{ V = $In; P = $script:Pr } { param($V, $P) Format-HUExValue $V 'number' $P } } | Should -Throw '*keine Zahl*'
    }
    It '[decimal] und [double] bleiben, wie sie sind' {
        InModuleScope HU.Exchange -Parameters @{ P = $script:Pr } { param($P) Format-HUExValue ([decimal]1234.5) 'number' $P } | Should -Be '1234,5'
        InModuleScope HU.Exchange -Parameters @{ P = $script:Pr } { param($P) Format-HUExValue 0.25 'number' $P } | Should -Be '0,25'
    }
    It 'Datum <In> wird <Out>' -TestCases @(
        @{ In = '05.11.2026'; Out = '05.11.2026' }
        @{ In = '13.11.2026'; Out = '13.11.2026' }
        @{ In = '5.1.2026'; Out = '05.01.2026' }
        @{ In = '05.11.2026 14:30'; Out = '05.11.2026' }
        @{ In = '2026-11-05'; Out = '05.11.2026' }
        @{ In = '2026-11-05T14:30:00'; Out = '05.11.2026' }
        @{ In = '2026-11-05 14:30'; Out = '05.11.2026' }
        @{ In = '2026-11-05T14:30:00.1234567'; Out = '05.11.2026' }
    ) {
        param($In, $Out)
        InModuleScope HU.Exchange -Parameters @{ V = $In; P = $script:Pr } { param($V, $P) Format-HUExValue $V 'date' $P } | Should -Be $Out
    }
    It 'Datum <In> ist kein Datum (kein freies Lesen)' -TestCases @(@{ In = '11/05/2026' }, @{ In = '32.01.2026' }, @{ In = 'morgen' }, @{ In = '05.13.2026' }) {
        param($In)
        { InModuleScope HU.Exchange -Parameters @{ V = $In; P = $script:Pr } { param($V, $P) Format-HUExValue $V 'date' $P } } | Should -Throw '*kein Datum*'
    }
    It '[datetime] bleibt' {
        InModuleScope HU.Exchange -Parameters @{ P = $script:Pr } { param($P) Format-HUExValue ([datetime]'2026-11-05') 'date' $P } | Should -Be '05.11.2026'
    }
    It 'Export ueber das Profil: 1,5 bleibt 1,5 und 05.11.2026 bleibt der 5. November' {
        $dir = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtB1-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        try {
            $p = New-HUTableProfile -Columns @(@{ Name = 'Betrag'; Type = 'number' }, @{ Name = 'Datum'; Type = 'date' })
            $f = Export-HUTable -Rows @([pscustomobject]@{ Betrag = '1,5'; Datum = '05.11.2026' }) -TableProfile $p -Path (Join-Path $dir 'b1.csv')
            ([IO.File]::ReadAllText($f, [Text.Encoding]::UTF8) -split "`r`n")[1] | Should -Be '1,5;05.11.2026'
        } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
