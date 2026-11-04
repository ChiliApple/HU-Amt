#Requires -Version 5.1
# Pester 5: Datei-Austausch (Core\HU.Exchange.psm1). Nur kuenstliche Daten in einem Testordner unter TEMP.

BeforeAll {
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
