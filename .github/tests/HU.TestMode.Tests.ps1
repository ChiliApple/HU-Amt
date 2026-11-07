#Requires -Version 5.1
# Pester 5: Testmodus (Functions\Core-TestMode.ps1). Kuenstlicher App-Ordner unter TEMP, keine Fenster.

BeforeAll {
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:Root 'Functions\Core-TestMode.ps1')
}

Describe 'Testbestand' {
    BeforeEach {
        $script:App = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTm-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path (Join-Path $script:App 'Data') -Force)
        [IO.File]::WriteAllText((Join-Path $script:App 'Data\Verwalter.txt'), "anna`r`n")
        [IO.File]::WriteAllText((Join-Path $script:App 'Data\echt.json'), '{}')
        $script:T = Join-Path $script:App 'Data\_Test'
    }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'legt den Testbestand mit Kennungsdatei an und uebernimmt die Verwalter-Liste' {
        Initialize-HUTestData -AppRoot $script:App | Should -BeFalse
        Test-Path (Join-Path $script:T '.hu-testbestand') | Should -BeTrue
        [IO.File]::ReadAllText((Join-Path $script:T 'Verwalter.txt')) | Should -Match 'anna'
    }
    It 'Zuruecksetzen loescht nur den Testbestand, echte Daten bleiben' {
        [void](Initialize-HUTestData -AppRoot $script:App)
        [void](New-Item -ItemType Directory -Path (Join-Path $script:T 'Buchung\buchungen') -Force)
        [IO.File]::WriteAllText((Join-Path $script:T 'Buchung\buchungen\1.json'), '{}')
        [IO.File]::WriteAllText((Join-Path $script:T '.zuruecksetzen'), 'x')
        Initialize-HUTestData -AppRoot $script:App | Should -BeTrue
        Test-Path (Join-Path $script:T 'Buchung') | Should -BeFalse
        Test-Path (Join-Path $script:T '.zuruecksetzen') | Should -BeFalse
        Test-Path (Join-Path $script:T '.hu-testbestand') | Should -BeTrue
        Test-Path (Join-Path $script:T 'Verwalter.txt') | Should -BeTrue
        Test-Path (Join-Path $script:App 'Data\echt.json') | Should -BeTrue
    }
    It 'Pfad-Pruefung: nur unter Data\_Test und nur mit Kennungsdatei' {
        Test-HUTestDataPath (Join-Path $script:T 'x') -AppRoot $script:App | Should -BeFalse
        [void](Initialize-HUTestData -AppRoot $script:App)
        Test-HUTestDataPath (Join-Path $script:T 'Buchung') -AppRoot $script:App | Should -BeTrue
        Test-HUTestDataPath $script:T -AppRoot $script:App | Should -BeTrue
        Test-HUTestDataPath (Join-Path $script:App 'Data') -AppRoot $script:App | Should -BeFalse
        Test-HUTestDataPath (Join-Path $script:App 'Data\_TestX') -AppRoot $script:App | Should -BeFalse
        Test-HUTestDataPath (Join-Path $script:T '..\echt.json') -AppRoot $script:App | Should -BeFalse
        Test-HUTestDataPath '' -AppRoot $script:App | Should -BeFalse
    }
    It 'scheitert die Pfad-Pruefung, wird nichts geloescht' {
        [void](Initialize-HUTestData -AppRoot $script:App)
        [IO.File]::WriteAllText((Join-Path $script:T 'wichtig.txt'), 'x')
        [IO.File]::WriteAllText((Join-Path $script:T '.zuruecksetzen'), 'x')
        Mock Test-HUTestDataPath { $false }
        { Initialize-HUTestData -AppRoot $script:App } | Should -Throw '*nicht gekennzeichnet*'
        Test-Path (Join-Path $script:T 'wichtig.txt') | Should -BeTrue
    }
}

Describe 'Testmodus erkennen' {
    AfterEach { Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue }
    It 'aus der Umgebungsvariable' {
        Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue
        Test-HUTestMode | Should -BeFalse
        $env:HUAMT_TESTMODE = '1'
        Test-HUTestMode | Should -BeTrue
    }
}

Describe 'Gefuehrter Test (Logik)' {
    BeforeAll {
        Import-Module (Join-Path $script:Root 'Core\HU.Exchange.psm1') -Force -DisableNameChecking
        . (Join-Path $script:Root 'Functions\UI-TestAssistant.ps1')
        $script:Sc = @(
            @{ Name = 'Saal'; Steps = @(
                    @{ Action = 'Saal 1/2 buchen'; Expected = 'gespeichert'; Check = { @{ Ok = $true; Text = '1 Buchung' } } }
                    @{ Action = 'Saal 2/3 zur selben Zeit'; Expected = 'abgelehnt'; Check = { throw 'Datei fehlt' } }
                    @{ Action = 'Preis ansehen'; Expected = 'Beispielwert' }
                ) }
            @{ Name = 'Museumshof'; Steps = @(@{ Action = 'buchen'; Expected = 'ok'; Check = { $false } }) }
        )
    }
    BeforeEach { $script:Run = New-HUTestRun -ModuleId 'Buchung' -ModuleVersion '0.1.1' -Scenarios $script:Sc -AppVersion '0.1.9' -UserName 'anna' -PcName 'PC-1' }

    It 'legt alle Schritte flach mit Nummer je Szenario an' {
        $script:Run.Steps.Count | Should -Be 4
        $script:Run.Steps[3].Szenario | Should -Be 'Museumshof'
        $script:Run.Steps[3].Nr | Should -Be 1
        @($script:Run.Steps | Where-Object { $_.Mensch -ne 'offen' }).Count | Should -Be 0
    }
    It 'automatische Pruefung setzt nie die Bestaetigung des Menschen' {
        $s = Invoke-HUTestCheck $script:Run.Steps[0]
        $s.Automatisch | Should -Be 'OK'
        $s.AutomatischText | Should -Be '1 Buchung'
        $s.Mensch | Should -Be 'offen'
    }
    It 'Fehler in der Pruefung zaehlt als Fehler, $false ebenso' {
        (Invoke-HUTestCheck $script:Run.Steps[1]).Automatisch | Should -Be 'Fehler'
        $script:Run.Steps[1].AutomatischText | Should -Match 'Datei fehlt'
        (Invoke-HUTestCheck $script:Run.Steps[3]).Automatisch | Should -Be 'Fehler'
    }
    It 'Schritt ohne Pruefung bleibt ungeprueft' {
        (Invoke-HUTestCheck $script:Run.Steps[2]).Automatisch | Should -Not -BeIn @('OK', 'Fehler')
    }
    It 'Bestaetigung des Menschen aendert das automatische Ergebnis nicht; Fehler braucht Kommentar' {
        [void](Invoke-HUTestCheck $script:Run.Steps[1])
        { Set-HUTestVerdict $script:Run.Steps[1] 'Fehler' ' ' } | Should -Throw
        [void](Set-HUTestVerdict $script:Run.Steps[1] 'Bestanden' 'sah richtig aus')
        $script:Run.Steps[1].Mensch | Should -Be 'Bestanden'
        $script:Run.Steps[1].Automatisch | Should -Be 'Fehler'
    }
    It 'Zusammenfassung trennt Mensch und Automatik und listet offene Fehler' {
        [void](Invoke-HUTestCheck $script:Run.Steps[0]); [void](Invoke-HUTestCheck $script:Run.Steps[1])
        [void](Set-HUTestVerdict $script:Run.Steps[0] 'Bestanden')
        [void](Set-HUTestVerdict $script:Run.Steps[2] 'Fehler' 'Preis fehlt')
        $sum = Get-HUTestSummary $script:Run
        $sum.Bestanden | Should -Be 1; $sum.FehlerMensch | Should -Be 1; $sum.Offen | Should -Be 2
        $sum.AutoOk | Should -Be 1; $sum.AutoFehler | Should -Be 1
        $sum.OffeneFehler.Count | Should -Be 2
    }
    It 'Protokoll: CSV und HTML mit getrennten Spalten fuer Automatik und Mensch' {
        $dir = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTp-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        try {
            [void](Invoke-HUTestCheck $script:Run.Steps[0]); [void](Set-HUTestVerdict $script:Run.Steps[2] 'Fehler' 'Preis <fehlt>')
            $r = Export-HUTestProtocol -Run $script:Run -Folder $dir -Now ([datetime]'2026-11-07 17:00')
            $csv = [IO.File]::ReadAllText($r.Csv, [Text.Encoding]::UTF8)
            $csv | Should -Match 'Automatisch geprüft;Automatisch Ist;Bestätigung Mensch;Ist laut Mensch / Kommentar'
            $csv | Should -Match '0\.1\.9;Buchung;0\.1\.1;Saal;1;'
            $html = [IO.File]::ReadAllText($r.Html, [Text.Encoding]::UTF8)
            $html | Should -Match 'Offene Fehler'
            $html | Should -Match 'Preis &lt;fehlt&gt;'
            $html | Should -Not -Match 'Preis <fehlt>'
            (Split-Path $r.Html -Leaf) | Should -Be 'Testprotokoll_Buchung_20261107-170000.html'
        } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    It 'Fenster-XAML ist gueltig' {
        if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { Set-ItResult -Skipped -Because 'XAML braucht STA'; return }
        Add-Type -AssemblyName PresentationFramework
        $src = [IO.File]::ReadAllText((Join-Path $script:Root 'Functions\UI-TestAssistant.ps1'), [Text.Encoding]::UTF8)
        $x = [regex]::Match($src, "(?s)@'\r?\n(<Window.*?</Window>)\r?\n'@").Groups[1].Value
        $theme = [IO.File]::ReadAllText((Join-Path $script:Root 'XAML\Theme.xaml'), [Text.Encoding]::UTF8)
        $x = $x.Replace('<!--HU:THEME-->', [regex]::Match($theme, '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>').Groups[1].Value)
        $w = [Windows.Markup.XamlReader]::Parse($x)
        $w.FindName('grdSteps') | Should -Not -BeNullOrEmpty
        $w.FindName('btnPass') | Should -Not -BeNullOrEmpty
    }
}

Describe 'Zeichen ausserhalb von [char]' {
    It 'kein [char] mit mehr als 4 Hex-Stellen (Emoji brauchen [char]::ConvertFromUtf32)' {
        $bad = @(Get-ChildItem -Path $script:Root -Recurse -File -Include *.ps1, *.psm1 | Where-Object { $_.FullName -notmatch '\\(Data|\.git)\\' } |
                Select-String -Pattern '\[char\]0x[0-9A-Fa-f]{5,}' | Where-Object { $_.Path -notlike '*HU.TestMode.Tests.ps1' } | ForEach-Object { "$($_.Filename):$($_.LineNumber)" })
        $bad | Should -BeNullOrEmpty
    }
}
