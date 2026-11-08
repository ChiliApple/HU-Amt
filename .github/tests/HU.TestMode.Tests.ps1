#Requires -Version 5.1
# Pester 5: Testmodus (Functions\Core-TestMode.ps1). Kuenstlicher App-Ordner unter TEMP, keine Fenster.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking   # Get-HUDataLocation
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
    It 'eingestellter Datenordner: Testbestand und Verwalter-Liste kommen von dort' {
        $ext = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTmExt-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        try {
            [void](New-Item -ItemType Directory -Path $ext)
            [IO.File]::WriteAllText((Join-Path $ext 'Verwalter.txt'), "berta`r`n")
            Set-HUDataLocation -AppRoot $script:App -Path $ext
            Get-HUTestDataRoot -AppRoot $script:App | Should -Be (Join-Path $ext '_Test')
            [void](Initialize-HUTestData -AppRoot $script:App)
            [IO.File]::ReadAllText((Join-Path $ext '_Test\Verwalter.txt')) | Should -Match 'berta'
            Test-Path (Join-Path $script:T '.hu-testbestand') | Should -BeFalse
        } finally { Remove-Item -LiteralPath $ext -Recurse -Force -ErrorAction SilentlyContinue }
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

Describe 'Beispieldaten der Module (Testdaten)' {
    BeforeEach {
        $script:App = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTd-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $script:Mod = Join-Path $script:App 'AddonsDev\Muster'
        [void](New-Item -ItemType Directory -Path (Join-Path $script:Mod 'Testdaten\vorgaenge') -Force)
        [IO.File]::WriteAllText((Join-Path $script:Mod 'Testdaten\einstellungen.json'), '{"schema":1}')
        [IO.File]::WriteAllText((Join-Path $script:Mod 'Testdaten\vorgaenge\v1.json'), '{"id":"v1"}')
        [void](New-Item -ItemType Directory -Path (Join-Path $script:App 'Data') -Force)
        [IO.File]::WriteAllText((Join-Path $script:App 'Data\Verwalter.txt'), "anna`r`n")
        $script:A = [pscustomobject]@{ Id = 'Muster'; Dir = $script:Mod }
        $env:HUAMT_TESTMODE = '1'
        [void](Initialize-HUTestData -AppRoot $script:App)
    }
    AfterEach {
        Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue
    }
    It 'legt die Beispieldaten in den leeren Testbestand des Moduls' {
        Initialize-HUAddonTestData $script:A -AppRoot $script:App | Should -Be 2
        Test-Path (Join-Path $script:App 'Data\_Test\Muster\vorgaenge\v1.json') | Should -BeTrue
        Test-Path (Join-Path $script:App 'Data\Muster') | Should -BeFalse
    }
    It 'ueberschreibt nichts, wenn im Testbestand schon Daten liegen' {
        [void](New-Item -ItemType Directory -Path (Join-Path $script:App 'Data\_Test\Muster') -Force)
        [IO.File]::WriteAllText((Join-Path $script:App 'Data\_Test\Muster\eigen.json'), '{}')
        Initialize-HUAddonTestData $script:A -AppRoot $script:App | Should -Be 0
        Test-Path (Join-Path $script:App 'Data\_Test\Muster\einstellungen.json') | Should -BeFalse
    }
    It 'ausserhalb des Testmodus passiert nichts' {
        Remove-Item -LiteralPath Env:\HUAMT_TESTMODE
        Initialize-HUAddonTestData $script:A -AppRoot $script:App | Should -Be 0
        Test-Path (Join-Path $script:App 'Data\_Test\Muster') | Should -BeFalse
    }
    It 'Modul ohne Testdaten: nichts' {
        Initialize-HUAddonTestData ([pscustomobject]@{ Id = 'Leer'; Dir = (Join-Path $script:App 'gibt-es-nicht') }) -AppRoot $script:App | Should -Be 0
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
    It 'Oeffnen im StrictMode: keine ungesetzte Variable vor dem Fenster' {
        $env:HUAMT_TESTMODE = '1'
        try {
            . (Join-Path $script:Root 'Functions\UI-TestAssistant.ps1')   # frisch geladen wie beim App-Start
            function New-HUWindow { throw 'Fenster wuerde jetzt aufgehen' }
            function Get-HUThemedXaml([string]$Xaml) { $Xaml }
            { Show-HUTestAssistant -ModuleId 'Buchung' -Scenarios $script:Sc } | Should -Throw '*Fenster wuerde jetzt aufgehen*'
        } finally { Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue }
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

Describe 'Verwalter-Liste im Testmodus (Pruefung 086, KERN-B17)' {
    BeforeEach {
        $script:App = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTm17-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path (Join-Path $script:App 'Data') -Force)
        [IO.File]::WriteAllText((Join-Path $script:App 'Data\Verwalter.txt'), "anna`r`n")
        $script:T = Join-Path $script:App 'Data\_Test'
    }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'kopiert Verwalter.txt bei jedem Start neu' {
        [void](Initialize-HUTestData -AppRoot $script:App)
        [IO.File]::ReadAllText((Join-Path $script:T 'Verwalter.txt')) | Should -Match 'anna'
        [IO.File]::WriteAllText((Join-Path $script:App 'Data\Verwalter.txt'), "berta`r`n")
        [void](Initialize-HUTestData -AppRoot $script:App)
        $t = [IO.File]::ReadAllText((Join-Path $script:T 'Verwalter.txt'))
        $t | Should -Match 'berta'
        $t | Should -Not -Match 'anna'
    }
}

Describe 'Gefuehrter Test: Fenster (Pruefung 086, KERN-U1/U2)' {
    BeforeAll {
        Import-Module (Join-Path $script:Root 'Core\HU.Exchange.psm1') -Force -DisableNameChecking
        . (Join-Path $script:Root 'Functions\UI-TestAssistant.ps1')
        $script:TaSrc =[IO.File]::ReadAllText((Join-Path $script:Root 'Functions\UI-TestAssistant.ps1'), [Text.Encoding]::UTF8)
    }
    It 'Fenster hat einen GridSplitter in der 8-px-Spalte und Schliessen reagiert auf Esc' {
        if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { Set-ItResult -Skipped -Because 'XAML braucht STA'; return }
        Add-Type -AssemblyName PresentationFramework
        $x = [regex]::Match($script:TaSrc, "(?s)@'\r?\n(<Window.*?</Window>)\r?\n'@").Groups[1].Value
        $theme = [IO.File]::ReadAllText((Join-Path $script:Root 'XAML\Theme.xaml'), [Text.Encoding]::UTF8)
        $x = $x.Replace('<!--HU:THEME-->', [regex]::Match($theme, '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>').Groups[1].Value)
        $w = [Windows.Markup.XamlReader]::Parse($x)
        $sp = $w.FindName('taSplit')
        $sp | Should -BeOfType ([System.Windows.Controls.GridSplitter])
        [System.Windows.Controls.Grid]::GetColumn($sp) | Should -Be 1
        $sp.Parent.ColumnDefinitions[1].Width.Value | Should -Be 8
        $sp.ToolTip | Should -Not -BeNullOrEmpty
        $w.FindName('btnClose').IsCancel | Should -BeTrue
    }
    It 'merkt Groesse und Lage (Restore vor dem Anzeigen, Save beim Schliessen)' {
        $script:TaSrc | Should -Match "Restore-HUDialogState \`$w 'TestAssistent'"
        $script:TaSrc | Should -Match "Save-HUDialogState \`$t\.W 'TestAssistent'"
        $script:TaSrc.IndexOf("Restore-HUDialogState `$w 'TestAssistent'") | Should -BeLessThan $script:TaSrc.IndexOf('$w.Show()')
    }
    It 'Rueckfrage beim Schliessen nur, solange Bestaetigungen offen sind' {
        $sc = @(@{ Name = 'A'; Steps = @(@{ Action = 'x'; Expected = 'y' }, @{ Action = 'z'; Expected = 'w' }) })
        $run = New-HUTestRun -ModuleId 'Probe' -Scenarios $sc -AppVersion '0.1.12' -UserName 'anna' -PcName 'PC-1'
        Get-HUTestAssistantCloseQuestion $run | Should -Match '^2 Schritte sind noch nicht bestätigt'
        [void](Set-HUTestVerdict $run.Steps[0] 'Bestanden')
        Get-HUTestAssistantCloseQuestion $run | Should -Match '^1 Schritt ist noch nicht bestätigt'
        [void](Set-HUTestVerdict $run.Steps[1] 'Fehler' 'kaputt')
        Get-HUTestAssistantCloseQuestion $run | Should -Be ''
        Get-HUTestAssistantCloseQuestion $null | Should -Be ''
    }
    It 'Hinweis ausserhalb des Testmodus nutzt ein gueltiges Symbol' {
        Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue
        $script:Icons = New-Object System.Collections.Generic.List[string]
        function Show-HUMessage { param([string]$Text, [string]$Title, [ValidateSet('Info', 'Information', 'Warning', 'Error')][string]$Icon = 'Info', $Owner = $null) $script:Icons.Add($Icon) }
        Show-HUTestAssistant -ModuleId 'Probe' -Scenarios @(@{ Name = 'A'; Steps = @(@{ Action = 'x'; Expected = 'y' }) })
        $script:Icons | Should -Be @('Info')
    }
}
