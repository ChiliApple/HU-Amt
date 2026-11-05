#Requires -Version 5.1
# Pester 5: Modulsystem (Functions\Core-Addons.ps1) mit dem Beispielmodul und kuenstlichen Modulen unter TEMP.

BeforeAll {
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking
    . (Join-Path $script:Root 'Functions\Core-Update.ps1')
    . (Join-Path $script:Root 'Functions\Core-Addons.ps1')
    $script:Example = Join-Path $PSScriptRoot 'Beispiel'

    # kuenstliches Modul in einem Ordner anlegen
    function New-TestAddon([string]$Parent, [string]$Id, [hashtable]$Extra = @{}, [string]$Folder = '') {
        $d = Join-Path $Parent $(if ($Folder) { $Folder } else { $Id })
        [void](New-Item -ItemType Directory -Path $d -Force)
        $j = [ordered]@{ id = $Id; title = $Id; version = '1.0.0'; minApp = '0.0.1'; tab = [ordered]@{ header = $Id; xaml = 'Tab.xaml'; script = 'Addon.ps1'; init = "Initialize-HUAddon$Id" } }
        foreach ($k in $Extra.Keys) { $j[$k] = $Extra[$k] }
        [IO.File]::WriteAllText((Join-Path $d 'addon.json'), ($j | ConvertTo-Json -Depth 5), (New-Object Text.UTF8Encoding $false))
        return $d
    }
}

Describe 'addon.json lesen' {
    It 'Beispielmodul ist gueltig, mit Reihenfolge und Endungen' {
        $i = Read-HUAddonInfo $script:Example
        $i.Error | Should -Be ''
        $i.Id | Should -Be 'Beispiel'
        $i.Order | Should -Be 90
        $i.Init | Should -Be 'Initialize-HUAddonBeispiel'
        @($i.Drop) | Should -Be @('.pdf', '.xlsx', '.txt')
    }
    It 'Beispielmodul: Tab.xaml hat ein TabItem als Wurzel, Addon.ps1 hat Init und Drop' {
        ([xml][IO.File]::ReadAllText((Join-Path $script:Example 'Tab.xaml'), [Text.Encoding]::UTF8)).DocumentElement.LocalName | Should -Be 'TabItem'
        $t = [IO.File]::ReadAllText((Join-Path $script:Example 'Addon.ps1'), [Text.Encoding]::UTF8)
        $t | Should -Match 'function Initialize-HUAddonBeispiel'
        $t | Should -Match 'function Invoke-HUAddonDropBeispiel'
    }
    Context 'Fehlerfaelle' {
        BeforeEach {
            $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtAdd-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            [void](New-Item -ItemType Directory -Path $script:Tmp)
        }
        AfterEach { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }

        It 'reservierte Kennung wird abgelehnt' { (Read-HUAddonInfo (New-TestAddon $script:Tmp 'Benutzer')).Error | Should -Match 'Kennung' }
        It 'Kennung mit Zeichen ausserhalb A-Z/0-9 wird abgelehnt' { (Read-HUAddonInfo (New-TestAddon $script:Tmp 'a-b' -Folder 'x')).Error | Should -Match 'Kennung' }
        It 'falscher Init-Name wird abgelehnt' {
            $d = New-TestAddon $script:Tmp 'Xx' @{ tab = @{ header = 'x'; xaml = 'Tab.xaml'; script = 'Addon.ps1'; init = 'Start-Irgendwas' } }
            (Read-HUAddonInfo $d).Error | Should -Match 'init'
        }
        It 'Pfad mit .. wird abgelehnt' {
            $d = New-TestAddon $script:Tmp 'Yy' @{ tab = @{ header = 'x'; xaml = '..\Tab.xaml'; script = 'Addon.ps1'; init = 'Initialize-HUAddonYy' } }
            (Read-HUAddonInfo $d).Error | Should -Match 'xaml'
        }
        It 'ungueltige Endung bei drop wird abgelehnt' { (Read-HUAddonInfo (New-TestAddon $script:Tmp 'Zz' @{ drop = @('pdf') })).Error | Should -Match 'drop' }
        It 'Ordnername muss zur Kennung passen' {
            [void](New-TestAddon $script:Tmp 'Aa' -Folder 'Bb')
            (@(Get-HUAddonsInDir $script:Tmp ([Version]'0.0.1') 'Test'))[0].Error | Should -Match 'Ordnername'
        }
        It 'zu neue Mindestversion wird abgelehnt' {
            [void](New-TestAddon $script:Tmp 'Cc' @{ minApp = '9.0.0' })
            (@(Get-HUAddonsInDir $script:Tmp ([Version]'0.0.1') 'Test'))[0].Error | Should -Match 'Programmversion'
        }
        It 'privates Modul ohne Signatur wird nicht geladen' {
            [void](New-TestAddon $script:Tmp 'Dd')
            (@(Get-HUAddonsInDir $script:Tmp ([Version]'0.0.1') 'Privat' -Signed -Thumbprint ('A' * 40)))[0].Error | Should -Match 'nicht vertrauenswuerdig'
        }
        It 'Ordner mit _ am Anfang ist abgeschaltet' {
            [void](New-TestAddon $script:Tmp 'Ee' -Folder '_Ee')
            @(Get-HUAddonsInDir $script:Tmp ([Version]'0.0.1') 'Entwicklung' -SkipDisabled).Count | Should -Be 0
        }
    }
}

Describe 'Reihenfolge und Ersetzen' {
    It 'sortiert nach order, dann Titel; spaetere Quelle ersetzt gleiche Kennung' {
        $mk = { param($id, $order, $src) [pscustomobject]@{ Id = $id; Title = $id; Order = $order; Source = $src; Error = '' } }
        $list = Merge-HUAddonList @((& $mk 'Ocr' 20 'Programm'), (& $mk 'Energie' 10 'Privat'), (& $mk 'Ocr' 20 'Entwicklung'), (& $mk 'Abc' 20 'Programm'))
        @($list.Id) | Should -Be @('Energie', 'Abc', 'Ocr')
        ($list | Where-Object Id -eq 'Ocr').Source | Should -Be 'Entwicklung'
    }
    It 'fehlerhafte Eintraege fallen weg' {
        @(Merge-HUAddonList @([pscustomobject]@{ Id = 'X'; Title = 'X'; Order = 1; Error = 'kaputt' })).Count | Should -Be 0
    }
}

Describe 'Dateien hineinziehen' {
    It 'Endung passt unabhaengig von Gross-/Kleinschreibung' {
        $i = Read-HUAddonInfo $script:Example
        Test-HUAddonDropMatch $i 'C:\x\Rechnung.PDF' | Should -BeTrue
        Test-HUAddonDropMatch $i 'C:\x\Bild.png' | Should -BeFalse
    }
    It 'Modul ohne drop nimmt nichts an' {
        Test-HUAddonDropMatch ([pscustomobject]@{ Drop = @() }) 'a.pdf' | Should -BeFalse
    }
}

Describe 'Datenordner der Module' {
    It 'liegt in Data\<Id>' {
        $script:DataRoot = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtData-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        try {
            $d = Get-HUAddonDataDir 'Beispiel'
            $d | Should -Be (Join-Path $script:DataRoot 'Beispiel')
            Test-Path $d | Should -BeTrue
        } finally { Remove-Item -LiteralPath $script:DataRoot -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Modul-Datei (*.humodule)' {
    BeforeEach {
        $script:KTmp = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtKey-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:KTmp)
        function Write-KeyFile([string]$Json) { $p = Join-Path $script:KTmp 'test.humodule'; [IO.File]::WriteAllText($p, $Json, (New-Object Text.UTF8Encoding $false)); return $p }
    }
    AfterEach { Remove-Item -LiteralPath $script:KTmp -Recurse -Force -ErrorAction SilentlyContinue }
    It 'erkennt die Endung unabhaengig von Gross- und Kleinschreibung' {
        Test-HUAddonKeyFile 'C:\x\a.humodule' | Should -BeTrue
        Test-HUAddonKeyFile 'C:\x\a.HUMODULE' | Should -BeTrue
        Test-HUAddonKeyFile 'C:\x\a.json' | Should -BeFalse
    }
    It 'liest mehrere Eintraege mit eigenem Token' {
        $r = @(Read-HUAddonKeyFile (Write-KeyFile '{"module":[{"repo":"Besitzer/ModulA","token":"tok-a"},{"repo":"Besitzer/ModulB","token":"tok-b"}]}'))
        $r.Count | Should -Be 2
        $r[0].Owner | Should -Be 'Besitzer'; $r[0].Repo | Should -Be 'ModulA'; $r[0].Token | Should -Be 'tok-a'; $r[0].Error | Should -Be ''
        $r[1].Token | Should -Be 'tok-b'
    }
    It 'gemeinsamer Token gilt fuer Eintraege ohne eigenen' {
        $r = @(Read-HUAddonKeyFile (Write-KeyFile '{"token":"tok-x","module":[{"repo":"Besitzer/ModulA"},{"repo":"Besitzer/ModulB","token":"tok-b"}]}'))
        $r[0].Token | Should -Be 'tok-x'
        $r[1].Token | Should -Be 'tok-b'
    }
    It 'meldet ungueltiges Repository und fehlenden Token je Eintrag' {
        $r = @(Read-HUAddonKeyFile (Write-KeyFile '{"module":[{"repo":"ohneSchraegstrich","token":"t"},{"repo":"Besitzer/ModulA"}]}'))
        $r[0].Error | Should -Match 'ungültig'
        $r[1].Error | Should -Be 'Token fehlt'
    }
    It 'wirft bei kaputtem JSON oder fehlender Liste' {
        { Read-HUAddonKeyFile (Write-KeyFile '{kaputt') } | Should -Throw
        { Read-HUAddonKeyFile (Write-KeyFile '{"module":[]}') } | Should -Throw
    }
}
