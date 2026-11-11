#Requires -Version 5.1
# Pester 5: Modulsystem (Functions\Core-Addons.ps1) mit dem Beispielmodul und kuenstlichen Modulen unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
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

Describe 'Postfach zwischen Modulen' {
    BeforeEach {
        $script:Data = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        [void](New-Item -ItemType Directory -Path $script:Data -Force)
        $script:Src = Join-Path $TestDrive 'scan.pdf'
        [IO.File]::WriteAllText($script:Src, 'kuenstlich')
    }
    It 'kopiert die Datei und liefert den Eintrag' {
        [void](Add-HUAddonInboxItem -DataRoot $script:Data -To 'Ziel' -Path $script:Src -From 'OCR' -User 'Test' -Pc 'PC1')
        $e = @(Get-HUAddonInboxItems -DataRoot $script:Data -Id 'Ziel' -Extensions '.pdf')
        $e.Count | Should -Be 1
        $e[0].Name | Should -Be 'scan.pdf'
        $e[0].From | Should -Be 'OCR'
        $e[0].Pc | Should -Be 'PC1'
        $e[0].Path | Should -Not -Be $script:Src
        Test-Path -LiteralPath $e[0].Path | Should -BeTrue
        Test-Path -LiteralPath $script:Src | Should -BeTrue
    }
    It 'laesst fremde Endungen, unfertige und manipulierte Eintraege aus' {
        [void](Add-HUAddonInboxItem -DataRoot $script:Data -To 'Ziel' -Path $script:Src -From 'OCR')
        @(Get-HUAddonInboxItems -DataRoot $script:Data -Id 'Ziel' -Extensions '.xlsx').Count | Should -Be 0
        $root = Get-HUAddonInboxRoot $script:Data 'Ziel'
        # ohne eintrag.json
        $d1 = Join-Path $root ('a' * 32); [void](New-Item -ItemType Directory -Path $d1); [IO.File]::WriteAllText((Join-Path $d1 'x.pdf'), 'x')
        # Name zeigt aus dem Ordner hinaus
        $d2 = Join-Path $root ('b' * 32); [void](New-Item -ItemType Directory -Path $d2)
        [IO.File]::WriteAllText((Join-Path $d2 'eintrag.json'), '{"name":"..\..\boese.pdf","zeit":"2026-11-05T10:00:00"}')
        @(Get-HUAddonInboxItems -DataRoot $script:Data -Id 'Ziel' -Extensions '.pdf').Count | Should -Be 1
    }
    It 'entfernt einen Eintrag samt Kopie' {
        [void](Add-HUAddonInboxItem -DataRoot $script:Data -To 'Ziel' -Path $script:Src -From 'OCR')
        $e = @(Get-HUAddonInboxItems -DataRoot $script:Data -Id 'Ziel')[0]
        Remove-HUAddonInboxEntry -DataRoot $script:Data -Item $e
        Test-Path -LiteralPath $e.Dir | Should -BeFalse
        @(Get-HUAddonInboxItems -DataRoot $script:Data -Id 'Ziel').Count | Should -Be 0
    }
    It 'laesst Datei-Verknuepfungen aus' {
        $root = Get-HUAddonInboxRoot $script:Data 'Ziel'
        $d = Join-Path $root ('c' * 32); [void](New-Item -ItemType Directory -Path $d -Force)
        [IO.File]::WriteAllText((Join-Path $d 'eintrag.json'), '{"name":"link.pdf","zeit":"2026-11-05T10:00:00"}')
        $link = Join-Path $d 'link.pdf'
        try { [void](New-Item -ItemType SymbolicLink -Path $link -Target $script:Src -ErrorAction Stop) }
        catch { Set-ItResult -Skipped -Because 'Symlink-Recht fehlt auf diesem PC'; return }
        @(Get-HUAddonInboxItems -DataRoot $script:Data -Id 'Ziel' -Extensions '.pdf').Count | Should -Be 0
    }
    It 'lehnt ungueltige Kennungen und Eintraege ab' {
        { Get-HUAddonInboxRoot $script:Data '..\x' } | Should -Throw
        { Add-HUAddonInboxItem -DataRoot $script:Data -To 'Ziel' -Path (Join-Path $TestDrive 'fehlt.pdf') } | Should -Throw
        { Remove-HUAddonInboxEntry -DataRoot $script:Data -Item ([pscustomobject]@{ Key = '..'; Inbox = 'Ziel' }) } | Should -Throw
        Test-HUAddonId 'Postfach' | Should -BeFalse
        Test-HUAddonId 'Wiederherstellung' | Should -BeFalse
    }
}

Describe 'Modul-Update: schneller Weg ohne alle Dateien (ab 1.0.5)' {
    BeforeAll {
        $script:AD = Join-Path ([IO.Path]::GetTempPath()) ('hu-addupd-' + [guid]::NewGuid().ToString('N'))
        [void](New-Item -ItemType Directory -Path $script:AD)
        [void](New-TestAddon $script:AD 'Reisen' @{ version = '1.0.4' })
        [void](New-TestAddon (Join-Path $script:AD '.pending') 'Fristen' @{ version = '1.0.5' })
        # GitHub nachgestellt: signiertes Release, Liste mit addon.json, addon.json mit der Version aus $script:RelVer
        Mock Get-HMReleases { @([pscustomobject]@{ Version = [version]'1.0.5'; Tag = 'v1.0.5'; Prerelease = $false; ManifestUrl = 'm'; ManifestApi = 'm'; SignatureUrl = 's'; SignatureApi = 's' }) }
        Mock Get-HMReleaseAsset { [byte[]](1, 2, 3) }
        Mock Test-HMManifestSignature { '' }
        Mock ConvertFrom-HMManifest { @{ 'addon.json' = 'HASH'; 'Addon.ps1' = 'HASH'; 'Tab.xaml' = 'HASH' } }
        Mock Get-HMFileSha256 { 'HASH' }
        Mock Invoke-WebRequest {
            $j = [ordered]@{ id = $script:RelId; title = $script:RelId; version = $script:RelVer }
            [IO.File]::WriteAllText($OutFile, ($j | ConvertTo-Json), (New-Object Text.UTF8Encoding $false))
        }
    }
    AfterAll { Remove-Item -LiteralPath $script:AD -Recurse -Force -ErrorAction SilentlyContinue }
    It 'gleiche Version installiert: nur addon.json geladen, Status Aktuell' {
        $script:RelId = 'Reisen'; $script:RelVer = '1.0.4'
        $r = Install-HUAddon -Owner 'o' -Repo 'r' -Token 't' -AddonsDir $script:AD -AppVersion ([version]'9.9.9') -Thumbprint 'x'
        $r.Status | Should -Be 'Aktuell'
        Should -Invoke Invoke-WebRequest -Times 1 -Exactly
    }
    It 'Fassung wartet schon auf den Neustart: nicht noch einmal laden' {
        $script:RelId = 'Fristen'; $script:RelVer = '1.0.5'
        $r = Install-HUAddon -Owner 'o' -Repo 'r' -Token 't' -AddonsDir $script:AD -AppVersion ([version]'9.9.9') -Thumbprint 'x'
        $r.Status | Should -Be 'Aktuell'
        $r.Text | Should -BeLike '*wartet auf den Neustart*'
        Should -Invoke Invoke-WebRequest -Times 1 -Exactly
    }
    It 'neuere Version: alle Dateien werden geladen' {
        $script:RelId = 'Reisen'; $script:RelVer = '1.0.5'
        [void](Install-HUAddon -Owner 'o' -Repo 'r' -Token 't' -AddonsDir $script:AD -AppVersion ([version]'9.9.9') -Thumbprint 'x')
        Should -Invoke Invoke-WebRequest -Times 4 -Exactly
    }
}
