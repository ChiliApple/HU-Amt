#Requires -Version 5.1
# Pester 5: App-Manifest, Version, Kodierung und Inhalte des Repos.

BeforeAll {
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:Files = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File |
            Where-Object { $_.FullName -notmatch '\\\.git\\' -and $_.Extension -in '.ps1', '.psm1', '.xaml', '.json', '.md', '.cmd', '.yml', '.html' })
}

Describe 'App-Manifest' {
    It 'Config\app.json hat Id, Titel, Repo und Signatur' {
        $a = [IO.File]::ReadAllText((Join-Path $script:Root 'Config\app.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
        $a.id | Should -Match '^[A-Za-z0-9][A-Za-z0-9\-]+$'
        $a.title | Should -Not -BeNullOrEmpty
        $a.repo.owner | Should -Not -BeNullOrEmpty
        $a.repo.name | Should -Not -BeNullOrEmpty
        $a.signer | Should -Match '^[0-9A-F]{40}$'
    }
    It 'Version in version.json steht oben im CHANGELOG' {
        $v = ([IO.File]::ReadAllText((Join-Path $script:Root 'Config\version.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json).version
        $v | Should -Match '^\d+\.\d+\.\d+$'
        $first = @(Get-Content -LiteralPath (Join-Path $script:Root 'CHANGELOG.md') -Encoding UTF8 | Where-Object { $_ -match '^## v' })[0]
        $first | Should -Match ("^## v" + [regex]::Escape($v) + '(\s|$)')
    }
}

Describe 'Aufraeumen und Benutzerordner' {
    It 'Start loescht alte Protokolle von App und Modulen (alle *.log in Logs)' {
        $t = [IO.File]::ReadAllText((Join-Path $script:Root 'Main.ps1'), [Text.Encoding]::UTF8)
        $t | Should -Match "-LiteralPath \`$script:LogDir -Filter '\*\.log'"
    }
    It 'Pull bildet den Benutzerordner wie Get-HUSafeName (Punkt und Leerzeichen am Ende weg)' {
        Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force
        $t = [IO.File]::ReadAllText((Join-Path $script:Root 'Pull.ps1'), [Text.Encoding]::UTF8)
        $m = [regex]::Match($t, '\$tokUser = (\(.+?\)\.Trim\(.+?\));')
        $m.Success | Should -BeTrue
        foreach ($n in 'anna', 'max.', 'a b ', 'x/y', '...') {
            $env:USERNAME_TEST = $n
            $pull = & ([scriptblock]::Create(($m.Groups[1].Value -replace '\$env:USERNAME', '$env:USERNAME_TEST')))
            if (-not $pull) { $pull = 'unbekannt' }
            $pull | Should -Be (Get-HUSafeName $n)
        }
        Remove-Item Env:\USERNAME_TEST -ErrorAction SilentlyContinue
    }
}

Describe 'Pull.ps1: Zielordner der Erstinstallation' {
    BeforeAll {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root 'Pull.ps1'), [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Test-HMPullEmptyHome' }, $true)
        . ([scriptblock]::Create($fn.Extent.Text))
        $script:Base = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtPull-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        [void](New-Item -ItemType Directory -Path $script:Base)
    }
    AfterAll { Remove-Item -LiteralPath $script:Base -Recurse -Force -ErrorAction SilentlyContinue }
    It 'Pull.ps1 allein im Ordner: dorthin installieren' {
        $d = Join-Path $script:Base 'leer'; [void](New-Item -ItemType Directory -Path $d)
        $p = Join-Path $d 'Pull.ps1'; Set-Content -LiteralPath $p -Value '#'
        Set-Content -LiteralPath (Join-Path $d 'desktop.ini') -Value 'x'
        Test-HMPullEmptyHome $d $p | Should -BeTrue
    }
    It 'weitere Dateien im Ordner: nicht dorthin' {
        $d = Join-Path $script:Base 'voll'; [void](New-Item -ItemType Directory -Path $d)
        $p = Join-Path $d 'Pull.ps1'; Set-Content -LiteralPath $p -Value '#'
        Set-Content -LiteralPath (Join-Path $d 'Rechnung.pdf') -Value 'x'
        Test-HMPullEmptyHome $d $p | Should -BeFalse
    }
    It 'Desktop, Downloads und Laufwerkswurzel: nie' {
        Test-HMPullEmptyHome (Join-Path $env:USERPROFILE 'Downloads') 'x' | Should -BeFalse
        Test-HMPullEmptyHome (Join-Path $env:USERPROFILE 'Desktop') 'x' | Should -BeFalse
        Test-HMPullEmptyHome 'C:\' 'x' | Should -BeFalse
    }
    It 'kein Befehl mit relativem Pfad zu Pull.ps1 in Doku und Skripten' {
        # .github ausgenommen: die CI laeuft im Repo-Ordner, dort ist der relative Pfad richtig
        $hits = @($script:Files | Where-Object { $_.FullName -notmatch '\\\.github\\' } | Select-String -Pattern '-File\s+\.?\\?Pull\.ps1')
        ($hits | ForEach-Object { "$($_.Filename):$($_.LineNumber)" }) -join ', ' | Should -BeNullOrEmpty
    }
}

Describe 'Inhalte' {
    It 'der Name steht nur in app.json, nicht fest im Code (kein Name der Vorlage)' {
        # Muster zusammengesetzt, damit diese Datei sich nicht selbst findet
        $old = 'HU-Multi' + 'Tenant'
        $hits = @($script:Files | Select-String -Pattern ($old + '|' + $old.Replace('-', '')))
        $hits.Count | Should -Be 0 -Because (($hits | ForEach-Object { "$($_.Filename):$($_.LineNumber)" }) -join ', ')
    }
    It 'keine Namen von KI-Werkzeugen' {
        $names = @(('cla' + 'ude'), ('anthro' + 'pic'), ('gr' + 'ok'), ('gem' + 'ini'), ('chat' + 'gpt'), ('open' + 'ai'))
        $hits = @($script:Files | Select-String -Pattern ('\b(' + ($names -join '|') + ')\b'))
        $hits.Count | Should -Be 0 -Because (($hits | ForEach-Object { "$($_.Filename):$($_.LineNumber)" }) -join ', ')
    }
    It 'keine internen IP-Adressen' {
        $hits = @($script:Files | Select-String -Pattern '\b(10\.\d{1,3}|172\.(1[6-9]|2\d|3[01])|192\.168)\.\d{1,3}\.\d{1,3}\b')
        $hits.Count | Should -Be 0 -Because (($hits | ForEach-Object { "$($_.Filename):$($_.LineNumber)" }) -join ', ')
    }
    It 'keine GitHub-Token' {
        @($script:Files | Select-String -Pattern (('github' + '_pat_') + '|' + ('gh' + 'p_[A-Za-z0-9]{20}'))).Count | Should -Be 0
    }
}

Describe 'Kodierung und Syntax' {
    It 'PowerShell-Dateien sind UTF-8 mit BOM und fehlerfrei' {
        foreach ($f in @($script:Files | Where-Object { $_.Extension -in '.ps1', '.psm1' })) {
            $b = [IO.File]::ReadAllBytes($f.FullName)
            ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) | Should -BeTrue -Because "$($f.Name) braucht BOM (PowerShell 5.1)"
            $tok = $null; $err = $null
            [void][Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tok, [ref]$err)
            @($err).Count | Should -Be 0 -Because "$($f.Name): $(@($err | ForEach-Object { $_.Message }) -join '; ')"
        }
    }
    It 'XAML-Dateien sind gueltiges XML' {
        foreach ($f in @($script:Files | Where-Object { $_.Extension -eq '.xaml' })) {
            { [xml][IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8) } | Should -Not -Throw -Because $f.Name
        }
    }
}
