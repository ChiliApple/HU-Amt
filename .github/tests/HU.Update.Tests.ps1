#Requires -Version 5.1
# Pester 5: Update (Pull.ps1, Functions\Core-Update.ps1). Kein Netzzugriff; Sperren mit kuenstlichen Dateien unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:PullText = [IO.File]::ReadAllText((Join-Path $script:Root 'Pull.ps1'), [Text.Encoding]::UTF8)
    $script:LibText = [IO.File]::ReadAllText((Join-Path $script:Root 'Functions\Core-Update.ps1'), [Text.Encoding]::UTF8)
    function Get-Region([string]$Text) {
        $m = [regex]::Match($Text, '(?s)#region HMUpdateLib.*?#endregion HMUpdateLib')
        return ($m.Value -replace "`r`n", "`n")
    }
    # eine Funktion aus Pull.ps1 holen, ohne Pull auszufuehren
    $tok = $null; $err = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput($script:PullText, [ref]$tok, [ref]$err)
    $fn = $ast.FindAll({ $args[0] -is [Management.Automation.Language.FunctionDefinitionAst] -and $args[0].Name -eq 'Get-HMOpenSessions' }, $true) | Select-Object -First 1
    . ([scriptblock]::Create($fn.Extent.Text))
}

Describe 'Update-Bibliothek' {
    It 'HMUpdateLib ist in Pull.ps1 und Core-Update.ps1 wortgleich' {
        $a = Get-Region $script:PullText; $b = Get-Region $script:LibText
        $a.Length | Should -BeGreaterThan 1000
        $a | Should -BeExactly $b
    }
    It 'Konstanten passen zu Config\app.json' {
        $app = [IO.File]::ReadAllText((Join-Path $script:Root 'Config\app.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json
        $r = Get-Region $script:LibText
        $r | Should -Match ([regex]::Escape("`$script:HMManifestName  = '$($app.id)-files.sha256'"))
        $r | Should -Match ([regex]::Escape("`$script:HMDefaultOwner  = '$($app.repo.owner)'"))
        $r | Should -Match ([regex]::Escape("`$script:HMDefaultRepo   = '$($app.repo.name)'"))
        $r | Should -Match ([regex]::Escape("`$script:HMDefaultSigner = '$($app.signer)'"))
    }
    It 'Pull laesst Data, Addons, AddonsDev, bin und Config in Ruhe' {
        $m = [regex]::Match($script:PullText, "\`$keepDirs = '([^']+)'")
        $m.Success | Should -BeTrue
        foreach ($d in 'Config/x.json', 'Data/Benutzer/a/ui-state.json', 'Addons/X/addon.json', 'AddonsDev/X/addon.json', 'bin/ocr.exe') { $d | Should -Match $m.Groups[1].Value }
        'Functions/UI-Common.ps1' | Should -Not -Match $m.Groups[1].Value
    }
}

Describe 'Update im gemeinsamen Netzordner' {
    BeforeEach {
        $script:App = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtPull-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $script:Run = Join-Path $script:App 'Data\Laufend'
        [void](New-Item -ItemType Directory -Path $script:Run -Force)
    }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'ohne offene Sitzung: nichts gemeldet' { @(Get-HMOpenSessions $script:App).Count | Should -Be 0 }
    It 'verwaiste Sperre wird entfernt und nicht gemeldet' {
        Set-Content -LiteralPath (Join-Path $script:Run 'alt_PC-9.lock') -Value 'alt|PC-9|2026-01-01T08:00:00|1'
        @(Get-HMOpenSessions $script:App).Count | Should -Be 0
        Test-Path (Join-Path $script:Run 'alt_PC-9.lock') | Should -BeFalse
    }
    It 'offene Sitzung wird mit Name und PC gemeldet' {
        $p = Join-Path $script:Run 'anna_PC-1.lock'
        $fs = [IO.File]::Open($p, 'CreateNew', 'ReadWrite', 'Read')
        try {
            $b = [Text.Encoding]::UTF8.GetBytes('anna|PC-1|2026-11-03T08:00:00|1'); $fs.Write($b, 0, $b.Length); $fs.Flush()
            $r = @(Get-HMOpenSessions $script:App)
            $r.Count | Should -Be 1
            $r[0] | Should -Be 'anna am PC PC-1'
        } finally { $fs.Dispose() }
    }
    It 'Pull prueft die Sitzungen, bevor etwas geladen wird' {
        $iCheck = $script:PullText.IndexOf('$open = @(Get-HMOpenSessions $Target)')
        $iLoad = $script:PullText.IndexOf('# --- 1. Version waehlen')
        $iCheck | Should -BeGreaterThan 0
        $iCheck | Should -BeLessThan $iLoad
    }
}

Describe 'Alle Releases signieren und freigeben: was offen ist (Get-HMReleasePlan, ab 1.0.5)' {
    BeforeAll {
        . (Join-Path $script:Root 'Functions\Core-Update.ps1')
        function New-R([string]$Tag, [bool]$Pre, [bool]$Man, [bool]$Sig) {
            [pscustomobject]@{ Version = [version]($Tag.TrimStart('v')); Tag = $Tag; Prerelease = $Pre; ManifestUrl = $(if ($Man) { 'm' } else { '' }); SignatureUrl = $(if ($Sig) { 's' } else { '' }) }
        }
    }
    It 'neues Vorab-Release ohne Signatur: signieren und freigeben' {
        $p = Get-HMReleasePlan @((New-R 'v1.0.5' $true $true $false), (New-R 'v1.0.4' $false $true $true))
        $p.ToSign | Should -Be @('v1.0.5')
        $p.Publish | Should -Be 'v1.0.5'
        $p.Stable | Should -Be 'v1.0.4'
    }
    It 'schon signiertes Vorab-Release: nur freigeben; mehrere: das neueste' {
        $p = Get-HMReleasePlan @((New-R 'v1.0.3' $true $true $true), (New-R 'v1.0.4' $true $true $true), (New-R 'v1.0.2' $false $true $true))
        @($p.ToSign).Count | Should -Be 0
        $p.Publish | Should -Be 'v1.0.4'
    }
    It 'alles freigegeben: nichts offen; Release ohne Pruefsummen-Datei wird gemeldet, nicht freigegeben' {
        $p = Get-HMReleasePlan @((New-R 'v1.0.4' $false $true $true), (New-R 'v1.0.3' $true $true $true))
        @($p.ToSign).Count | Should -Be 0
        $p.Publish | Should -Be ''
        $q = Get-HMReleasePlan @((New-R 'v1.0.5' $true $false $false), (New-R 'v1.0.4' $false $true $true))
        $q.Publish | Should -Be ''
        $q.NoManifest | Should -Be @('v1.0.5')
    }
    It 'ohne freigegebenes Release: neuestes Vorab-Release ist der Kandidat' {
        (Get-HMReleasePlan @((New-R 'v0.1.0' $true $true $false))).Publish | Should -Be 'v0.1.0'
        (Get-HMReleasePlan @()).Publish | Should -Be ''
    }
}
