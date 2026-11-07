#Requires -Version 5.1
# Pester 5: Verwalter-Liste und Reihenfolge der Hilfeseiten. Kuenstliche App-Ordner unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking
    $script:AppRoot = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtSet-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](New-Item -ItemType Directory -Path (Join-Path $script:AppRoot 'Config') -Force)
    [void](New-Item -ItemType Directory -Path (Join-Path $script:AppRoot 'Docs') -Force)
    # erfundener Fingerabdruck: das echte Herausgeber-Zertifikat auf dem Test-PC darf nicht mitspielen
    [IO.File]::WriteAllText((Join-Path $script:AppRoot 'Config\update.json'), ('{ "SignerThumbprint": "' + ('A' * 40) + '" }'))
    $script:App = [pscustomobject]@{ id = 'HU-Amt'; title = 'HU-Amt'; signer = ('A' * 40); repo = [pscustomobject]@{ owner = 'x'; name = 'y' } }
    $script:Paths = Initialize-HUDataPaths -AppRoot $script:AppRoot -UserName 'anna' -PcName 'PC-1'
    $script:DataRoot = $script:Paths.DataRoot; $script:UserRoot = $script:Paths.UserRoot; $script:UserName = 'anna'
    function Read-HUJsonFile([string]$Path) { if (Test-Path -LiteralPath $Path) { return ([IO.File]::ReadAllText($Path) | ConvertFrom-Json) } return $null }
    . (Join-Path $script:Root 'Functions\Core-Update.ps1')
    . (Join-Path $script:Root 'Functions\UI-Update.ps1')
    . (Join-Path $script:Root 'Functions\UI-Settings.ps1')
    . (Join-Path $script:Root 'Functions\UI-Help.ps1')
}
AfterAll { Remove-Item -LiteralPath $script:AppRoot -Recurse -Force -ErrorAction SilentlyContinue }

Describe 'Verwalter' {
    BeforeEach { $script:IsVerwalter = $null; Remove-Item -LiteralPath (Get-HUVerwalterFile) -Force -ErrorAction SilentlyContinue }

    It 'ohne Liste und ohne Zertifikat: niemand ist Verwalter' {
        Test-HUVerwalter | Should -BeFalse
    }
    It 'Name in der Liste macht zum Verwalter (Gross-/Kleinschreibung egal)' {
        Save-HUVerwalterList @('ANNA', 'bernd') | Should -Be ''
        $script:IsVerwalter = $null
        Test-HUVerwalter | Should -BeTrue
    }
    It 'nicht in der Liste: kein Verwalter, auch wenn die Liste existiert' {
        Save-HUVerwalterList @('bernd') | Should -Be ''
        $script:IsVerwalter = $null
        Test-HUVerwalter | Should -BeFalse
    }
    It 'Kommentarzeilen und Leerzeilen zaehlen nicht' {
        [IO.File]::WriteAllText((Get-HUVerwalterFile), "# anna`r`n`r`n  bernd  `r`n")
        @(Get-HUVerwalterList) | Should -Be @('bernd')
    }
    It 'Speichern legt vorher eine Kopie an' {
        Save-HUVerwalterList @('anna') | Should -Be ''
        Save-HUVerwalterList @('bernd') | Should -Be ''
        @(Get-ChildItem -LiteralPath (Join-Path $script:DataRoot '_Kopien') -Filter 'Verwalter_*.txt').Count | Should -BeGreaterThan 0
    }
    It 'gesperrte Liste wird nicht ueberschrieben' {
        $l = Lock-HUDataFile -Path (Get-HUVerwalterFile) -UserName 'bernd' -PcName 'PC-2'
        try { Save-HUVerwalterList @('anna') | Should -Match 'bernd' } finally { Unlock-HUDataFile $l }
    }
}

Describe 'Verwalter beim Erststart' {
    BeforeAll {
        # Stellvertreter fuer Oberflaeche und Fensterzustand (keine sichtbaren Fenster im Test)
        $script:Controls = @{}; $script:Window = $null
        function Get-HUStateValue([string]$Name, $Default = $null) { if ($script:TestState.ContainsKey($Name)) { return $script:TestState[$Name] } return $Default }
        function Set-HUStateValue([string]$Name, $Value) { $script:TestState[$Name] = $Value }
        function Save-HUUIState { }
        function Confirm-HU { param([string]$Text, [string]$Title, [switch]$Warning, $Owner) $script:Asked++; return $script:Answer }
        function Show-HUMessage { param($Text, $Icon, $Owner) }
        function Write-HULogInfo([string]$m) { }
        function Write-HULogOK([string]$m) { }
    }
    BeforeEach {
        $script:IsVerwalter = $null; $script:TestState = @{}; $script:Asked = 0
        Remove-Item -LiteralPath (Get-HUVerwalterFile) -Force -ErrorAction SilentlyContinue
    }

    It 'ohne Liste und ohne Zertifikat: Frage kommt, Ja traegt den Benutzer ein' {
        $script:Answer = $true
        Request-HUFirstVerwalter | Should -BeTrue
        $script:Asked | Should -Be 1
        @(Get-HUVerwalterList) | Should -Be @('anna')
        Test-HUVerwalter | Should -BeTrue
    }
    It 'Nein: keine Liste, und dieser Benutzer wird nicht mehr gefragt' {
        $script:Answer = $false
        Request-HUFirstVerwalter | Should -BeFalse
        Get-HUVerwalterList | Should -BeNullOrEmpty
        Request-HUFirstVerwalter | Should -BeFalse
        $script:Asked | Should -Be 1
    }
    It 'bestehende Liste: keine Frage, auch wenn der Benutzer fehlt' {
        Save-HUVerwalterList @('bernd') | Should -Be ''
        $script:IsVerwalter = $null; $script:Answer = $true
        Request-HUFirstVerwalter | Should -BeFalse
        $script:Asked | Should -Be 0
        @(Get-HUVerwalterList) | Should -Be @('bernd')
    }
}

Describe 'Hilfe' {
    BeforeAll {
        $script:Mod = Join-Path $script:AppRoot 'Modules\Beispiel'
        [void](New-Item -ItemType Directory -Path $script:Mod -Force)
        $script:Info = [pscustomobject]@{ Id = 'Beispiel'; Dir = $script:Mod }
        [IO.File]::WriteAllText((Join-Path $script:AppRoot 'Docs\Grundhilfe.html'), 'g')
    }
    It 'ohne Hilfe-Repo und ohne Modulhilfe: Grundhilfe' {
        Get-HUHelpPage $script:Info | Should -Be (Join-Path $script:AppRoot 'Docs\Grundhilfe.html')
        Get-HUHelpPage $null | Should -Be (Join-Path $script:AppRoot 'Docs\Grundhilfe.html')
    }
    It 'Modulhilfe vor Grundhilfe' {
        [IO.File]::WriteAllText((Join-Path $script:Mod 'Hilfe.html'), 'm')
        Get-HUHelpPage $script:Info | Should -Be (Join-Path $script:Mod 'Hilfe.html')
    }
    It 'Hilfe aus Data\Docs vor allem anderen' {
        $d = Join-Path $script:DataRoot 'Docs\Beispiel'; [void](New-Item -ItemType Directory -Path $d -Force)
        [IO.File]::WriteAllText((Join-Path $d 'index.html'), 'd')
        Get-HUHelpPage $script:Info | Should -Be (Join-Path $d 'index.html')
    }
    It 'mitgelieferte Grundhilfe ist vorhanden und dunkel' {
        $t = [IO.File]::ReadAllText((Join-Path $script:Root 'Docs\Grundhilfe.html'), [Text.Encoding]::UTF8)
        $t | Should -Match '<html lang="de"'
        $t | Should -Match 'data-theme="dark"'
    }
}
