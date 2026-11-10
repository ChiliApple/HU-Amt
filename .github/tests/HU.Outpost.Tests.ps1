#Requires -Version 5.1
# Pester 5: Zentrale / Aussenstelle (Functions\Core-Outpost.ps1) mit einer kuenstlichen Installation unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Paths.psm1') -Force -DisableNameChecking
    . (Join-Path $script:Root 'Functions\Core-Update.ps1')
    . (Join-Path $script:Root 'Functions\Core-Addons.ps1')
    . (Join-Path $script:Root 'Functions\Core-Outpost.ps1')
    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }

    function New-TestDir { $d = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtOut-' + [guid]::NewGuid().ToString('N').Substring(0, 8)); [void](New-Item -ItemType Directory -Path $d); return $d }
    function Set-TestFile([string]$Path, [string]$Text) { $d = Split-Path $Path -Parent; if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }; [IO.File]::WriteAllText($Path, $Text, (New-Object Text.UTF8Encoding $false)) }
    # kuenstliche Installation: Programmdateien + Pruefsummen-Datei, Module, Daten der Zentrale
    function New-TestInstall {
        $app = New-TestDir
        Set-TestFile (Join-Path $app 'Main.ps1') '# Main'
        Set-TestFile (Join-Path $app 'Config\app.json') '{ "id": "HU-Amt" }'
        Set-TestFile (Join-Path $app 'Functions\A.ps1') '# A'
        $lines = foreach ($k in 'Config/app.json', 'Functions/A.ps1', 'Main.ps1') { "$(Get-HMFileSha256 (Join-Path $app ($k -replace '/', '\')))  $k" }
        $man = [Text.Encoding]::UTF8.GetBytes(($lines -join "`n") + "`n")
        foreach ($id in 'Reiseabrechnung', 'PDF', 'Buchung') {
            $j = [ordered]@{ id = $id; title = $id; version = '1.0.1'; minApp = '0.0.1'; tab = [ordered]@{ header = $id; xaml = 'Tab.xaml'; script = 'Addon.ps1'; init = "Initialize-HUAddon$id" } }
            Set-TestFile (Join-Path $app "Addons\$id\addon.json") ($j | ConvertTo-Json -Depth 5)
            Set-TestFile (Join-Path $app "Addons\$id\Addon.ps1") "# $id"
        }
        $data = Join-Path $app 'Data'
        Set-TestFile (Join-Path $data 'Verwalter.txt') "# Verwalter`r`nzentrale`r`n"
        Set-TestFile (Join-Path $data 'Reiseabrechnung\einstellungen.json') '{ "routing": { "schluessel": "geheim" } }'
        Set-TestFile (Join-Path $data 'Reiseabrechnung\reisen\r1.json') '{ "echt": true }'
        Set-TestFile (Join-Path $data 'Benutzer\zentrale\settings.json') '{}'
        return @{ App = $app; Data = $data; Manifest = $man }
    }
    function Get-ZipEntries([string]$Zip) {
        $z = [IO.Compression.ZipFile]::OpenRead($Zip)
        try { return @($z.Entries | ForEach-Object { $_.FullName -replace '\\', '/' }) } finally { $z.Dispose() }
    }
    function Read-ZipText([string]$Zip, [string]$Name) {
        $z = [IO.Compression.ZipFile]::OpenRead($Zip)
        try {
            $e = @($z.Entries | Where-Object { ($_.FullName -replace '\\', '/') -eq $Name })[0]
            if (-not $e) { return $null }
            $r = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
            try { return $r.ReadToEnd() } finally { $r.Dispose() }
        } finally { $z.Dispose() }
    }
    $script:Sources = @(
        [pscustomobject]@{ Owner = 'Besitzer'; Repo = 'HU-Amt-Reiseabrechnung'; Token = 'tok-reise' },
        [pscustomobject]@{ Owner = 'Besitzer'; Repo = 'HU-Amt-PDF'; Token = 'tok-pdf' },
        [pscustomobject]@{ Owner = 'Besitzer'; Repo = 'HU-Amt-Buchung'; Token = 'tok-buchung' }
    )
}

Describe 'Art der Installation (Config\installation.json)' {
    BeforeEach { $script:Cfg = New-TestDir }
    AfterEach { Remove-Item -LiteralPath $script:Cfg -Recurse -Force -ErrorAction SilentlyContinue }

    It 'ohne Datei: Zentrale, alle Module aktiv' {
        $c = Get-HUInstallationConfig $script:Cfg
        $c.Mode | Should -Be 'Zentrale'
        @($c.ModulesOff).Count | Should -Be 0
        Get-HUOutpostLabel $c | Should -Be ''
    }
    It 'Aussenstelle mit Namen und abgeschalteten Modulen bleibt erhalten' {
        Save-HUInstallationConfig $script:Cfg ([pscustomobject]@{ Mode = 'Aussenstelle'; Name = ' Kindergarten '; ModulesOff = @('PDF', 'OCR', 'PDF') })
        $c = Get-HUInstallationConfig $script:Cfg
        $c.Mode | Should -Be 'Aussenstelle'
        $c.Name | Should -Be 'Kindergarten'
        @($c.ModulesOff) | Should -Be @('OCR', 'PDF')
        Get-HUOutpostLabel $c | Should -Be "Au$([char]0xDF)enstelle Kindergarten"
    }
    It 'Aussenstelle ohne Namen wird abgelehnt' {
        { Save-HUInstallationConfig $script:Cfg ([pscustomobject]@{ Mode = 'Aussenstelle'; Name = ''; ModulesOff = @() }) } | Should -Throw '*Namen*'
    }
    It 'unbekannte Art wird abgelehnt, unlesbare Datei gilt als Zentrale' {
        { Save-HUInstallationConfig $script:Cfg ([pscustomobject]@{ Mode = 'Filiale'; Name = 'x'; ModulesOff = @() }) } | Should -Throw
        Set-TestFile (Join-Path $script:Cfg 'installation.json') '{ kaputt'
        (Get-HUInstallationConfig $script:Cfg).Mode | Should -Be 'Zentrale'
    }
    It 'ungueltige Modul-Kennungen werden beim Lesen verworfen' {
        Set-TestFile (Join-Path $script:Cfg 'installation.json') '{ "modus": "Zentrale", "moduleAus": [ "PDF", "..\\x", "" ] }'
        @((Get-HUInstallationConfig $script:Cfg).ModulesOff) | Should -Be @('PDF')
    }
}

Describe 'Hilfsfunktionen' {
    It 'Name fuer Ordner und ZIP' {
        ConvertTo-HUPackageName 'Kindergarten Süd / 2' | Should -Be 'Kindergarten-Süd-2'
        { ConvertTo-HUPackageName ' / ' } | Should -Throw
    }
    It 'Token zum Modul ueber den Repository-Namen' {
        (Find-HUAddonSourceFor 'PDF' $script:Sources).Token | Should -Be 'tok-pdf'
        Find-HUAddonSourceFor 'OCR' $script:Sources | Should -BeNullOrEmpty
    }
    It 'Programmdateien gegen die Pruefsummen-Datei' {
        $i = New-TestInstall
        try {
            @(Test-HUInstallFiles $i.App $i.Manifest).Count | Should -Be 0
            Set-TestFile (Join-Path $i.App 'Functions\A.ps1') '# geaendert'
            Remove-Item -LiteralPath (Join-Path $i.App 'Main.ps1')
            $bad = @(Test-HUInstallFiles $i.App $i.Manifest)
            $bad | Should -Contain 'verändert: Functions/A.ps1'
            $bad | Should -Contain 'fehlt: Main.ps1'
        } finally { Remove-Item -LiteralPath $i.App -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Paket fuer eine Aussenstelle' {
    BeforeEach { $script:I = New-TestInstall; $script:Out = New-TestDir }
    AfterEach {
        Remove-Item -LiteralPath $script:I.App -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'enthaelt Programm, gewaehlte Module, Verwalter, Einstellungen und Zugang, aber keine Daten der Zentrale' {
        $zip = Join-Path $script:Out 'HU-Amt-Kindergarten.zip'
        $mods = @([pscustomobject]@{ Id = 'Reiseabrechnung'; Active = $true; Settings = $true }, [pscustomobject]@{ Id = 'PDF'; Active = $false; Settings = $false })
        $r = New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'Kindergarten' -Admins @('verwalter1') -Modules $mods `
            -Sources $script:Sources -IncludeKeys -Manifest $script:I.Manifest -Signature ([byte[]](1, 2, 3)) -Version '1.0.1' -Channel 'Stable'
        $r.Path | Should -Be $zip
        $r.Files | Should -Be 3
        $r.Folder | Should -Be 'HU-Amt-Kindergarten'
        $e = Get-ZipEntries $zip
        $p = 'HU-Amt-Kindergarten/'
        foreach ($f in 'Main.ps1', 'Functions/A.ps1', 'Config/app.json', 'Config/HU-Amt-files.sha256', 'Config/HU-Amt-files.sha256.p7s', 'Config/installation.json',
            'Config/installed.json', 'Config/update.json', 'Config/paket.json', 'Data/Verwalter.txt', 'Data/Reiseabrechnung/einstellungen.json',
            'Addons/Reiseabrechnung/addon.json', 'Addons/PDF/Addon.ps1', 'Zugang.humodule', 'LIESMICH.txt') { $e | Should -Contain "$p$f" }
        @($e | Where-Object { $_ -match '/Addons/Buchung/|/Benutzer/|/reisen/' }).Count | Should -Be 0
        (Read-ZipText $zip "${p}Data/Verwalter.txt") | Should -Match '(?m)^verwalter1\r?$'
        (Read-ZipText $zip "${p}Data/Verwalter.txt") | Should -Not -Match 'zentrale'
        $inst = (Read-ZipText $zip "${p}Config/installation.json") | ConvertFrom-Json
        $inst.modus | Should -Be 'Aussenstelle'
        $inst.name | Should -Be 'Kindergarten'
        @($inst.moduleAus) | Should -Be @('PDF')
        $key = (Read-ZipText $zip "${p}Zugang.humodule") | ConvertFrom-Json
        @($key.module | ForEach-Object { $_.repo }) | Should -Be @('Besitzer/HU-Amt-Reiseabrechnung', 'Besitzer/HU-Amt-PDF')
        (Read-ZipText $zip "${p}Zugang.humodule") | Should -Not -Match 'tok-buchung'
        $pk = (Read-ZipText $zip "${p}Config/paket.json") | ConvertFrom-Json
        $pk.zugang | Should -Be 'Zugang.humodule'
        $pk.geprueft | Should -Be $false
        ((Read-ZipText $zip "${p}Config/installed.json") | ConvertFrom-Json).FileList | Should -Be @('Config/app.json', 'Functions/A.ps1', 'Main.ps1')
    }
    It 'ohne Zugang: keine Zugangsdatei, Hinweis bei Modul ohne Token' {
        $zip = Join-Path $script:Out 'a.zip'
        $r = New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'Bauhof' -Admins @('x') -Manifest $script:I.Manifest `
            -Modules @([pscustomobject]@{ Id = 'PDF'; Active = $true; Settings = $true })
        Get-ZipEntries $zip | Should -Not -Contain 'HU-Amt-Bauhof/Zugang.humodule'
        @($r.Warnings) -join ' ' | Should -Match 'keine Einstellungen'
        $zip2 = Join-Path $script:Out 'b.zip'
        $r2 = New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip2 -Name 'Bauhof' -Admins @('x') -Manifest $script:I.Manifest `
            -Modules @([pscustomobject]@{ Id = 'PDF'; Active = $true; Settings = $false }) -Sources @() -IncludeKeys
        @($r2.Warnings) -join ' ' | Should -Match 'kein Token'
    }
    It 'veraenderte Programmdatei: kein Paket' {
        Set-TestFile (Join-Path $script:I.App 'Main.ps1') '# fremd'
        $zip = Join-Path $script:Out 'c.zip'
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @('x') -Manifest $script:I.Manifest } | Should -Throw '*nicht zur freigegebenen Version*'
        Test-Path -LiteralPath $zip | Should -BeFalse
    }
    It 'vorhandene Datei, fehlendes Modul, fehlender Verwalter werden abgelehnt' {
        $zip = Join-Path $script:Out 'd.zip'; Set-TestFile $zip 'x'
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @('x') -Manifest $script:I.Manifest } | Should -Throw '*gibt es schon*'
        $zip = Join-Path $script:Out 'e.zip'
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @('x') -Manifest $script:I.Manifest -Modules @([pscustomobject]@{ Id = 'OCR'; Active = $true; Settings = $false }) } | Should -Throw '*nicht vorhanden*'
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @(' ') -Manifest $script:I.Manifest } | Should -Throw '*Verwalter*'
        Test-Path -LiteralPath $zip | Should -BeFalse
    }
}

Describe 'Erster Start eines Pakets' {
    BeforeEach { $script:App = New-TestDir }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'ohne paket.json passiert nichts' {
        (Test-HUPackageFirstStart $script:App ('A' * 40)).Status | Should -Be 'Keins'
        Get-HUPackageKeyFile $script:App | Should -Be ''
    }
    It 'fehlende Signatur wird gemeldet, danach nicht noch einmal' {
        Set-TestFile (Join-Path $script:App 'Config\paket.json') '{ "name": "K", "erstellt": "x", "geprueft": false, "zugangSha256": "" }'
        (Test-HUPackageFirstStart $script:App ('A' * 40)).Status | Should -Be 'Fehler'
        (Test-HUPackageFirstStart $script:App ('A' * 40)).Status | Should -Be 'Keins'
    }
    It 'Zugangsdatei gilt nur mit passender Pruefsumme' {
        $kf = Join-Path $script:App 'Zugang.humodule'
        Set-TestFile $kf '{ "module": [ { "repo": "Besitzer/HU-Amt-PDF", "token": "t" } ] }'
        Set-TestFile (Join-Path $script:App 'Config\paket.json') ('{ "name": "K", "geprueft": true, "zugangSha256": "' + (Get-HMFileSha256 $kf) + '" }')
        Get-HUPackageKeyFile $script:App | Should -Be $kf
        Set-TestFile $kf '{ "module": [ { "repo": "Fremd/HU-Amt-PDF", "token": "t" } ] }'
        Get-HUPackageKeyFile $script:App | Should -Be ''
    }
}
