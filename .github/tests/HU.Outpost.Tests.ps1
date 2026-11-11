#Requires -Version 5.1
# Pester 5: Zentrale / Aussenstelle (Functions\Core-Outpost.ps1, Einrichten.ps1) mit einer kuenstlichen Installation unter TEMP.

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
    function New-TestManifest([string]$Dir, [string[]]$Files) {
        $lines = foreach ($k in ($Files | Sort-Object)) { "$(Get-HMFileSha256 (Join-Path $Dir ($k -replace '/', '\')))  $k" }
        return , [Text.Encoding]::UTF8.GetBytes(($lines -join "`n") + "`n")
    }
    # kuenstliche Installation: Programmdateien (mit dem echten Einrichten.ps1) + Pruefsummen-Datei, Module, Daten der Zentrale
    function New-TestInstall {
        $app = New-TestDir
        Set-TestFile (Join-Path $app 'Main.ps1') '# Main'
        Set-TestFile (Join-Path $app 'Start-HU-Amt.cmd') '@echo off'
        Set-TestFile (Join-Path $app 'Config\app.json') '{ "id": "HU-Amt" }'
        Set-TestFile (Join-Path $app 'Functions\A.ps1') '# A'
        Copy-Item -LiteralPath (Join-Path $script:Root 'Einrichten.ps1') -Destination (Join-Path $app 'Einrichten.ps1')
        $man = New-TestManifest $app @('Config/app.json', 'Einrichten.ps1', 'Functions/A.ps1', 'Main.ps1', 'Start-HU-Amt.cmd')
        foreach ($id in 'Reiseabrechnung', 'PDF', 'Buchung') {
            $j = [ordered]@{ id = $id; title = $id; version = '1.0.1'; minApp = '0.0.1'; tab = [ordered]@{ header = $id; xaml = 'Tab.xaml'; script = 'Addon.ps1'; init = "Initialize-HUAddon$id" } }
            Set-TestFile (Join-Path $app "Addons\$id\addon.json") ($j | ConvertTo-Json -Depth 5)
            Set-TestFile (Join-Path $app "Addons\$id\Addon.ps1") "# $id"
        }
        # Reise mit signierter Liste (Signatur hier ohne Belang) und einer fremden Datei, die nicht mitkommen darf
        $rd = Join-Path $app 'Addons\Reiseabrechnung'
        [IO.File]::WriteAllBytes((Join-Path $rd 'addon.sha256'), (New-TestManifest $rd @('Addon.ps1', 'addon.json')))
        Set-TestFile (Join-Path $rd 'addon.sha256.p7s') 'sig'
        Set-TestFile (Join-Path $rd 'fremd.dll') 'x'
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
        [pscustomobject]@{ Owner = 'Besitzer'; Repo = 'HU-Amt-Buchung'; Token = 'tok-buchung' },
        [pscustomobject]@{ Owner = 'Besitzer'; Repo = 'HU-Amt-Kassa-PDF'; Token = 'tok-falsch' }
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
    It 'Token zum Modul nur bei genauem Repository-Namen' {
        (Find-HUAddonSourceFor 'PDF' $script:Sources).Token | Should -Be 'tok-pdf'
        Find-HUAddonSourceFor 'OCR' $script:Sources | Should -BeNullOrEmpty
        (Find-HUAddonSourceFor 'PDF' @($script:Sources[3], $script:Sources[1])).Token | Should -Be 'tok-pdf'
        Find-HUAddonSourceFor 'PDF' @($script:Sources[3]) | Should -BeNullOrEmpty
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
    It 'Moduldateien nur laut signierter Liste' {
        $i = New-TestInstall
        try {
            $f = @(Get-HUAddonPackageFiles (Join-Path $i.App 'Addons\Reiseabrechnung'))
            $f | Should -Contain 'Addon.ps1'
            $f | Should -Contain 'addon.sha256'
            $f | Should -Not -Contain 'fremd.dll'
            @(Get-HUAddonPackageFiles (Join-Path $i.App 'Addons\PDF')).Count | Should -Be 2
        } finally { Remove-Item -LiteralPath $i.App -Recurse -Force -ErrorAction SilentlyContinue }
    }
    It 'alte Paket-Reste in TEMP werden entfernt, frische bleiben' {
        $old = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtPaket-test' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $new = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtPaket-test' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        [void](New-Item -ItemType Directory -Path $old); [void](New-Item -ItemType Directory -Path $new)
        (Get-Item -LiteralPath $old).LastWriteTime = (Get-Date).AddHours(-3)
        try {
            Remove-HUOldPackageTemp
            Test-Path -LiteralPath $old | Should -BeFalse
            Test-Path -LiteralPath $new | Should -BeTrue
        } finally { Remove-Item -LiteralPath $old, $new -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Paket fuer eine Aussenstelle' {
    BeforeEach { $script:I = New-TestInstall; $script:Out = New-TestDir }
    AfterEach {
        Remove-Item -LiteralPath $script:I.App -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:Out -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'enthaelt Einrichten, Programm, gewaehlte Module, Verwalter, Einstellungen und Zugang, aber keine Daten der Zentrale' {
        $zip = Join-Path $script:Out 'HU-Amt-Kindergarten.zip'
        $mods = @([pscustomobject]@{ Id = 'Reiseabrechnung'; Active = $true; Settings = $true }, [pscustomobject]@{ Id = 'PDF'; Active = $false; Settings = $true })
        $r = New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'Kindergarten' -Admins @('verwalter1') -Modules $mods `
            -Sources $script:Sources -IncludeKeys -Manifest $script:I.Manifest -Signature ([byte[]](1, 2, 3)) -Version '1.0.2' -Channel 'Stable'
        $r.Path | Should -Be $zip
        $r.Files | Should -Be 5
        @($r.Warnings).Count | Should -Be 0   # PDF ohne Einstellungen ist kein Hinweis mehr
        $e = Get-ZipEntries $zip
        foreach ($f in 'Einrichten.cmd', 'LIESMICH.txt') { $e | Should -Contain $f }
        $p = 'Programm/'
        foreach ($f in 'Main.ps1', 'Einrichten.ps1', 'Functions/A.ps1', 'Config/app.json', 'Config/HU-Amt-files.sha256', 'Config/HU-Amt-files.sha256.p7s', 'Config/installation.json',
            'Config/installed.json', 'Config/update.json', 'Config/paket.json', 'Data/Verwalter.txt', 'Data/Reiseabrechnung/einstellungen.json',
            'Addons/Reiseabrechnung/addon.json', 'Addons/Reiseabrechnung/addon.sha256', 'Addons/PDF/Addon.ps1', 'Zugang.humodule') { $e | Should -Contain "$p$f" }
        @($e | Where-Object { $_ -match '/Addons/Buchung/|/Benutzer/|/reisen/|fremd\.dll' }).Count | Should -Be 0
        (Read-ZipText $zip "${p}Data/Verwalter.txt") | Should -Match '(?m)^verwalter1\r?$'
        (Read-ZipText $zip "${p}Data/Verwalter.txt") | Should -Not -Match 'zentrale'
        $inst = (Read-ZipText $zip "${p}Config/installation.json") | ConvertFrom-Json
        $inst.modus | Should -Be 'Aussenstelle'
        $inst.name | Should -Be 'Kindergarten'
        @($inst.moduleAus) | Should -Be @('PDF')
        $key = (Read-ZipText $zip "${p}Zugang.humodule") | ConvertFrom-Json
        @($key.module | ForEach-Object { $_.repo }) | Should -Be @('Besitzer/HU-Amt-Reiseabrechnung', 'Besitzer/HU-Amt-PDF')
        (Read-ZipText $zip "${p}Zugang.humodule") | Should -Not -Match 'tok-buchung|tok-falsch'
        $pk = (Read-ZipText $zip "${p}Config/paket.json") | ConvertFrom-Json
        $pk.zugang | Should -Be 'Zugang.humodule'
        $pk.geprueft | Should -Be $false
        @($pk.verwalter) | Should -Be @('verwalter1')
        ((Read-ZipText $zip "${p}Config/installed.json") | ConvertFrom-Json).FileList | Should -Be @('Config/app.json', 'Einrichten.ps1', 'Functions/A.ps1', 'Main.ps1', 'Start-HU-Amt.cmd')
        $lies = Read-ZipText $zip 'LIESMICH.txt'
        $lies | Should -Match 'Alle extrahieren'
        $lies | Should -Match 'verwalter1'
        $lies | Should -Not -Match '[„“…–]'
        (Read-ZipText $zip 'Einrichten.cmd') | Should -Match 'Programm\\Einrichten\.ps1'
    }
    It 'ohne Zugang: keine Zugangsdatei, Hinweis bei Modul ohne Token' {
        $zip = Join-Path $script:Out 'a.zip'
        [void](New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'Bauhof' -Admins @('x') -Manifest $script:I.Manifest `
            -Modules @([pscustomobject]@{ Id = 'PDF'; Active = $true; Settings = $true }))
        Get-ZipEntries $zip | Should -Not -Contain 'Programm/Zugang.humodule'
        $zip2 = Join-Path $script:Out 'b.zip'
        $r2 = New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip2 -Name 'Bauhof' -Admins @('x') -Manifest $script:I.Manifest `
            -Modules @([pscustomobject]@{ Id = 'PDF'; Active = $true; Settings = $false }) -Sources @($script:Sources[3]) -IncludeKeys
        @($r2.Warnings) -join ' ' | Should -Match 'kein Zugang'
    }
    It 'vorhandene Datei: nur mit -Overwrite, ersetzt erst nach Erfolg' {
        $zip = Join-Path $script:Out 'c.zip'; Set-TestFile $zip 'alt'
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @('x') -Manifest $script:I.Manifest } | Should -Throw '*gibt es schon*'
        # Fehler mitten im Paket: alte Datei bleibt, keine Reste
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @('x') -Manifest $script:I.Manifest -Overwrite -Modules @([pscustomobject]@{ Id = 'OCR'; Active = $true; Settings = $false }) } | Should -Throw '*nicht vorhanden*'
        [IO.File]::ReadAllText($zip) | Should -Be 'alt'
        Test-Path -LiteralPath "$zip.tmp" | Should -BeFalse
        [void](New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @('x') -Manifest $script:I.Manifest -Overwrite)
        Get-ZipEntries $zip | Should -Contain 'Einrichten.cmd'
    }
    It 'veraenderte Programmdatei: kein Paket, mit naechstem Schritt' {
        Set-TestFile (Join-Path $script:I.App 'Main.ps1') '# fremd'
        $zip = Join-Path $script:Out 'd.zip'
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @('x') -Manifest $script:I.Manifest } | Should -Throw '*Update*'
        Test-Path -LiteralPath $zip | Should -BeFalse
    }
    It 'fehlender Verwalter wird abgelehnt' {
        $zip = Join-Path $script:Out 'e.zip'
        { New-HUOutpostPackage -AppRoot $script:I.App -DataRoot $script:I.Data -ZipPath $zip -Name 'K' -Admins @(' ') -Manifest $script:I.Manifest } | Should -Throw '*Verwalter*'
    }
}

Describe 'Einrichten in der Aussenstelle (Einrichten.ps1)' {
    BeforeAll {
        $script:I2 = New-TestInstall; $script:Out2 = New-TestDir
        $script:Zip2 = Join-Path $script:Out2 'p.zip'
        [void](New-HUOutpostPackage -AppRoot $script:I2.App -DataRoot $script:I2.Data -ZipPath $script:Zip2 -Name 'Kindergarten' -Admins @('verwalter1') -Manifest $script:I2.Manifest `
            -Modules @([pscustomobject]@{ Id = 'Reiseabrechnung'; Active = $true; Settings = $true }))
        $script:Ex = Join-Path $script:Out2 'ausgepackt'
        [IO.Compression.ZipFile]::ExtractToDirectory($script:Zip2, $script:Ex)
        $script:Tgt = Join-Path $script:Out2 'lokal'
        $script:Ps = (Get-Process -Id $PID).Path
        function Invoke-Einrichten { & $script:Ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $script:Ex 'Programm\Einrichten.ps1') -TargetRoot $script:Tgt -NoStart -Quiet | Out-Null; return $LASTEXITCODE }
    }
    AfterAll {
        Remove-Item -LiteralPath $script:I2.App -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:Out2 -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'kopiert das Paket in den Zielordner HU-Amt-<Name>' {
        Invoke-Einrichten | Should -Be 0
        $t = Join-Path $script:Tgt 'HU-Amt-Kindergarten'
        foreach ($f in 'Main.ps1', 'Config\installation.json', 'Config\paket.json', 'Data\Verwalter.txt', 'Data\Reiseabrechnung\einstellungen.json', 'Addons\Reiseabrechnung\addon.json') {
            Test-Path -LiteralPath (Join-Path $t $f) | Should -BeTrue
        }
    }
    It 'erneutes Einrichten erneuert das Programm, Daten bleiben' {
        $t = Join-Path $script:Tgt 'HU-Amt-Kindergarten'
        Set-TestFile (Join-Path $t 'Data\Reiseabrechnung\reisen\eigene.json') '{ "meine": 1 }'
        Set-TestFile (Join-Path $t 'Main.ps1') '# kaputt'
        $pj = Join-Path $t 'Config\paket.json'
        $p = [IO.File]::ReadAllText($pj) | ConvertFrom-Json; $p.geprueft = $true; [IO.File]::WriteAllText($pj, ($p | ConvertTo-Json))
        Invoke-Einrichten | Should -Be 0
        [IO.File]::ReadAllText((Join-Path $t 'Data\Reiseabrechnung\reisen\eigene.json')) | Should -Match 'meine'
        [IO.File]::ReadAllText((Join-Path $t 'Main.ps1')) | Should -Be '# Main'
        ([IO.File]::ReadAllText($pj) | ConvertFrom-Json).geprueft | Should -Be $false
    }
    It 'neues Paket ohne ein Modul: das Modul wird entfernt, seine Daten bleiben' {
        $t = Join-Path $script:Tgt 'HU-Amt-Kindergarten'
        Set-TestFile (Join-Path $t 'Addons\PDF\addon.json') '{ "id": "PDF" }'
        Set-TestFile (Join-Path $t 'Data\PDF\eigene.json') '{ "pdf": 1 }'
        Invoke-Einrichten | Should -Be 0
        Test-Path -LiteralPath (Join-Path $t 'Addons\PDF') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $t 'Addons\Reiseabrechnung\addon.json') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $t 'Data\PDF\eigene.json') | Should -BeTrue
    }
    It 'Modul als Verknuepfung (Junction): nur die Verknuepfung geht, das Ziel bleibt (Pruefung 092)' {
        $t = Join-Path $script:Tgt 'HU-Amt-Kindergarten'
        $outside = New-TestDir
        Set-TestFile (Join-Path $outside 'wichtig.txt') 'bleibt'
        $link = Join-Path $t 'Addons\Verlinkt'
        [void](New-Item -ItemType Junction -Path $link -Target $outside)
        try {
            Invoke-Einrichten | Should -Be 0
            Test-Path -LiteralPath $link | Should -BeFalse
            Test-Path -LiteralPath (Join-Path $outside 'wichtig.txt') | Should -BeTrue
        } finally {
            if (Test-Path -LiteralPath $link) { [IO.Directory]::Delete($link, $false) }
            Remove-Item -LiteralPath $outside -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    It 'aus einem unvollstaendigen Ordner: Abbruch mit Hinweis aufs Auspacken' {
        $half = Join-Path $script:Out2 'halb'
        [void](New-Item -ItemType Directory -Path $half)
        Copy-Item -LiteralPath (Join-Path $script:Ex 'Programm\Einrichten.ps1') -Destination $half
        $o = & $script:Ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $half 'Einrichten.ps1') -TargetRoot $script:Tgt -NoStart -Quiet
        $LASTEXITCODE | Should -Be 1
        ($o -join ' ') | Should -Match 'Alle extrahieren'
    }
}

Describe 'Zugaenge ohne Modul nach einem neuen Paket' {
    BeforeAll {
        $script:App3 = New-TestDir
        Set-TestFile (Join-Path $script:App3 'Addons\Reiseabrechnung\addon.json') '{ "id": "Reiseabrechnung" }'
        Set-TestFile (Join-Path $script:App3 'Addons\OhneJson\Addon.ps1') '# kein addon.json'
    }
    AfterAll { Remove-Item -LiteralPath $script:App3 -Recurse -Force -ErrorAction SilentlyContinue }

    It 'meldet nur Zugaenge, zu denen kein Modul in Addons liegt' {
        $r = @(Get-HUPackageStaleSources -AppRoot $script:App3 -Sources $script:Sources)
        @($r | ForEach-Object { $_.Repo }) | Should -Be @('HU-Amt-PDF', 'HU-Amt-Buchung', 'HU-Amt-Kassa-PDF')
    }
    It 'ohne Addons-Ordner gelten alle Zugaenge als ueberzaehlig' {
        $e = New-TestDir
        try { @(Get-HUPackageStaleSources -AppRoot $e -Sources $script:Sources).Count | Should -Be 4 } finally { Remove-Item -LiteralPath $e -Recurse -Force }
    }
}

Describe 'Erster Start eines Pakets' {
    BeforeEach { $script:App = New-TestDir }
    AfterEach { Remove-Item -LiteralPath $script:App -Recurse -Force -ErrorAction SilentlyContinue }

    It 'ohne paket.json passiert nichts' {
        (Test-HUPackageFirstStart $script:App ('A' * 40)).Status | Should -Be 'Keins'
        Get-HUPackageKeyFile $script:App | Should -Be ''
    }
    It 'fehlende Signatur wird bei jedem Start gemeldet, bis die Pruefung klappt' {
        Set-TestFile (Join-Path $script:App 'Config\paket.json') '{ "name": "K", "erstellt": "x", "geprueft": false, "zugangSha256": "" }'
        (Test-HUPackageFirstStart $script:App ('A' * 40)).Status | Should -Be 'Fehler'
        (Test-HUPackageFirstStart $script:App ('A' * 40)).Status | Should -Be 'Fehler'
        ([IO.File]::ReadAllText((Join-Path $script:App 'Config\paket.json')) | ConvertFrom-Json).geprueft | Should -Be $false
    }
    It 'schon erfolgreich geprueft: nichts mehr zu tun' {
        Set-TestFile (Join-Path $script:App 'Config\paket.json') '{ "name": "K", "geprueft": true }'
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
