#Requires -Version 5.1
# Pester 5: Wächter für externe Werkzeuge der Module (.github\waechter). Ohne Netz, nur künstliche Daten unter TEMP.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:Root '.github\waechter\Waechter.Funktionen.ps1')
    function New-TestModul([hashtable]$Konfig, [hashtable]$Paket = $null, [string]$Req = '') {
        $d = Join-Path ([IO.Path]::GetTempPath()) ('huw-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path (Join-Path $d '.github') -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $d '.github\waechter.json'), ($Konfig | ConvertTo-Json -Depth 6))
        if ($Paket) { [IO.File]::WriteAllText((Join-Path $d 'werkzeuge.json'), ($Paket | ConvertTo-Json -Depth 6)) }
        if ($Req) { [IO.File]::WriteAllText((Join-Path $d 'requirements.txt'), $Req) }
        return $d
    }
}

Describe 'Versionen und Tags' {
    It 'vergleicht Versionen' {
        Compare-HUWVersion '10.08.0' '10.04.0' | Should -Be 1
        Compare-HUWVersion '3.12.7' '3.12.16' | Should -Be -1
        Compare-HUWVersion '26.00' '26.0' | Should -Be 0
        Compare-HUWVersion 'v5.4.0.20240606' '5.4.0.20240606' | Should -Be 0
        Compare-HUWVersion '1.2' '1.2.1' | Should -Be -1
    }
    It 'setzt Tags in Versionen um' {
        ConvertFrom-HUWTag 'v12.4.2' | Should -Be '12.4.2'
        ConvertFrom-HUWTag 'gs10080' '^gs(\d{2})(\d{2})(\d)$' '{1}.{2:00}.{3}' | Should -Be '10.08.0'
        ConvertFrom-HUWTag 'gs10080rc1' '^gs(\d{2})(\d{2})(\d)$' '{1}.{2:00}.{3}' | Should -Be ''
    }
    It 'kürzt die Version für die CPE-Abfrage' {
        Get-HUWCpeVersion '5.4.0.20240606' 3 | Should -Be '5.4.0'
        Get-HUWCpeVersion '10.08.0' 0 | Should -Be '10.08.0'
    }
}

Describe 'Status, Signatur, Bericht' {
    It 'bestimmt den Status' {
        Get-HUWStatus @{ Fehler = 'x'; Luecken = @(@{ Id = 'a' }) } | Should -Be 'fehler'
        Get-HUWStatus @{ Fehler = ''; Luecken = @(@{ Id = 'a' }); Neuer = $true } | Should -Be 'luecke'
        Get-HUWStatus @{ Fehler = ''; Luecken = @(); EolVorbei = $true } | Should -Be 'eol'
        Get-HUWStatus @{ Fehler = ''; Luecken = @(); Neuer = $true } | Should -Be 'neuer'
        Get-HUWStatus @{ Fehler = ''; Luecken = @() } | Should -Be 'aktuell'
    }
    It 'ändert die Signatur nur mit den Befunden' {
        $a = @(@{ Id = 'gs'; Status = 'luecke'; Neueste = '10.08.0'; Luecken = @(@{ Id = 'CVE-1' }, @{ Id = 'CVE-2' }) }, @{ Id = 'qpdf'; Status = 'aktuell'; Neueste = '1'; Luecken = @() })
        $b = @(@{ Id = 'qpdf'; Status = 'aktuell'; Neueste = '2'; Luecken = @() }, @{ Id = 'gs'; Status = 'luecke'; Neueste = '10.08.0'; Luecken = @(@{ Id = 'CVE-2' }, @{ Id = 'CVE-1' }) })
        $c = @(@{ Id = 'gs'; Status = 'luecke'; Neueste = '10.08.0'; Luecken = @(@{ Id = 'CVE-1' }) })
        Get-HUWSignatur $a | Should -Be (Get-HUWSignatur $b)
        Get-HUWSignatur $a | Should -Not -Be (Get-HUWSignatur $c)
        Get-HUWSignatur $a | Should -Match '^[0-9a-f]{16}$'
    }
    It 'liest requirements.txt' {
        $f = Join-Path ([IO.Path]::GetTempPath()) ('huw-' + [guid]::NewGuid().ToString('N') + '.txt')
        try {
            [IO.File]::WriteAllText($f, "# Kopf`r`nPillow==12.0.0 --hash=sha256:aa`r`n`r`npdfminer.six==20250506 --hash=sha256:bb`r`n")
            $r = Read-HUWRequirements $f
            $r.Count | Should -Be 2
            $r[0].Name | Should -Be 'Pillow'
            $r[1].Version | Should -Be '20250506'
        } finally { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
    }
    It 'schreibt den Bericht mit Signatur, schwerste Lücke zuerst' {
        $e = [ordered]@{ modul = 'Test'; paket = '2'; stand = '2026-11-05T00:00:00Z'; pythonPakete = 3; befunde = 1; signatur = 'abcdef0123456789'
            eintraege = @(
                [ordered]@{ id = 'gs'; version = '1.0'; neueste = '2.0'; status = 'luecke'; luecken = @([ordered]@{ id = 'CVE-A'; wert = '4.3'; stufe = 'MEDIUM'; aliase = ''; behoben = '' }, [ordered]@{ id = 'CVE-B'; wert = '9.8'; stufe = 'CRITICAL'; aliase = ''; behoben = '' }); hinweise = @('a|b'); fehler = '' },
                [ordered]@{ id = 'qpdf'; version = '1'; neueste = '1'; status = 'aktuell'; luecken = @(); hinweise = @(); fehler = '' }) }
        $t = ConvertTo-HUWBericht $e
        $t | Should -Match 'huwaechter:abcdef0123456789'
        $t | Should -Match '1 Befund'
        $t.IndexOf('CVE-B') | Should -BeLessThan $t.IndexOf('CVE-A')
        $t | Should -Not -Match 'a\|b'
    }
    It 'Teams ohne Webhook wirft nicht' { Send-HUWTeams '' 'x' @('y') | Should -Be 'kein Webhook' }
}

Describe 'waechter.json eines Moduls (Test-HUWKonfig)' {
    It 'nimmt eine vollständige Liste an' {
        $d = New-TestModul @{ modul = 'Test'; werkzeugpaket = 'werkzeuge.json'; pythonPakete = 'requirements.txt'
            eintraege = @(@{ id = 'gs'; neueste = @{ art = 'github'; repo = 'a/b' }; cpe = 'cpe:2.3:a:artifex:ghostscript' }, @{ id = 'lib'; version = '1.0'; osv = @{ oekosystem = 'NuGet'; name = 'X' } }) } `
            @{ dateien = @(@{ id = 'gs'; version = '10.08.0' }) } "x==1 --hash=sha256:aa`n"
        try { @(Test-HUWKonfig $d).Count | Should -Be 0 } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
    It 'meldet fehlende Datei' {
        $d = Join-Path ([IO.Path]::GetTempPath()) ('huw-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $d | Out-Null
        try { @(Test-HUWKonfig $d) -join ';' | Should -Match 'fehlt' } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
    It 'meldet Werkzeug ohne Eintrag, unbekannte Quelle, doppelte id, Ausnahme ohne Grund' {
        $d = New-TestModul @{ modul = 'Test'; werkzeugpaket = 'werkzeuge.json'
            eintraege = @(@{ id = 'a'; neueste = @{ art = 'ftp' } }, @{ id = 'a'; cpe = 'cpe:2.3:a:x:y'; ignorieren = @(@{ id = 'CVE-1'; grund = '' }) }) } `
            @{ dateien = @(@{ id = 'a'; version = '1' }, @{ id = 'b'; version = '2' }) }
        try {
            $p = @(Test-HUWKonfig $d) -join ';'
            $p | Should -Match 'b: kein Wächter-Eintrag'
            $p | Should -Match 'Quelle ftp'
            $p | Should -Match 'a: doppelt'
            $p | Should -Match 'Ausnahme ohne Kennung oder Grund'
        } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
    It 'meldet Eintrag ohne Version und ohne Quelle' {
        $d = New-TestModul @{ modul = ''; eintraege = @(@{ id = 'x' }) }
        try {
            $p = @(Test-HUWKonfig $d) -join ';'
            $p | Should -Match 'modul fehlt'
            $p | Should -Match 'x: keine Quelle'
            $p | Should -Match 'x: keine Version'
        } finally { Remove-Item -LiteralPath $d -Recurse -Force }
    }
}

Describe 'Workflow und Vorlage' {
    It 'Workflow des Kerns ist wiederverwendbar und nutzt die Skripte des Kerns' {
        $w = Get-Content -LiteralPath (Join-Path $script:Root '.github\workflows\waechter.yml') -Raw -Encoding UTF8
        $w | Should -Match 'workflow_call:'
        $w | Should -Match 'repository: ChiliApple/HU-Amt'
        $w | Should -Match 'kern/\.github/waechter/Waechter\.ps1'
        $w | Should -Not -Match '(?m)^\s+schedule:'
    }
    It 'Vorlage ruft den Workflow des Kerns auf' {
        $v = Get-Content -LiteralPath (Join-Path $script:Root '.github\waechter\Vorlage-Modul.yml') -Raw -Encoding UTF8
        $v | Should -Match 'uses: ChiliApple/HU-Amt/\.github/workflows/waechter\.yml@main'
        $v | Should -Match 'issues: write'
        $v | Should -Match 'secrets: inherit'
    }
}
