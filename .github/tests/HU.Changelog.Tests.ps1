#Requires -Version 5.1
# Pester 5: CHANGELOG lesen fuer das Fenster "Neu in dieser Version" (Functions\Core-Changelog.ps1, ab 1.0.4).

BeforeAll {
    Set-StrictMode -Version 1.0
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    . (Join-Path $script:Root 'Functions\Core-Changelog.ps1')
    $script:Text = @'
# Changelog

## v1.0.4

Stable-Release der Testversionen 1.0.3 bis 1.0.4 - "Thema".

### Outlook
- **Anhänge:** direkt aus der Mail.
  - Auch mehrere.
    - Dritte Ebene mit `Code`.
- **Zweiter Punkt:** Text.

## v1.0.3 (Test, 10.11.2026)

### Hilfe
- **Hilfe:** neu.

## v1.0.2

- **Alt:** alt.
'@
}

Describe 'Abschnitte' {
    It 'erkennt jede Version mit Zusatz, neueste zuerst' {
        $s = @(Get-HUChangelogSections $script:Text)
        $s.Count | Should -Be 3
        $s[0].Version | Should -Be ([version]'1.0.4')
        $s[1].Heading | Should -Be 'v1.0.3 (Test, 10.11.2026)'
        $s[2].Version | Should -Be ([version]'1.0.2')
    }
    It 'waehlt neuer als gesehen bis installiert' {
        $s = @(Get-HUChangelogSections $script:Text)
        @(Select-HUChangelogSections $s '1.0.2' '1.0.4' | ForEach-Object { "$($_.Version)" }) | Should -Be @('1.0.4', '1.0.3')
        @(Select-HUChangelogSections $s '1.0.3' '1.0.3').Count | Should -Be 0
        @(Select-HUChangelogSections $s '1.0.2' '1.0.3' | ForEach-Object { "$($_.Version)" }) | Should -Be @('1.0.3')
        @(Select-HUChangelogSections $s $null '1.0.4' | ForEach-Object { "$($_.Version)" }) | Should -Be @('1.0.4')
    }
}

Describe 'Bausteine' {
    It 'Kopf, Absatz, Punkte mit Ebenen, fett und Code' {
        $s = @(Get-HUChangelogSections $script:Text)[0]
        $b = @(ConvertTo-HUChangelogBlocks @($s.Lines))
        @($b | ForEach-Object { $_.Kind }) | Should -Be @('Para', 'Heading', 'Bullet', 'Bullet', 'Bullet', 'Bullet')
        @($b | Where-Object { $_.Kind -eq 'Bullet' } | ForEach-Object { $_.Level }) | Should -Be @(0, 1, 2, 0)
        $b[2].Parts[0].Kind | Should -Be 'Bold'
        $b[2].Parts[0].Text | Should -Be 'Anhänge:'
        @($b[4].Parts | Where-Object { $_.Kind -eq 'Code' }).Text | Should -Be 'Code'
    }
    It 'unvollstaendige Marken bleiben Text' {
        $p = @(ConvertTo-HUChangelogInline 'a **b ohne Ende')
        $p.Count | Should -Be 1
        $p[0].Kind | Should -Be 'Text'
    }
}

Describe 'Was der Benutzer noch nicht gesehen hat' {
    BeforeAll {
        $script:Dir = Join-Path ([IO.Path]::GetTempPath()) ("hu-cl-" + [guid]::NewGuid().ToString('N'))
        [void](New-Item -ItemType Directory -Path $script:Dir)
        $script:File = Join-Path $script:Dir 'CHANGELOG.md'
        [IO.File]::WriteAllText($script:File, $script:Text, (New-Object Text.UTF8Encoding $true))
    }
    AfterAll { Remove-Item -LiteralPath $script:Dir -Recurse -Force -ErrorAction SilentlyContinue }
    It 'zeigt nur Eintraege mit neuerer Version und bekanntem Stand' {
        $cur = @(
            [pscustomobject]@{ Key = 'HU-Amt'; Title = 'HU-Amt'; Version = '1.0.4'; Path = $script:File }
            [pscustomobject]@{ Key = 'addon:Neu'; Title = 'Neu'; Version = '1.0.4'; Path = $script:File }
            [pscustomobject]@{ Key = 'addon:Gleich'; Title = 'Gleich'; Version = '1.0.3'; Path = $script:File }
        )
        $n = @(Get-HUChangelogNews -Current $cur -Seen @{ 'HU-Amt' = '1.0.2'; 'addon:Gleich' = '1.0.3' })
        $n.Count | Should -Be 1
        $n[0].Title | Should -Be 'HU-Amt'
        @($n[0].Sections | ForEach-Object { "$($_.Version)" }) | Should -Be @('1.0.4', '1.0.3')
    }
    It 'fehlende Datei oder aelterer Stand: nichts' {
        $cur = @([pscustomobject]@{ Key = 'HU-Amt'; Title = 'HU-Amt'; Version = '1.0.4'; Path = 'C:\gibt\es\nicht.md' })
        @(Get-HUChangelogNews -Current $cur -Seen @{ 'HU-Amt' = '1.0.2' }).Count | Should -Be 0
        $cur2 = @([pscustomobject]@{ Key = 'HU-Amt'; Title = 'HU-Amt'; Version = '1.0.2'; Path = $script:File })
        @(Get-HUChangelogNews -Current $cur2 -Seen @{ 'HU-Amt' = '1.0.4' }).Count | Should -Be 0
    }
    It 'der echte CHANGELOG von HU-Amt laesst sich lesen und hat einen Abschnitt zur Version' {
        $ver = "$(([IO.File]::ReadAllText((Join-Path $script:Root 'Config\version.json'), [Text.Encoding]::UTF8) | ConvertFrom-Json).version)"
        $s = @(Get-HUChangelogSections ([IO.File]::ReadAllText((Join-Path $script:Root 'CHANGELOG.md'), [Text.Encoding]::UTF8)))
        @($s | Where-Object { "$($_.Version)" -eq $ver }).Count | Should -Be 1
        @(ConvertTo-HUChangelogBlocks @($s[0].Lines)).Count | Should -BeGreaterThan 2
    }
}
