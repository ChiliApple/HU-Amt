#Requires -Version 5.1
<#
.SYNOPSIS
    CHANGELOG.md lesen (ab 1.0.4): Abschnitte je Version und ihre Bausteine fuer das Fenster "Neu in dieser Version".
.DESCRIPTION
    Reine Funktionen ohne Oberflaeche (getestet in .github\tests\HU.Changelog.Tests.ps1).
    - Abschnitt beginnt mit "## v1.2.3" (Rest der Zeile ist Zusatz, z. B. "(Stable, 11.11.2026)").
    - Bausteine: Kopf (### ...), Absatz, Punkt mit Ebene (Einrueckung je 2 Leerzeichen), Inhalt mit **fett** und `Code`.
    - Gezeigt werden die Versionen, die neuer sind als die zuletzt gesehene, hoechstens bis zur installierten.
#>

function ConvertTo-HUChangelogVersion([string]$Text) {
    $m = [regex]::Match("$Text", '(\d+)\.(\d+)\.(\d+)')
    if (-not $m.Success) { return $null }
    try { return [version]$m.Value } catch { return $null }
}

<#
.SYNOPSIS
    Abschnitte aus dem Text eines CHANGELOG: Version ([version]), Heading (Text nach "## "), Lines (Zeilen darunter).
    Reihenfolge wie in der Datei (neueste zuerst). Zeilen vor dem ersten Abschnitt (Titel "# Changelog") fallen weg.
#>
function Get-HUChangelogSections([string]$Text) {
    $out = New-Object System.Collections.Generic.List[object]
    $cur = $null
    foreach ($raw in @("$Text" -split "\r?\n")) {
        $line = $raw.TrimEnd()
        $hm = [regex]::Match($line, '^##\s+(.+)$')
        if ($hm.Success) {
            $v = ConvertTo-HUChangelogVersion $hm.Groups[1].Value
            if ($v) {
                $cur = [pscustomobject]@{ Version = $v; Heading = $hm.Groups[1].Value.Trim(); Lines = (New-Object System.Collections.Generic.List[string]) }
                $out.Add($cur)
                continue
            }
        }
        if ($cur) { $cur.Lines.Add($line) }
    }
    return $out.ToArray()
}

<#
.SYNOPSIS
    Abschnitte neuer als -Since und hoechstens -Upto (beides [version] oder Text). -Since leer: nur der Abschnitt von -Upto.
#>
function Select-HUChangelogSections([object[]]$Sections, $Since, $Upto) {
    $s = $(if ($Since -is [version]) { $Since } else { ConvertTo-HUChangelogVersion "$Since" })
    $u = $(if ($Upto -is [version]) { $Upto } else { ConvertTo-HUChangelogVersion "$Upto" })
    return @(foreach ($x in @($Sections)) {
            if (-not $x) { continue }
            if ($u -and $x.Version -gt $u) { continue }
            if ($s) { if ($x.Version -gt $s) { $x } }
            elseif ($u -and $x.Version -eq $u) { $x }
        })
}

<#
.SYNOPSIS
    Inhalt einer Zeile in Stuecke: Text, Bold (**...**), Code (`...`). Unvollstaendige Marken bleiben Text.
#>
function ConvertTo-HUChangelogInline([string]$Text) {
    $out = New-Object System.Collections.Generic.List[object]
    $rest = "$Text"
    while ($rest.Length) {
        $m = [regex]::Match($rest, '\*\*(.+?)\*\*|`([^`]+)`')
        if (-not $m.Success) { $out.Add([pscustomobject]@{ Kind = 'Text'; Text = $rest }); break }
        if ($m.Index -gt 0) { $out.Add([pscustomobject]@{ Kind = 'Text'; Text = $rest.Substring(0, $m.Index) }) }
        if ($m.Groups[1].Success) { $out.Add([pscustomobject]@{ Kind = 'Bold'; Text = $m.Groups[1].Value }) }
        else { $out.Add([pscustomobject]@{ Kind = 'Code'; Text = $m.Groups[2].Value }) }
        $rest = $rest.Substring($m.Index + $m.Length)
    }
    return $out.ToArray()
}

<#
.SYNOPSIS
    Zeilen eines Abschnitts in Bausteine: Kind = Heading | Para | Bullet, Level (Punkte: 0, 1, 2 ...), Parts (Inline-Stuecke).
    Fortsetzungszeilen eines Absatzes werden angehaengt, Leerzeilen trennen.
#>
function ConvertTo-HUChangelogBlocks([string[]]$Lines) {
    $out = New-Object System.Collections.Generic.List[object]
    $para = $null
    foreach ($l in @($Lines)) {
        $line = "$l"
        if (-not $line.Trim()) { $para = $null; continue }
        if ($line -match '^###\s+(.+)$') {
            $out.Add([pscustomobject]@{ Kind = 'Heading'; Level = 0; Parts = @(ConvertTo-HUChangelogInline $Matches[1].Trim()) }); $para = $null; continue
        }
        if ($line -match '^(\s*)[-*]\s+(.+)$') {
            $lvl = [int][Math]::Floor($Matches[1].Replace("`t", '  ').Length / 2)
            $out.Add([pscustomobject]@{ Kind = 'Bullet'; Level = $lvl; Parts = @(ConvertTo-HUChangelogInline $Matches[2].Trim()) }); $para = $null; continue
        }
        if ($para) { $para.Text = "$($para.Text) $($line.Trim())"; $para.Parts = @(ConvertTo-HUChangelogInline $para.Text); continue }
        $para = [pscustomobject]@{ Kind = 'Para'; Level = 0; Text = $line.Trim(); Parts = @(ConvertTo-HUChangelogInline $line.Trim()) }
        $out.Add($para)
    }
    return $out.ToArray()
}

<#
.SYNOPSIS
    Was dieser Benutzer seit dem letzten Start noch nicht gesehen hat.
    -Current: Liste @{ Key; Title; Version; Path } (Programm und Module, Path = CHANGELOG.md).
    -Seen: Hashtable Key -> zuletzt gesehene Version (Text). Fehlt ein Schluessel (neues Modul), wird nichts gezeigt.
    Rueckgabe: je Eintrag mit neuen Abschnitten Title, Key, Version, Sections.
#>
function Get-HUChangelogNews([object[]]$Current, [hashtable]$Seen) {
    return @(foreach ($c in @($Current)) {
            if (-not $c -or -not "$($c.Key)") { continue }
            if (-not $Seen.ContainsKey("$($c.Key)")) { continue }
            $old = ConvertTo-HUChangelogVersion "$($Seen["$($c.Key)"])"
            $new = ConvertTo-HUChangelogVersion "$($c.Version)"
            if (-not $old -or -not $new -or $new -le $old) { continue }
            if (-not "$($c.Path)" -or -not (Test-Path -LiteralPath "$($c.Path)")) { continue }
            $txt = ''; try { $txt = [IO.File]::ReadAllText("$($c.Path)", [Text.Encoding]::UTF8) } catch { continue }
            $sec = @(Select-HUChangelogSections (Get-HUChangelogSections $txt) $old $new)
            if ($sec.Count) { [pscustomobject]@{ Title = "$($c.Title)"; Key = "$($c.Key)"; Version = "$($c.Version)"; Sections = $sec } }
        })
}
