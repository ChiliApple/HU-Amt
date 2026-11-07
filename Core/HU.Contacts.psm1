#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Contacts - Kontakte fuer Module (z. B. Mieter, Betriebe): einheitliches Kontakt-Objekt und vCard-Leser (.vcf).
.DESCRIPTION
    Kontakt-Objekt: Name, Firma, Mail, Telefon, Mobil, Strasse, Plz, Ort, Land, Quelle.
    vCard: 4.0 (RFC 6350), 3.0 und 2.1 (Outlook-Export). Zeilen werden entfaltet (CRLF + Leerzeichen/Tab),
    QUOTED-PRINTABLE mit CHARSET wird entschluesselt, Maskierungen (\n \, \; \\) aufgeloest. Mehrere Karten je Datei.
    Fuer vCard 2.1 gibt es keine RFC; Aufbau nach den gaengigen Exporten (unsicher bei exotischen Dateien).
    Die Outlook-Kontakte liefert Core\HU.Office.psm1 (Get-HUOutlookContacts) im selben Format.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7. Keine fremden Bibliotheken.
#>

Set-StrictMode -Version 1.0

<#
.SYNOPSIS
    Einheitliches Kontakt-Objekt anlegen. Name faellt auf die Firma zurueck, wenn er fehlt.
#>
function New-HUContact {
    param([string]$Name = '', [string]$Firma = '', [string]$Mail = '', [string]$Telefon = '', [string]$Mobil = '',
        [string]$Strasse = '', [string]$Plz = '', [string]$Ort = '', [string]$Land = '', [string]$Quelle = '')
    $t = { param($v) ("$v" -replace '\s+', ' ').Trim() }
    $n = & $t $Name; $f = & $t $Firma
    if (-not $n) { $n = $f }
    return [pscustomobject]@{
        Name = $n; Firma = $f; Mail = (& $t $Mail); Telefon = (& $t $Telefon); Mobil = (& $t $Mobil)
        Strasse = (& $t $Strasse); Plz = (& $t $Plz); Ort = (& $t $Ort); Land = (& $t $Land); Quelle = $Quelle
    }
}

# Bytes einer Datei als Text: UTF-8 (mit oder ohne BOM), sonst Windows-1252 (aeltere Exporte)
function ConvertFrom-HUContactBytes([byte[]]$Bytes) {
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($Bytes, 2, $Bytes.Length - 2) }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($Bytes, 2, $Bytes.Length - 2) }
    $start = $(if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { 3 } else { 0 })
    try { return (New-Object System.Text.UTF8Encoding($false, $true)).GetString($Bytes, $start, $Bytes.Length - $start) }
    catch { return [Text.Encoding]::GetEncoding(1252).GetString($Bytes) }
}

# ohne CHARSET: zuerst streng UTF-8, sonst Windows-1252 (Outlook-Exporte ohne Angabe)
function ConvertFrom-HUQuotedPrintable([string]$Text, [string]$Charset = '') {
    $bytes = New-Object System.Collections.Generic.List[byte]
    $i = 0
    while ($i -lt $Text.Length) {
        $ch = $Text[$i]
        if ($ch -eq '=' -and $i + 2 -lt $Text.Length -and "$($Text[$i+1])$($Text[$i+2])" -match '^[0-9A-Fa-f]{2}$') {
            $bytes.Add([Convert]::ToByte("$($Text[$i+1])$($Text[$i+2])", 16)); $i += 3; continue
        }
        foreach ($b in [Text.Encoding]::GetEncoding(1252).GetBytes([string]$ch)) { $bytes.Add($b) }
        $i++
    }
    $arr = $bytes.ToArray()
    if (-not $Charset) {
        try { return (New-Object System.Text.UTF8Encoding($false, $true)).GetString($arr) } catch { return [Text.Encoding]::GetEncoding(1252).GetString($arr) }
    }
    $enc = $null
    try { $enc = [Text.Encoding]::GetEncoding($Charset) } catch { $enc = [Text.Encoding]::UTF8 }
    return $enc.GetString($arr)
}

function ConvertFrom-HUVCardEscape([string]$Value) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Value.Length; $i++) {
        $c = $Value[$i]
        if ($c -eq '\' -and $i + 1 -lt $Value.Length) {
            $n = $Value[$i + 1]
            switch -CaseSensitive ($n) { 'n' { [void]$sb.Append("`n") } 'N' { [void]$sb.Append("`n") } default { [void]$sb.Append($n) } }
            $i++; continue
        }
        [void]$sb.Append($c)
    }
    return $sb.ToString()
}

# Strukturierten Wert an ; teilen (maskierte \; bleiben im Teil)
function Split-HUVCardValue([string]$Value) {
    $parts = New-Object System.Collections.Generic.List[string]
    $cur = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Value.Length; $i++) {
        $c = $Value[$i]
        if ($c -eq '\' -and $i + 1 -lt $Value.Length) { [void]$cur.Append($c).Append($Value[$i + 1]); $i++; continue }
        if ($c -eq ';') { $parts.Add($cur.ToString()); [void]$cur.Clear(); continue }
        [void]$cur.Append($c)
    }
    $parts.Add($cur.ToString())
    return @($parts | ForEach-Object { ConvertFrom-HUVCardEscape $_ })
}

# Weiche QP-Umbrueche zusammenfuegen (Text mit `n): nur, wenn die Eigenschaft der Zeile Quoted-Printable ist.
# Eingerueckte Folgezeilen (RFC-Faltung) gehoeren zur Eigenschaft davor.
function Join-HUVCardSoftBreaks([string]$Raw) {
    $out = New-Object System.Collections.Generic.List[string]
    $buf = $null; $lastQp = $false
    foreach ($line in $Raw -split "`n") {
        if ($null -ne $buf) { $buf += $line }
        else {
            $buf = $line
            if ($line -notmatch '^[ \t]') {
                $ci = $line.IndexOf(':')
                $lastQp = ($ci -gt 0 -and $line.Substring(0, $ci) -match '(?i)(^|;)\s*(encoding\s*=\s*)?quoted-printable\s*(;|$)')
            }
        }
        if ($lastQp -and $buf.EndsWith('=')) { $buf = $buf.Substring(0, $buf.Length - 1); continue }
        $out.Add($buf); $buf = $null
    }
    if ($null -ne $buf) { $out.Add($buf) }
    return ($out -join "`n")
}

<#
.SYNOPSIS
    vCard-Text (eine oder mehrere Karten) in Kontakt-Objekte umwandeln.
#>
function ConvertFrom-HUVCard {
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text, [string]$Quelle = 'vCard')
    $raw = $Text -replace "`r`n", "`n" -replace "`r", "`n"
    # Quoted-Printable: weiche Umbrueche (= am Zeilenende) zuerst zusammenfuegen, aber nur bei Eigenschaften mit
    # ENCODING=QUOTED-PRINTABLE (KERN-B2). Base64 (z. B. PHOTO) endet oft auf = und darf nicht mit der Folgezeile verschmelzen.
    $raw = Join-HUVCardSoftBreaks $raw
    # Entfalten: Zeilenumbruch + ein Leerzeichen/Tab (RFC 6350 3.2)
    $raw = [regex]::Replace($raw, "`n[ `t]", '')
    $out = New-Object System.Collections.Generic.List[object]
    $card = $null
    foreach ($line in $raw -split "`n") {
        if (-not $line.Trim()) { continue }
        if ($line -match '^(?i)BEGIN:VCARD\s*$') { $card = @{ FN = ''; N = @(); ORG = ''; EMAIL = @(); TEL = @(); ADR = @() }; continue }
        if ($line -match '^(?i)END:VCARD\s*$') {
            if ($card) {
                $name = $card.FN
                if (-not $name -and $card.N.Count) { $name = (@($card.N[1], $card.N[2], $card.N[0]) | Where-Object { $_ }) -join ' ' }
                $mail = @($card.EMAIL | Sort-Object { -not $_.Pref } | ForEach-Object { $_.Value })[0]
                $tels = @($card.TEL)
                $mob = @($tels | Where-Object { $_.Types -match 'cell|mobile' } | ForEach-Object { $_.Value })[0]
                $tel = @($tels | Where-Object { $_.Types -notmatch 'cell|mobile|fax' } | Sort-Object { -not $_.Pref }, { $_.Types -notmatch 'work' } | ForEach-Object { $_.Value })[0]
                $adr = @($card.ADR | Sort-Object { -not $_.Pref })[0]
                $p = $(if ($adr) { $adr.Parts } else { @() })
                $g = { param($i) if ($p.Count -gt $i) { $p[$i] } else { '' } }
                $c = New-HUContact -Name $name -Firma $card.ORG -Mail $mail -Telefon $tel -Mobil $mob -Strasse ((@((& $g 1), (& $g 2)) | Where-Object { $_ }) -join ' ') -Plz (& $g 5) -Ort (& $g 3) -Land (& $g 6) -Quelle $Quelle
                if ($c.Name -or $c.Mail -or $c.Telefon) { $out.Add($c) }
            }
            $card = $null; continue
        }
        if (-not $card) { continue }
        $ci = $line.IndexOf(':'); if ($ci -lt 1) { continue }
        $head = $line.Substring(0, $ci); $val = $line.Substring($ci + 1)
        $hp = $head.Split(';')
        $prop = $hp[0].ToUpperInvariant(); if ($prop.Contains('.')) { $prop = $prop.Substring($prop.LastIndexOf('.') + 1) }   # Gruppe item1.EMAIL
        $params = @($hp | Select-Object -Skip 1)
        $ptext = ($params -join ';').ToLowerInvariant()
        if ($ptext -match 'encoding=quoted-printable' -or $params -contains 'QUOTED-PRINTABLE') {
            $cs = ''; foreach ($pp in $params) { if ($pp -match '^(?i)charset=(.+)$') { $cs = $Matches[1] } }
            $val = ConvertFrom-HUQuotedPrintable $val $cs
        }
        $pref = ($ptext -match 'pref')
        switch ($prop) {
            'FN' { $card.FN = (ConvertFrom-HUVCardEscape $val).Trim() }
            'N' { $card.N = @(Split-HUVCardValue $val) }
            'ORG' { $card.ORG = (@(Split-HUVCardValue $val) | Where-Object { $_ } | Select-Object -First 1) }
            'EMAIL' { $card.EMAIL += [pscustomobject]@{ Value = (ConvertFrom-HUVCardEscape $val).Trim(); Pref = $pref } }
            'TEL' { $card.TEL += [pscustomobject]@{ Value = (ConvertFrom-HUVCardEscape $val).Trim() -replace '^tel:', ''; Types = $ptext; Pref = $pref } }
            'ADR' { $card.ADR += [pscustomobject]@{ Parts = @(Split-HUVCardValue $val); Pref = $pref } }
            default { }
        }
    }
    return $out.ToArray()
}

<#
.SYNOPSIS
    vCard-Datei (.vcf) lesen. Rueckgabe: Kontakt-Objekte (leer, wenn keine Karte darin ist).
#>
function Import-HUVCardFile {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Datei nicht gefunden: $Path" }
    $t = ConvertFrom-HUContactBytes ([IO.File]::ReadAllBytes($Path))
    return @(ConvertFrom-HUVCard -Text $t -Quelle ("vCard " + [IO.Path]::GetFileName($Path)))
}

<#
.SYNOPSIS
    Kontakte nach Suchtext filtern (Name, Firma, Mail, Ort; Gross/klein egal).
#>
function Find-HUContact {
    param([object[]]$Contacts = @(), [string]$Search = '')
    $s = "$Search".Trim()
    if (-not $s) { return @($Contacts) }
    # IndexOf statt -like: Zeichen wie [ ] * ? im Suchtext sind kein Platzhalter (KERN-B7)
    return @($Contacts | Where-Object { "$($_.Name) $($_.Firma) $($_.Mail) $($_.Ort)".IndexOf($s, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
}

Export-ModuleMember -Function New-HUContact, ConvertFrom-HUVCard, Import-HUVCardFile, Find-HUContact
