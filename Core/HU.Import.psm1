#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Import - Excel (.xlsx/.xlsm) und CSV lesen sowie einfache xlsx schreiben, ohne Excel und ohne fremde Bibliotheken.
.DESCRIPTION
    Lesen (Read-HUXlsx, Read-HUCsv, Read-HUImportFile) liefert Blaetter als [pscustomobject]:
        Name     Name des Blatts ('CSV' bei CSV)
        Rows     Liste von string[] je Zeile der Datei; Zeilen, die in der Datei fehlen, kommen als leeres string[]
    Standard-Werte der Zellen: Text wie in der Datei, Zahlen mit Punkt (ganze Zahlen ohne Nachkommastellen),
    Wahrheitswerte WAHR/FALSCH, Fehler leer. Ein Datum steht in einer xlsx als Excel-Seriennummer (Zahl).
    Mit -SheetJs liest Read-HUXlsx wie SheetJS (XLSX.read + sheet_to_json mit defval ''), fuer Module, die
    dieselben Werte wie eine JavaScript-Auswertung brauchen. Dann gilt fuer Rows: Zahlen wie String(zahl) in
    JavaScript, Wahrheitswerte true/false, Fehler als Text (#N/A). Zusaetzliche Felder:
        Ref      Bereich @(sr, sc, er, ec) (0-basiert) wie !ref: <dimension> der Datei, sonst aus den Zellen
        Werte    Zeile -> (Spalte -> Text) nur fuer Zellen mit Wert
        Anzeige  'zeile,spalte' -> angezeigter Text (fuer Spaltennamen)
    ConvertTo-HUXlsxRowObject macht daraus Zeilen-Objekte wie sheet_to_json.
    CSV: Kodierung UTF-8 mit BOM, UTF-16 mit BOM, UTF-8 ohne BOM, sonst Windows-1252. Trennzeichen ; , oder Tabulator
    (das haeufigste in der ersten Zeile, Standard ;), Felder mit Anfuehrungszeichen auch ueber mehrere Zeilen.
    Schreiben (New-HUXlsxBytes): Blaetter mit Kopfzeile (fett, fixiert, Autofilter) und Textzellen.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7, auch im Hintergrund-Runspace ohne Fenster.
    XML wird ohne DTD gelesen (DtdProcessing Prohibit, kein XmlResolver); eine xlsx enthaelt nie eine DTD.
#>

Set-StrictMode -Version 1.0
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Numerics

$script:HUImportMaxZeilen = 200000
$script:HUImportNsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# --- Hilfen -----------------------------------------------------------------------------------------
# Text eines Eintrags im ZIP (UTF-8) oder $null
function Get-HUImportZipText($Zip, [string]$Name) {
    $e = $Zip.GetEntry($Name)
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}
# XML ohne DTD laden; -Leerraum: Leerzeichen in Zellen bleiben erhalten (wie SheetJS)
function ConvertTo-HUImportXml([string]$Text, [switch]$Leerraum) {
    $set = New-Object Xml.XmlReaderSettings
    $set.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $set.XmlResolver = $null
    $d = New-Object Xml.XmlDocument
    $d.PreserveWhitespace = [bool]$Leerraum
    $d.XmlResolver = $null
    $sr = New-Object IO.StringReader($Text)
    $xr = [Xml.XmlReader]::Create($sr, $set)
    try { $d.Load($xr) } finally { $xr.Dispose(); $sr.Dispose() }
    return $d
}
# Text aus <t>-Knoten wie SheetJS (Rich Text zusammengesetzt, Lautschrift rPh weggelassen, _xHHHH_ aufgeloest)
function Get-HUImportXmlText($Node) {
    if (-not $Node) { return '' }
    $sb = New-Object Text.StringBuilder
    foreach ($t in $Node.SelectNodes("./*[local-name()='t'] | ./*[local-name()='r']/*[local-name()='t']")) { [void]$sb.Append($t.InnerText) }
    return [regex]::Replace($sb.ToString(), '_x([0-9A-Fa-f]{4})_', { param($mm) [string][char][Convert]::ToInt32($mm.Groups[1].Value, 16) })
}
# Zahl aus der xlsx als Text (Standard): ganze Zahlen ohne Nachkommastellen, sonst Rundreise-Darstellung mit Punkt
function ConvertTo-HUImportZahlText([string]$Raw) {
    $d = 0.0
    if (-not [double]::TryParse($Raw, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $Raw }
    if ($d -eq [Math]::Floor($d) -and [Math]::Abs($d) -lt 1e15) { return ([long]$d).ToString([Globalization.CultureInfo]::InvariantCulture) }
    return $d.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
}
# Text fuer XML: unerlaubte Zeichen entfernen, maskieren
function ConvertTo-HUImportXmlText([string]$Text) {
    $t = [regex]::Replace("$Text", '[^\u0009\u000A\u000D -퟿-�]', '')
    return [Security.SecurityElement]::Escape($t)
}

<#
.SYNOPSIS
    Spaltenbuchstaben einer Zellangabe ("C7", "AA1") als Index (0-basiert). Ohne Buchstaben: -1.
#>
function Get-HUXlsxColumnIndex([string]$Ref) {
    $n = 0
    foreach ($ch in $Ref.ToUpperInvariant().ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { break }
        $n = $n * 26 + ([int]$ch - 64)
    }
    return $n - 1
}
<#
.SYNOPSIS
    Spaltenbuchstaben zu einem Index (0-basiert): 0 = A, 26 = AA.
#>
function Get-HUXlsxColumnName([int]$Index) {
    $s = ''; $n = $Index + 1
    while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }
    return $s
}
<#
.SYNOPSIS
    Bereich "A1:D6" -> @(sr, sc, er, ec) (0-basiert) oder $null, wie SheetJS (auch $-Zeichen, fehlende Teile = 0).
#>
function ConvertFrom-HUXlsxRange([string]$Ref) {
    $m = [regex]::Match("$Ref", '^\$?([A-Za-z]*)\$?(\d*):\$?([A-Za-z]*)\$?(\d*)$')
    if (-not $m.Success) { return $null }
    $sc = $(if ($m.Groups[1].Value) { Get-HUXlsxColumnIndex $m.Groups[1].Value } else { 0 })
    $sr = $(if ($m.Groups[2].Value) { [int]$m.Groups[2].Value - 1 } else { 0 })
    $ec = $(if ($m.Groups[3].Value) { Get-HUXlsxColumnIndex $m.Groups[3].Value } else { 0 })
    $er = $(if ($m.Groups[4].Value) { [int]$m.Groups[4].Value - 1 } else { 0 })
    if ($ec -lt $sc -or $er -lt $sr -or $ec -lt 0 -or $er -lt 0) { return $null }
    return @($sr, $sc, $er, $ec)
}

<#
.SYNOPSIS
    Zahl als Text genau wie String(zahl) in JavaScript (kuerzeste eindeutige Ziffernfolge, ab 1e21 und unter 1e-6
    mit Exponent, z. B. 1e+21, 1.5e-7).
.DESCRIPTION
    Rechnet exakt mit BigInteger, weil Windows PowerShell 5.1 (.NET Framework) Zahlen mit mehr als 15 Stellen
    nicht verlaesslich formatiert.
#>
function ConvertTo-HUJsNumberText([double]$D) {
    if ([double]::IsNaN($D)) { return 'NaN' }
    if ([double]::IsInfinity($D)) { return $(if ($D -gt 0) { 'Infinity' } else { '-Infinity' }) }
    if ($D -eq 0) { return '0' }
    $a = [Math]::Abs($D); $vz = $(if ($D -lt 0) { '-' } else { '' })
    if ($a -lt 9007199254740992 -and $a -eq [Math]::Floor($a)) { return $vz + ([Numerics.BigInteger]$a).ToString() }   # ganze Zahl, exakt
    $bits = [BitConverter]::DoubleToInt64Bits($a)
    $be = [int](($bits -shr 52) -band 0x7FF); $fr = $bits -band 0xFFFFFFFFFFFFF
    if ($be -eq 0) { $m = [Numerics.BigInteger]$fr; $e = -1074 } else { $m = [Numerics.BigInteger]($fr -bor 0x10000000000000); $e = $be - 1075 }
    # Wert und Rundungsgrenzen als Bruch Z/N (vierfach, damit die halben Abstaende ganzzahlig bleiben)
    $p2 = [Numerics.BigInteger]::Pow(2, [Math]::Abs($e))
    $N = $(if ($e -lt 0) { 4 * $p2 } else { [Numerics.BigInteger]4 })
    $f = $(if ($e -lt 0) { [Numerics.BigInteger]1 } else { $p2 })
    $V = 4 * $m * $f
    $lo = $(if ($fr -eq 0 -and $be -gt 1) { (4 * $m - 1) * $f } else { (4 * $m - 2) * $f })
    $hi = (4 * $m + 2) * $f
    $gleich = $m.IsEven   # Grenzen gehoeren dazu, wenn die Mantisse gerade ist (Runden auf gerade)
    # Zehnerexponent k10 = floor(log10(a)), geschaetzt und exakt nachgeprueft
    $p10 = { param([int]$x) [Numerics.BigInteger]::Pow(10, $x) }
    $abZehn = { param([int]$k) if ($k -ge 0) { $V -ge (& $p10 $k) * $N } else { $V * (& $p10 (-$k)) -ge $N } }   # a >= 10^k ?
    $k10 = [int][Math]::Floor([Math]::Log10($a))
    while (-not (& $abZehn $k10)) { $k10-- }
    while (& $abZehn ($k10 + 1)) { $k10++ }
    $C = [Numerics.BigInteger]0; $t = 0
    for ($p = 1; $p -le 17; $p++) {
        $t = $k10 - $p + 1          # Kandidat = C * 10^t
        if ($t -le 0) { $zz = $V * (& $p10 (-$t)); $nn = $N } else { $zz = $V; $nn = $N * (& $p10 $t) }
        # auf p Stellen gerundet, bei genau der Haelfte auf die gerade Ziffer (wie JavaScript)
        $rest = [Numerics.BigInteger]::Zero
        $C = [Numerics.BigInteger]::DivRem($zz, $nn, [ref]$rest)
        $vgl = (2 * $rest).CompareTo($nn)
        if ($vgl -gt 0 -or ($vgl -eq 0 -and -not $C.IsEven)) { $C = $C + 1 }
        # liegt der Kandidat im Rundungsbereich der Zahl? (alles mit N bzw. 10^-t erweitert)
        if ($t -le 0) { $s = & $p10 (-$t); $vk = $C * $N; $vlo = $lo * $s; $vhi = $hi * $s }
        else { $vk = $C * (& $p10 $t) * $N; $vlo = $lo; $vhi = $hi }
        $ok = $(if ($gleich) { $vk -ge $vlo -and $vk -le $vhi } else { $vk -gt $vlo -and $vk -lt $vhi })
        if ($ok) { break }
    }
    $ziff = $C.ToString(); $n = $ziff.Length + $t
    $ziff = $ziff.TrimEnd('0'); $kz = $ziff.Length
    if ($kz -le $n -and $n -le 21) { $txt = $ziff + ('0' * ($n - $kz)) }
    elseif ($n -gt 0 -and $n -le 21) { $txt = $ziff.Substring(0, $n) + '.' + $ziff.Substring($n) }
    elseif ($n -gt -6 -and $n -le 0) { $txt = '0.' + ('0' * (-$n)) + $ziff }
    else {
        $ex = $n - 1
        $txt = $ziff.Substring(0, 1) + $(if ($kz -gt 1) { '.' + $ziff.Substring(1) } else { '' }) + 'e' + $(if ($ex -ge 0) { '+' } else { '-' }) + [Math]::Abs($ex)
    }
    return $vz + $txt
}

# --- Excel lesen ------------------------------------------------------------------------------------
<#
.SYNOPSIS
    Blaetter einer xlsx/xlsm lesen (ohne Excel).
.PARAMETER Bytes
    Inhalt der Datei.
.PARAMETER SheetJs
    Lesen wie SheetJS: Werte wie in JavaScript, dazu Ref, Werte, Anzeige (siehe Kopf der Datei).
.PARAMETER MaxRows
    Hoechstzahl Zeilen je Blatt (danach wird abgebrochen). Standard: 200000, mit -SheetJs ohne Grenze. 0 = ohne Grenze.
.OUTPUTS
    object[] der Blaetter (Name, Rows[, Ref, Werte, Anzeige]).
#>
function Read-HUXlsx {
    param([Parameter(Mandatory, Position = 0)][byte[]]$Bytes, [switch]$SheetJs, [int]$MaxRows = -1)
    $grenze = $(if ($MaxRows -ge 0) { $MaxRows } elseif ($SheetJs) { 0 } else { $script:HUImportMaxZeilen })
    $ms = New-Object IO.MemoryStream(, $Bytes)
    $zip = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $shared = @()
        $sst = Get-HUImportZipText $zip 'xl/sharedStrings.xml'
        if ($sst) { $shared = @((ConvertTo-HUImportXml $sst -Leerraum).DocumentElement.SelectNodes("./*[local-name()='si']") | ForEach-Object { Get-HUImportXmlText $_ }) }
        $wbText = Get-HUImportZipText $zip 'xl/workbook.xml'
        $relText = Get-HUImportZipText $zip 'xl/_rels/workbook.xml.rels'
        if (-not $wbText -or -not $relText) { throw 'Die Datei ist keine gültige Excel-Datei (Arbeitsmappe fehlt).' }
        $wb = ConvertTo-HUImportXml $wbText
        $rels = ConvertTo-HUImportXml $relText
        $ziele = @{}
        foreach ($r in $rels.DocumentElement.SelectNodes("./*[local-name()='Relationship']")) { $ziele[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }
        $out = New-Object Collections.Generic.List[object]
        foreach ($s in $wb.DocumentElement.SelectNodes("./*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rid = $s.GetAttribute('id', $script:HUImportNsRel)
            $t = "$($ziele[$rid])"
            if (-not $t) { continue }
            $pfad = $(if ($t.StartsWith('/')) { $t.TrimStart('/') } else { 'xl/' + $t })
            $xml = Get-HUImportZipText $zip $pfad
            if (-not $xml) { continue }
            $rows = New-Object Collections.Generic.List[object]
            $werte = $null; $anzeige = $null
            if ($SheetJs) {
                $werte = New-Object 'Collections.Generic.SortedDictionary[int,object]'
                $anzeige = New-Object 'Collections.Generic.Dictionary[string,string]'
            }
            $g = @(2000000, 2000000, -1, -1)   # Bereich aus den Zellen (wie SheetJS: Zeilen mit Inhalt, Spalten mit Wert oder Typ)
            $ri = -1
            foreach ($row in (ConvertTo-HUImportXml $xml -Leerraum).DocumentElement.SelectNodes("./*[local-name()='sheetData']/*[local-name()='row']")) {
                # leere Zeilen stehen nicht in der Datei: mit der Zeilennummer auffuellen, damit Luecken erhalten bleiben
                $rn = 0
                if ([int]::TryParse($row.GetAttribute('r'), [ref]$rn)) { $ri = $rn - 1; while ($rows.Count -lt $ri) { $rows.Add([string[]]@()) } } else { $ri++ }
                if ($grenze -gt 0 -and $rows.Count -ge $grenze) { break }
                if ($SheetJs -and -not $row.IsEmpty) { if ($ri -lt $g[0]) { $g[0] = $ri }; if ($ri -gt $g[2]) { $g[2] = $ri } }
                $text = @{}; $max = -1; $next = 0
                foreach ($c in $row.SelectNodes("./*[local-name()='c']")) {
                    $ref = $c.GetAttribute('r')
                    $i = $(if ($ref) { Get-HUXlsxColumnIndex $ref } else { $next })
                    $next = $i + 1
                    $typ = $c.GetAttribute('t')
                    $vn = $c.SelectSingleNode("./*[local-name()='v']")
                    $vt = $(if ($vn) { $vn.InnerText } else { '' })
                    if ($SheetJs) {
                        # $wert = Wert fuer sheet_to_json ($null = kein Wert), $zeig = angezeigter Text
                        $wert = $null; $zeig = ''
                        switch ($typ) {
                            'inlineStr' { $wert = Get-HUImportXmlText $c.SelectSingleNode("./*[local-name()='is']"); $zeig = $wert }
                            's' { $j = 0; $wert = $(if ([int]::TryParse($vt, [ref]$j) -and $j -ge 0 -and $j -lt $shared.Count) { "$($shared[$j])" } else { '' }); $zeig = $wert }
                            'str' { $wert = $vt; $zeig = $vt }
                            'b' { if ($vt -ne '') { $wert = $(if ($vt -eq '1' -or $vt -eq 'true') { 'true' } else { 'false' }); $zeig = $wert.ToUpperInvariant() } }
                            'e' { $zeig = $vt }
                            default {
                                if ($vt -ne '') {
                                    $d = 0.0
                                    $wert = $(if ([double]::TryParse($vt, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { ConvertTo-HUJsNumberText $d } else { $vt })
                                    $zeig = $wert
                                }
                            }
                        }
                        $stub = (-not $typ) -and $vt -eq ''   # nur Format, kein Wert: zaehlt bei SheetJS nicht
                        if (-not $stub) {
                            if ($i -lt $g[1]) { $g[1] = $i }; if ($i -gt $g[3]) { $g[3] = $i }
                            $anzeige["$ri,$i"] = $zeig
                        }
                        if ($null -ne $wert) {
                            if (-not $werte.ContainsKey($ri)) { $werte[$ri] = New-Object 'Collections.Generic.Dictionary[int,string]' }
                            $werte[$ri][$i] = $wert
                        }
                        $text[$i] = $(if ($typ -eq 'e') { $vt } elseif ($null -ne $wert) { $wert } else { '' })
                    } else {
                        $text[$i] = "$(switch ($typ) {
                            'inlineStr' { Get-HUImportXmlText $c.SelectSingleNode("./*[local-name()='is']") }
                            's' { $j = 0; if ([int]::TryParse($vt, [ref]$j) -and $j -ge 0 -and $j -lt $shared.Count) { "$($shared[$j])" } else { '' } }
                            'str' { $vt }
                            'b' { if ($vt -eq '1' -or $vt -eq 'true') { 'WAHR' } elseif ($vt -ne '') { 'FALSCH' } else { '' } }
                            'e' { '' }
                            default { if ($vt -ne '') { ConvertTo-HUImportZahlText $vt } else { '' } }
                        })"
                    }
                    if ($i -gt $max) { $max = $i }
                }
                $arr = New-Object string[] ($max + 1)
                for ($k = 0; $k -le $max; $k++) { $arr[$k] = $(if ($text.ContainsKey($k)) { "$($text[$k])" } else { '' }) }
                $rows.Add($arr)
            }
            if (-not $SheetJs) {
                $out.Add([pscustomobject]@{ Name = $s.GetAttribute('name'); Rows = $rows })
                continue
            }
            # Bereich: <dimension ref="A1:D6"> der Datei gilt (auch wenn kleiner oder groesser als die Daten), sonst aus den Zellen
            $bereich = $null
            $dm = [regex]::Match($xml, '<(?:\w*:)?dimension')
            if ($dm.Success -and $dm.Index -gt 0) {
                $rm = [regex]::Match($xml.Substring($dm.Index, [Math]::Min(50, $xml.Length - $dm.Index)), '"(\w*:\w*)"')
                if ($rm.Success) { $bereich = ConvertFrom-HUXlsxRange $rm.Groups[1].Value }
            }
            if (-not $bereich -and $g[3] -ge $g[1] -and $g[2] -ge $g[0]) { $bereich = @($g[0], $g[1], $g[2], $g[3]) }
            $out.Add([pscustomobject]@{ Name = $s.GetAttribute('name'); Rows = $rows; Ref = $bereich; Werte = $werte; Anzeige = $anzeige })
        }
        return $out.ToArray()
    } finally { $zip.Dispose(); $ms.Dispose() }
}

<#
.SYNOPSIS
    Zeilen eines mit Read-HUXlsx -SheetJs gelesenen Blatts als Objekte wie sheet_to_json(ws, { defval: '' }).
.DESCRIPTION
    Kopf = erste Zeile des Bereichs (leere Zellen heissen __EMPTY, __EMPTY_1, ...; doppelte Namen Name_1, Name_2, ...),
    jede Spalte des Bereichs mit '' vorbelegt, Zeilen ohne Wert fallen weg, Spaltennamen unterscheiden Gross/Klein.
.OUTPUTS
    Hashtables (Ordinal-Vergleich der Schluessel) einzeln in die Pipeline, wie bisher in den Modulen; ohne Bereich nichts.
#>
function ConvertTo-HUXlsxRowObject {
    param([Parameter(Position = 0)]$Sheet)
    $list = New-Object Collections.Generic.List[hashtable]
    if (-not $Sheet -or -not $Sheet.Ref) { return $list }
    $sr = $Sheet.Ref[0]; $sc = $Sheet.Ref[1]; $er = $Sheet.Ref[2]; $ec = $Sheet.Ref[3]
    $zaehler = New-Object 'Collections.Generic.Dictionary[string,int]' ([StringComparer]::Ordinal)
    $kopf = New-Object string[] ($ec - $sc + 1)
    for ($c = $sc; $c -le $ec; $c++) {
        $v = $(if ($Sheet.Anzeige.ContainsKey("$sr,$c")) { $Sheet.Anzeige["$sr,$c"] } else { '__EMPTY' })
        $vv = $v
        if (-not $zaehler.ContainsKey($v) -or $zaehler[$v] -eq 0) { $zaehler[$v] = 1 }
        else {
            $n = $zaehler[$v]
            do { $vv = "${v}_$n"; $n++ } while ($zaehler.ContainsKey($vv) -and $zaehler[$vv])
            $zaehler[$v] = $n; $zaehler[$vv] = 1
        }
        $kopf[$c - $sc] = $vv
    }
    foreach ($r in $Sheet.Werte.Keys) {
        if ($r -le $sr -or $r -gt $er) { continue }
        $z = $Sheet.Werte[$r]
        if (-not @($z.Keys | Where-Object { $_ -ge $sc -and $_ -le $ec }).Count) { continue }   # Zeile ohne Wert im Bereich
        $h = New-Object Collections.Hashtable ([StringComparer]::Ordinal)
        for ($c = $sc; $c -le $ec; $c++) { $h[$kopf[$c - $sc]] = $(if ($z.ContainsKey($c)) { $z[$c] } else { '' }) }
        $list.Add($h)
    }
    return $list
}

# --- CSV lesen --------------------------------------------------------------------------------------
<#
.SYNOPSIS
    CSV-Text in Zeilen zerlegen (Anfuehrungszeichen, "" als Zeichen, Zeilenumbrueche in Feldern, CR wird ignoriert).
.OUTPUTS
    List[object] von string[] je Zeile.
#>
function Split-HUCsvText {
    param([Parameter(Mandatory, Position = 0)][AllowEmptyString()][string]$Text, [Parameter(Position = 1)][char]$Delimiter = ';')
    $rows = New-Object Collections.Generic.List[object]
    $cur = New-Object Collections.Generic.List[string]
    $sb = New-Object Text.StringBuilder
    $inQ = $false; $i = 0; $n = $Text.Length
    while ($i -lt $n) {
        $ch = $Text[$i]
        if ($inQ) {
            if ($ch -eq '"') { if ($i + 1 -lt $n -and $Text[$i + 1] -eq '"') { [void]$sb.Append('"'); $i++ } else { $inQ = $false } }
            else { [void]$sb.Append($ch) }
        } elseif ($ch -eq '"') { $inQ = $true }
        elseif ($ch -eq $Delimiter) { $cur.Add($sb.ToString()); [void]$sb.Clear() }
        elseif ($ch -eq "`r") { }
        elseif ($ch -eq "`n") { $cur.Add($sb.ToString()); [void]$sb.Clear(); $rows.Add($cur.ToArray()); $cur.Clear() }
        else { [void]$sb.Append($ch) }
        $i++
    }
    if ($sb.Length -or $cur.Count) { $cur.Add($sb.ToString()); $rows.Add($cur.ToArray()) }
    return , $rows
}
<#
.SYNOPSIS
    Bytes einer Textdatei als Text: UTF-8 mit BOM, UTF-16 mit BOM, UTF-8 ohne BOM, sonst Windows-1252.
#>
function ConvertFrom-HUImportTextBytes([byte[]]$Bytes) {
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { return [Text.Encoding]::UTF8.GetString($Bytes, 3, $Bytes.Length - 3) }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($Bytes, 2, $Bytes.Length - 2) }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return [Text.Encoding]::BigEndianUnicode.GetString($Bytes, 2, $Bytes.Length - 2) }
    try { return (New-Object Text.UTF8Encoding($false, $true)).GetString($Bytes) }
    catch { return [Text.Encoding]::GetEncoding(1252).GetString($Bytes) }
}
<#
.SYNOPSIS
    CSV lesen. Kodierung und Trennzeichen werden erkannt (siehe Kopf der Datei), -Delimiter gibt das Trennzeichen fest vor.
.OUTPUTS
    object[] mit einem Blatt (Name = -Name, Standard 'CSV'; Rows).
#>
function Read-HUCsv {
    param([Parameter(Mandatory, Position = 0)][AllowEmptyCollection()][byte[]]$Bytes, [string]$Delimiter = '', [string]$Name = 'CSV')
    $text = ConvertFrom-HUImportTextBytes $Bytes
    $trenn = ';'
    if ($Delimiter) { $trenn = $Delimiter[0] }
    else {
        $first = ($text -split "`n", 2)[0]
        $best = 0
        foreach ($c in ';', ',', "`t") { $k = @($first.ToCharArray() | Where-Object { $_ -eq $c }).Count; if ($k -gt $best) { $best = $k; $trenn = $c } }
    }
    return @([pscustomobject]@{ Name = $Name; Rows = (Split-HUCsvText $text ([char]$trenn)) })
}
<#
.SYNOPSIS
    Datei fuer einen Import lesen (.xlsx, .xlsm, .csv, .txt). Rueckgabe: Blaetter (Name, Rows).
.PARAMETER SheetJs
    Nur fuer xlsx/xlsm: wie Read-HUXlsx -SheetJs.
#>
function Read-HUImportFile {
    param([Parameter(Mandatory, Position = 0)][string]$Path, [switch]$SheetJs, [int]$MaxRows = -1)
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($ext -notin '.xlsx', '.xlsm', '.csv', '.txt') { throw "Dateityp $ext wird nicht unterstützt (bitte als .xlsx oder .csv speichern)." }
    $b = [IO.File]::ReadAllBytes($Path)
    if ($ext -in '.xlsx', '.xlsm') { return @(Read-HUXlsx $b -SheetJs:$SheetJs -MaxRows $MaxRows) }
    return @(Read-HUCsv $b)
}

# --- Excel schreiben --------------------------------------------------------------------------------
# Blatt-XML: Kopf (string[]) fett und fixiert, Zeilen (string[][]) als Text, Breiten nach dem laengsten Text
function New-HUImportSheetXml([string[]]$Kopf, $Zeilen) {
    $alle = New-Object Collections.Generic.List[object]
    $alle.Add($Kopf)
    foreach ($z in @($Zeilen)) { if ($null -ne $z) { $alle.Add(@($z)) } }
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>')
    for ($i = 0; $i -lt $Kopf.Length; $i++) {
        $w = $Kopf[$i].Length
        for ($r = 1; $r -lt $alle.Count; $r++) { $z = $alle[$r]; if ($i -lt $z.Length -and "$($z[$i])".Length -gt $w) { $w = "$($z[$i])".Length } }
        [void]$sb.AppendFormat('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>', ($i + 1), [Math]::Min(80, $w + 4))
    }
    [void]$sb.Append('</cols><sheetData>')
    for ($r = 0; $r -lt $alle.Count; $r++) {
        [void]$sb.AppendFormat('<row r="{0}">', ($r + 1))
        $z = $alle[$r]
        for ($i = 0; $i -lt $z.Length; $i++) {
            if ("$($z[$i])" -eq '') { continue }
            $st = $(if ($r -eq 0) { ' s="1"' } else { '' })
            [void]$sb.AppendFormat('<c r="{0}{1}" t="inlineStr"{2}><is><t xml:space="preserve">{3}</t></is></c>', (Get-HUXlsxColumnName $i), ($r + 1), $st, (ConvertTo-HUImportXmlText $z[$i]))
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($alle.Count -gt 1 -and $Kopf.Length -gt 0) { [void]$sb.AppendFormat('<autoFilter ref="A1:{0}{1}"/>', (Get-HUXlsxColumnName ($Kopf.Length - 1)), $alle.Count) }
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}
<#
.SYNOPSIS
    Einfache xlsx im Speicher bauen (ohne Excel).
.PARAMETER Sheets
    Blaetter mit Name, Kopf (string[]) und Zeilen (Liste von string[]). Alle Zellen werden als Text geschrieben.
    Der Blattname wird nicht geprueft (Excel erlaubt hoechstens 31 Zeichen und keines von : \ / ? * [ ]).
.OUTPUTS
    byte[]
#>
function New-HUXlsxBytes {
    param([Parameter(Mandatory, Position = 0)][Alias('Blaetter')]$Sheets)
    $liste = @($Sheets)
    $ms = New-Object IO.MemoryStream
    $zip = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Create, $true)
    $utf8 = New-Object Text.UTF8Encoding($false)
    $add = {
        param($Name, $Text)
        $e = $zip.CreateEntry($Name, [IO.Compression.CompressionLevel]::Optimal)
        $sw = New-Object IO.StreamWriter($e.Open(), $utf8)
        try { $sw.Write($Text) } finally { $sw.Dispose() }
    }
    try {
        $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        $wbS = ''; $wbR = ''
        for ($i = 0; $i -lt $liste.Count; $i++) {
            $n = $i + 1
            $ct += "<Override PartName=`"/xl/worksheets/sheet$n.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
            $wbS += "<sheet name=`"$(ConvertTo-HUImportXmlText $liste[$i].Name)`" sheetId=`"$n`" r:id=`"rId$n`"/>"
            $wbR += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$n.xml`"/>"
        }
        $nS = $liste.Count + 1
        $wbR += "<Relationship Id=`"rId$nS`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"
        & $add '[Content_Types].xml' ($ct + '</Types>')
        & $add '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        & $add 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbS + '</sheets></workbook>')
        & $add 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $wbR + '</Relationships>')
        & $add 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
        for ($i = 0; $i -lt $liste.Count; $i++) {
            & $add "xl/worksheets/sheet$($i + 1).xml" (New-HUImportSheetXml ([string[]]@($liste[$i].Kopf)) $liste[$i].Zeilen)
        }
    } finally { $zip.Dispose() }
    $b = $ms.ToArray(); $ms.Dispose()
    return , $b
}

Export-ModuleMember -Function Read-HUXlsx, ConvertTo-HUXlsxRowObject, Read-HUCsv, Split-HUCsvText, Read-HUImportFile, New-HUXlsxBytes,
    ConvertTo-HUJsNumberText, Get-HUXlsxColumnIndex, Get-HUXlsxColumnName, ConvertFrom-HUXlsxRange
