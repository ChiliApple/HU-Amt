#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Pdf - kleiner PDF-Schreiber (PDF 1.4) fuer Module, ohne fremde Programme oder Bibliotheken.
.DESCRIPTION
    - Schrift Helvetica und Helvetica-Bold (Standard-14-Schriften, in jedem PDF-Programm vorhanden), Zeichensatz WinAnsi.
      Zeichen ausserhalb von WinAnsi werden ersetzt (z. B. Pfeil -> "->") oder als ? ausgegeben.
    - Breiten werden mit den Massen von Arial gemessen (gleiche Laufweite wie Helvetica), ueber WPF (PresentationCore).
      Fehlt Arial, gilt eine mittlere Breite von 0,556 em je Zeichen.
    - Seiten, Text (links, rechts, zentriert), Linie, gefuelltes Rechteck, JPEG-Bilder (JPG/PNG/BMP... werden ueber WPF
      verkleinert und als JPEG RGB eingebettet), Tabellen ueber mehrere Seiten mit wiederholter Kopfzeile, Fusszeile mit Seitenzahl.
    - Koordinaten in Punkt (1/72 Zoll), Y wird von OBEN gemessen (bei Text: Grundlinie).
    - Ausgabe als byte[] (Get-HUPdfBytes); Speichern macht das Modul selbst.
    Laeuft auch in einem nackten Hintergrund-Runspace: PresentationCore und WindowsBase werden bei Bedarf selbst geladen.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7 (Windows).
    Quelle Aufbau: PDF-Referenz 1.4 (xref-Eintraege genau 20 Byte, Standard-14-Schriften, WinAnsiEncoding Anhang D).
#>

Set-StrictMode -Version 1.0

# --- Zeichen und Breiten ----------------------------------------------------------------------
# Unicode -> WinAnsi (Bereich 0x80-0x9F, siehe PDF-Referenz Anhang D)
$script:HUPdfMap = @{ 0x20AC = 0x80; 0x201A = 0x82; 0x0192 = 0x83; 0x201E = 0x84; 0x2026 = 0x85; 0x2020 = 0x86; 0x2021 = 0x87; 0x02C6 = 0x88; 0x2030 = 0x89; 0x0160 = 0x8A; 0x2039 = 0x8B; 0x0152 = 0x8C; 0x017D = 0x8E; 0x2018 = 0x91; 0x2019 = 0x92; 0x201C = 0x93; 0x201D = 0x94; 0x2022 = 0x95; 0x2013 = 0x96; 0x2014 = 0x97; 0x02DC = 0x98; 0x2122 = 0x99; 0x0161 = 0x9A; 0x203A = 0x9B; 0x0153 = 0x9C; 0x017E = 0x9E; 0x0178 = 0x9F }
# Ersatz fuer haeufige Zeichen ausserhalb von WinAnsi
$script:HUPdfSubst = @{ 0x2192 = '->'; 0x2212 = '-'; 0x2714 = 'x'; 0x2009 = ' '; 0x00A0 = ' '; 0x202F = ' ' }
# WinAnsi -> Unicode (zum Messen)
$script:HUPdfRevMap = @{}
foreach ($kv in $script:HUPdfMap.GetEnumerator()) { $script:HUPdfRevMap[[int]$kv.Value] = [int]$kv.Key }
$script:HUPdfGlyph = @{}
$script:HUPdfWidthCache = @{}
$script:HUPdfInv = [Globalization.CultureInfo]::InvariantCulture

<#
.SYNOPSIS
    Text in WinAnsi-Bytes. Steuerzeichen und Emoji fallen weg, andere Zeichen ausserhalb: Ersatz oder ?. Rueckgabe: byte[].
#>
function ConvertTo-HUPdfBytes {
    param([string]$Text)
    $out = New-Object System.Collections.Generic.List[byte]
    foreach ($ch in "$Text".ToCharArray()) {
        $c = [int]$ch
        if ($script:HUPdfSubst.ContainsKey($c)) { foreach ($x in $script:HUPdfSubst[$c].ToCharArray()) { $out.Add([byte][int]$x) }; continue }
        if ($c -lt 0x20) { continue }
        if ($c -lt 0x80 -or ($c -ge 0xA0 -and $c -le 0xFF)) { $out.Add([byte]$c); continue }
        if ($script:HUPdfMap.ContainsKey($c)) { $out.Add([byte]$script:HUPdfMap[$c]); continue }
        if ([char]::IsSurrogate($ch)) { continue }   # Emoji u. a. fallen weg (wie in den Vorlagen)
        $out.Add([byte][int][char]'?')
    }
    return , $out.ToArray()
}

# Text als PDF-Literal (WinAnsi, Klammern und Backslash maskiert), ohne die umgebenden Klammern
function ConvertTo-HUPdfLiteral([string]$Text) {
    $b = ConvertTo-HUPdfBytes $Text
    $esc = New-Object System.Collections.Generic.List[byte]
    foreach ($c in $b) { if ($c -eq 0x28 -or $c -eq 0x29 -or $c -eq 0x5C) { $esc.Add(0x5C) }; $esc.Add($c) }
    return , $esc.ToArray()
}

# GlyphTypeface von Arial (normal/fett), $null wenn nicht vorhanden
function Get-HUPdfGlyphFace([bool]$Bold) {
    $k = "$Bold"
    if ($script:HUPdfGlyph.ContainsKey($k)) { return $script:HUPdfGlyph[$k] }
    $g = $null
    try {
        Add-Type -AssemblyName PresentationCore, WindowsBase
        $weight = $(if ($Bold) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal })
        $tf = New-Object System.Windows.Media.Typeface ((New-Object System.Windows.Media.FontFamily 'Arial'), [System.Windows.FontStyles]::Normal, $weight, [System.Windows.FontStretches]::Normal)
        if (-not $tf.TryGetGlyphTypeface([ref]$g)) { $g = $null }
    } catch { $g = $null }
    $script:HUPdfGlyph[$k] = $g
    return $g
}

<#
.SYNOPSIS
    Breite eines Textes in Punkt (Zeichen nach WinAnsi, gemessen mit Arial). Rueckgabe: double.
#>
function Get-HUPdfTextWidth {
    param([string]$Text, [double]$Size = 9, [switch]$Bold)
    $g = Get-HUPdfGlyphFace ([bool]$Bold)
    $w = 0.0
    foreach ($b in (ConvertTo-HUPdfBytes $Text)) {
        $key = "$([bool]$Bold)|$b"
        if (-not $script:HUPdfWidthCache.ContainsKey($key)) {
            $code = [int]$b
            if ($script:HUPdfRevMap.ContainsKey($code)) { $code = $script:HUPdfRevMap[$code] }
            $cw = 0.556
            if ($g) { $gi = [uint16]0; if ($g.CharacterToGlyphMap.TryGetValue($code, [ref]$gi)) { $cw = $g.AdvanceWidths[$gi] } }
            $script:HUPdfWidthCache[$key] = [double]$cw
        }
        $w += $script:HUPdfWidthCache[$key]
    }
    return ($w * $Size)
}

<#
.SYNOPSIS
    Text auf eine Breite kuerzen, mit "…" am Ende. Rueckgabe: string (unveraendert, wenn er passt).
#>
function Get-HUPdfFitText {
    param([string]$Text, [double]$Width, [double]$Size = 9, [switch]$Bold)
    $t = "$Text"
    if ((Get-HUPdfTextWidth $t $Size -Bold:$Bold) -le $Width) { return $t }
    $ell = [string][char]0x2026
    while ($t.Length -gt 0 -and (Get-HUPdfTextWidth ($t + $ell) $Size -Bold:$Bold) -gt $Width) { $t = $t.Substring(0, $t.Length - 1) }
    return ($t.TrimEnd() + $ell)
}

<#
.SYNOPSIS
    Text in Zeilen umbrechen (an Leerzeichen und Zeilenumbruechen). Ein einzelnes zu langes Wort bleibt ganz.
    Rueckgabe: string[] (immer ein Array, auch bei einer Zeile).
#>
function Split-HUPdfTextLines {
    param([string]$Text, [double]$Width, [double]$Size = 9, [switch]$Bold)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($para in ("$Text" -split "`r?`n")) {
        $cur = ''
        foreach ($w in ($para -split ' ')) {
            $try = $(if ($cur) { "$cur $w" } else { $w })
            if ((Get-HUPdfTextWidth $try $Size -Bold:$Bold) -le $Width -or -not $cur) { $cur = $try } else { $lines.Add($cur); $cur = $w }
        }
        $lines.Add($cur)
    }
    return , $lines.ToArray()
}

# --- PDF-Schreiber ----------------------------------------------------------------------------
<#
.SYNOPSIS
    Neues PDF-Dokument (noch ohne Seite). Standard A4 hoch; -Landscape tauscht Breite und Hoehe (A4 quer).
    Rueckgabe: Dokument-Objekt (W, H, Pages, Images, Cur) fuer die uebrigen Funktionen.
#>
function New-HUPdf {
    param([double]$Width = 595.28, [double]$Height = 841.89, [switch]$Landscape)
    $w = $Width; $h = $Height
    if ($Landscape -and $w -lt $h) { $w = $Height; $h = $Width }
    return [pscustomobject]@{ W = $w; H = $h; Pages = (New-Object System.Collections.Generic.List[object]); Images = (New-Object System.Collections.Generic.List[object]); Cur = $null }
}

<#
.SYNOPSIS
    Neue Seite anhaengen und zur aktuellen machen. Ohne Groesse gilt die des Dokuments. Rueckgabe: Seiten-Objekt (W, H, Ops, Images).
#>
function Add-HUPdfPage {
    param($Pdf, [double]$Width = 0, [double]$Height = 0)
    $p = [pscustomobject]@{ W = $(if ($Width) { $Width } else { $Pdf.W }); H = $(if ($Height) { $Height } else { $Pdf.H }); Ops = (New-Object System.Collections.Generic.List[byte[]]); Images = (New-Object System.Collections.Generic.List[int]) }
    $Pdf.Pages.Add($p); $Pdf.Cur = $p
    return $p
}

function Add-HUPdfOp($Pdf, [string]$Op) {
    if ($null -eq $Pdf.Cur) { [void](Add-HUPdfPage $Pdf) }
    $Pdf.Cur.Ops.Add([Text.Encoding]::ASCII.GetBytes($Op + "`n"))
}
function Format-HUPdfNum([double]$v) { return $v.ToString('0.##', $script:HUPdfInv) }

<#
.SYNOPSIS
    Text schreiben. Y von oben gemessen (Grundlinie). -Align right: X ist der rechte Rand, center: X ist die Mitte.
    -Gray 0 = schwarz, 1 = weiss. Leerer Text schreibt nichts.
#>
function Add-HUPdfText {
    param($Pdf, [double]$X, [double]$Y, [string]$Text, [double]$Size = 9, [switch]$Bold, [double]$Gray = 0, [ValidateSet('left', 'right', 'center')][string]$Align = 'left')
    if (-not "$Text") { return }
    if ($null -eq $Pdf.Cur) { [void](Add-HUPdfPage $Pdf) }
    $px = $X
    if ($Align -eq 'right') { $px = $X - (Get-HUPdfTextWidth $Text $Size -Bold:$Bold) }
    elseif ($Align -eq 'center') { $px = $X - (Get-HUPdfTextWidth $Text $Size -Bold:$Bold) / 2 }
    $esc = ConvertTo-HUPdfLiteral $Text
    $font = $(if ($Bold) { 'F2' } else { 'F1' })
    $head = [Text.Encoding]::ASCII.GetBytes("BT $(Format-HUPdfNum $Gray) g /$font $(Format-HUPdfNum $Size) Tf $(Format-HUPdfNum $px) $(Format-HUPdfNum ($Pdf.Cur.H - $Y)) Td (")
    $tail = [Text.Encoding]::ASCII.GetBytes(") Tj ET`n")
    $all = New-Object byte[] ($head.Length + $esc.Length + $tail.Length)
    [Array]::Copy($head, 0, $all, 0, $head.Length)
    [Array]::Copy($esc, 0, $all, $head.Length, $esc.Length)
    [Array]::Copy($tail, 0, $all, $head.Length + $esc.Length, $tail.Length)
    $Pdf.Cur.Ops.Add($all)
}

<#
.SYNOPSIS
    Linie von (X1,Y1) nach (X2,Y2), Y von oben. -Width Strichstaerke in Punkt, -Gray 0 = schwarz.
#>
function Add-HUPdfLine {
    param($Pdf, [double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [double]$Width = 0.5, [double]$Gray = 0)
    if ($null -eq $Pdf.Cur) { [void](Add-HUPdfPage $Pdf) }
    Add-HUPdfOp $Pdf "$(Format-HUPdfNum $Gray) G $(Format-HUPdfNum $Width) w $(Format-HUPdfNum $X1) $(Format-HUPdfNum ($Pdf.Cur.H - $Y1)) m $(Format-HUPdfNum $X2) $(Format-HUPdfNum ($Pdf.Cur.H - $Y2)) l S"
}

<#
.SYNOPSIS
    Gefuelltes Rechteck, (X,Y) = linke obere Ecke (Y von oben). -Gray Fuellung (Standard hellgrau 0.9).
#>
function Add-HUPdfRect {
    param($Pdf, [double]$X, [double]$Y, [double]$W, [double]$H, [double]$Gray = 0.9)
    if ($null -eq $Pdf.Cur) { [void](Add-HUPdfPage $Pdf) }
    Add-HUPdfOp $Pdf "$(Format-HUPdfNum $Gray) g $(Format-HUPdfNum $X) $(Format-HUPdfNum ($Pdf.Cur.H - $Y - $H)) $(Format-HUPdfNum $W) $(Format-HUPdfNum $H) re f"
}

<#
.SYNOPSIS
    Bild (alles, was WPF lesen kann: JPG, PNG, BMP, GIF, TIFF) als JPEG RGB vorbereiten, laengste Seite hoechstens -MaxPx.
    Quelle -Path (Datei) oder -Bytes (Inhalt). Rueckgabe: Objekt mit Bytes, PxW, PxH (fuer Add-HUPdfImage).
.NOTES
    Transparenz faellt weg (Umwandlung nach Bgr24). Drehung aus EXIF wird nicht beachtet (wie in der Vorlage).
#>
function ConvertTo-HUPdfJpeg {
    param([string]$Path = '', [byte[]]$Bytes = $null, [int]$MaxPx = 2000, [ValidateRange(1, 100)][int]$Quality = 85)
    Add-Type -AssemblyName PresentationCore, WindowsBase
    if ($Path) { $fs = [IO.File]::OpenRead($Path) }
    elseif ($Bytes) { $fs = New-Object IO.MemoryStream (, $Bytes) }
    else { throw 'ConvertTo-HUPdfJpeg: -Path oder -Bytes fehlt.' }
    try {
        $dec = [System.Windows.Media.Imaging.BitmapDecoder]::Create($fs, [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $src = $dec.Frames[0]
    } finally { $fs.Dispose() }
    $scale = 1.0
    if ($MaxPx -gt 0) { $scale = [Math]::Min(1.0, $MaxPx / [double][Math]::Max($src.PixelWidth, $src.PixelHeight)) }
    $bmp = $src
    if ($scale -lt 1.0) { $bmp = New-Object System.Windows.Media.Imaging.TransformedBitmap ($src, (New-Object System.Windows.Media.ScaleTransform ($scale, $scale))) }
    $rgb = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap ($bmp, [System.Windows.Media.PixelFormats]::Bgr24, $null, 0)
    $enc = New-Object System.Windows.Media.Imaging.JpegBitmapEncoder
    $enc.QualityLevel = $Quality
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rgb))
    $ms = New-Object IO.MemoryStream
    try { $enc.Save($ms); $out = $ms.ToArray() } finally { $ms.Dispose() }
    return [pscustomobject]@{ Bytes = $out; PxW = $rgb.PixelWidth; PxH = $rgb.PixelHeight }
}

<#
.SYNOPSIS
    Bild (Objekt aus ConvertTo-HUPdfJpeg) auf die aktuelle Seite setzen, (X,Y) = linke obere Ecke, W/H in Punkt.
    Dasselbe Objekt mehrfach eingesetzt wird nur einmal eingebettet.
#>
function Add-HUPdfImage {
    param($Pdf, $Jpeg, [double]$X, [double]$Y, [double]$W, [double]$H)
    if ($null -eq $Pdf.Cur) { [void](Add-HUPdfPage $Pdf) }
    $i = 0
    for ($k = 0; $k -lt $Pdf.Images.Count; $k++) { if ([object]::ReferenceEquals($Pdf.Images[$k], $Jpeg)) { $i = $k + 1; break } }
    if (-not $i) { $Pdf.Images.Add($Jpeg); $i = $Pdf.Images.Count }
    if (-not $Pdf.Cur.Images.Contains($i)) { $Pdf.Cur.Images.Add($i) }
    Add-HUPdfOp $Pdf "q $(Format-HUPdfNum $W) 0 0 $(Format-HUPdfNum $H) $(Format-HUPdfNum $X) $(Format-HUPdfNum ($Pdf.Cur.H - $Y - $H)) cm /Im$i Do Q"
}

<#
.SYNOPSIS
    PDF als Bytes. Objekte: 1 Katalog, 2 Seitenbaum, 3/4 Schriften, dann Bilder, dann je Seite Seite + Inhalt, zuletzt Info.
    -Title und -Producer stehen in den Dokumenteigenschaften. Ein Dokument ohne Seite bekommt eine leere Seite.
    Rueckgabe: byte[].
#>
function Get-HUPdfBytes {
    param($Pdf, [string]$Title = '', [string]$Producer = 'HU-Amt', [datetime]$CreationDate = (Get-Date))
    if (-not $Pdf.Pages.Count) { [void](Add-HUPdfPage $Pdf) }
    $ms = New-Object IO.MemoryStream
    $offs = New-Object System.Collections.Generic.List[long]
    $w = { param([string]$s) $b = [Text.Encoding]::ASCII.GetBytes($s); $ms.Write($b, 0, $b.Length) }
    $wb = { param([byte[]]$b) $ms.Write($b, 0, $b.Length) }
    & $w "%PDF-1.4`n%"; & $wb ([byte[]](0xE2, 0xE3, 0xCF, 0xD3)); & $w "`n"
    $nImg = $Pdf.Images.Count
    $firstPage = 5 + $nImg
    $infoObj = $firstPage + 2 * $Pdf.Pages.Count
    $kids = (@(for ($i = 0; $i -lt $Pdf.Pages.Count; $i++) { "$($firstPage + 2 * $i) 0 R" }) -join ' ')
    $obj = {
        param([int]$n, [scriptblock]$body)
        while ($offs.Count -lt $n) { $offs.Add(0) }
        $offs[$n - 1] = $ms.Position
        & $w "$n 0 obj`n"; & $body; & $w "`nendobj`n"
    }
    & $obj 1 { & $w '<< /Type /Catalog /Pages 2 0 R >>' }
    & $obj 2 { & $w "<< /Type /Pages /Kids [$kids] /Count $($Pdf.Pages.Count) >>" }
    & $obj 3 { & $w '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>' }
    & $obj 4 { & $w '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>' }
    for ($i = 0; $i -lt $nImg; $i++) {
        $im = $Pdf.Images[$i]
        & $obj (5 + $i) { & $w "<< /Type /XObject /Subtype /Image /Width $($im.PxW) /Height $($im.PxH) /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length $($im.Bytes.Length) >>`nstream`n"; & $wb $im.Bytes; & $w "`nendstream" }
    }
    for ($i = 0; $i -lt $Pdf.Pages.Count; $i++) {
        $p = $Pdf.Pages[$i]
        $pn = $firstPage + 2 * $i
        $xo = $(if ($p.Images.Count) { ' /XObject << ' + ((@($p.Images | ForEach-Object { "/Im$_ $(4 + $_) 0 R" })) -join ' ') + ' >>' } else { '' })
        & $obj $pn { & $w "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 $(Format-HUPdfNum $p.W) $(Format-HUPdfNum $p.H)] /Resources << /Font << /F1 3 0 R /F2 4 0 R >>$xo >> /Contents $($pn + 1) 0 R >>" }
        $cs = New-Object IO.MemoryStream
        foreach ($o in $p.Ops) { $cs.Write($o, 0, $o.Length) }
        $cb = $cs.ToArray()
        & $obj ($pn + 1) { & $w "<< /Length $($cb.Length) >>`nstream`n"; & $wb $cb; & $w "`nendstream" }
    }
    $te = ConvertTo-HUPdfLiteral $Title
    $pe = ConvertTo-HUPdfLiteral $Producer
    & $obj $infoObj { & $w '<< /Title ('; & $wb $te; & $w ') /Producer ('; & $wb $pe; & $w ") /CreationDate (D:$($CreationDate.ToString('yyyyMMddHHmmss'))) >>" }
    $xref = $ms.Position
    & $w "xref`n0 $($offs.Count + 1)`n0000000000 65535 f `n"
    foreach ($o in $offs) { & $w (('{0:D10} 00000 n ' -f $o) + "`n") }
    & $w "trailer`n<< /Size $($offs.Count + 1) /Root 1 0 R /Info $infoObj 0 R >>`nstartxref`n$xref`n%%EOF`n"
    return , $ms.ToArray()
}

# --- Tabellen und Fusszeile -------------------------------------------------------------------
# Eigenschaft aus Hashtable oder Objekt, sonst Standard
function Get-HUPdfProp($Obj, [string]$Name, $Default = $null) {
    if ($null -eq $Obj) { return $Default }
    if ($Obj -is [System.Collections.IDictionary]) { if ($Obj.Contains($Name)) { return $Obj[$Name] }; return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $Default
}
# Zelltext einer Zeile: Hashtable/Objekt ueber Key, Array ueber die Spaltennummer
function Get-HUPdfCell($Row, $Col, [int]$Index) {
    if ($Row -is [System.Collections.IList] -and -not ($Row -is [string])) { if ($Index -lt $Row.Count) { return "$($Row[$Index])" }; return '' }
    $key = "$(Get-HUPdfProp $Col 'Key' '')"
    if (-not $key) { $key = "$(Get-HUPdfProp $Col 'Name' '')" }
    return "$(Get-HUPdfProp $Row $key '')"
}
# Neue Seite mit Ueberschriften und Tabellenkopf. Rueckgabe: Y der ersten Zeile.
function Add-HUPdfTableHead($Ctx, [bool]$NewPage, [double]$StartY) {
    $pdf = $Ctx.Pdf
    $yy = $StartY
    if ($NewPage) {
        [void](Add-HUPdfPage $pdf)
        $yy = $Ctx.Top
        if ($Ctx.Heading) { Add-HUPdfText $pdf $Ctx.L $yy $Ctx.Heading 10 -Bold; $yy += 20 }
        if ($Ctx.Title) { Add-HUPdfText $pdf $Ctx.L $yy $Ctx.Title 14 -Bold; $yy += 16 }
        if ($Ctx.Subtitle) { Add-HUPdfText $pdf $Ctx.L $yy $Ctx.Subtitle 9 -Gray 0.35; $yy += 20 }
        elseif ($Ctx.Title) { $yy += 8 }
    }
    Add-HUPdfRect $pdf $Ctx.L ($yy - 11) ($Ctx.R - $Ctx.L) 16 0.88
    Add-HUPdfTableRow $Ctx $yy $null $true
    return ($yy + 16)
}
# Eine Zeile schreiben ($Row = $null: Kopfzeile mit den Spaltennamen)
function Add-HUPdfTableRow($Ctx, [double]$Y, $Row, [bool]$Bold) {
    $pdf = $Ctx.Pdf; $x = $Ctx.L
    for ($i = 0; $i -lt $Ctx.Cols.Count; $i++) {
        $c = $Ctx.Cols[$i]; $w = $Ctx.W[$i]
        $raw = $(if ($null -eq $Row) { "$(Get-HUPdfProp $c 'Name' '')" } else { Get-HUPdfCell $Row $c $i })
        $t = Get-HUPdfFitText $raw ($w - 8) $Ctx.Size -Bold:$Bold
        switch ("$(Get-HUPdfProp $c 'Align' 'left')") {
            'right' { Add-HUPdfText $pdf ($x + $w - 5) $Y $t $Ctx.Size -Bold:$Bold -Align right }
            'center' { Add-HUPdfText $pdf ($x + $w / 2) $Y $t $Ctx.Size -Bold:$Bold -Align center }
            default { Add-HUPdfText $pdf ($x + 4) $Y $t $Ctx.Size -Bold:$Bold }
        }
        $x += $w
    }
}

<#
.SYNOPSIS
    Einfache Tabelle ueber mehrere Seiten, Kopfzeile (und Ueberschriften) auf jeder Seite wiederholt.
.DESCRIPTION
    -Columns: je Spalte Hashtable/Objekt mit Name (Kopftext), Key (Feld der Zeile, Standard = Name),
              Align (left|right|center, Standard left) und wahlweise Width (Punkt, 0 = nach Inhalt).
    -Rows:    Hashtables/Objekte (ueber Key) oder Arrays (ueber die Spaltennummer). Werte werden als Text ausgegeben,
              formatiert wird vorher im Modul. Zu lange Zellen werden mit "…" gekuerzt.
    -SumRow:  wahlweise Summenzeile am Ende, fett, mit Linie darueber.
    Breite nach Inhalt: laengster Text + 10 pt, hoechstens -MaxColumnWidth, danach auf die Seitenbreite verkleinert (nie vergroessert).
    Ohne -Y beginnt die Tabelle auf einer neuen Seite mit -Heading (10 fett), -Title (14 fett), -Subtitle (9 grau).
    Mit -Y laeuft sie auf der aktuellen Seite ab dieser Hoehe weiter (nur Tabellenkopf), Folgeseiten bekommen die Ueberschriften.
    Querformat: New-HUPdf -Landscape (die Vorlage nahm quer ab mehr als 6 Spalten).
    Rueckgabe: Y unter der letzten Zeile (double), die aktuelle Seite ist die letzte der Tabelle.
#>
function Add-HUPdfTable {
    param(
        [Parameter(Mandatory)]$Pdf,
        [Parameter(Mandatory)][object[]]$Columns,
        [object[]]$Rows = @(),
        $SumRow = $null,
        [string]$Title = '', [string]$Subtitle = '', [string]$Heading = '',
        [double]$Y = 0,
        [double]$Size = 8.5, [double]$Margin = 40, [double]$Top = 44, [double]$BottomMargin = 50,
        [double]$RowHeight = 13, [double]$MaxColumnWidth = 260,
        [string]$EmptyText = 'Keine Einträge.'
    )
    $rowsAll = @($Rows | Where-Object { $null -ne $_ })
    $pageW = $(if ($Y -gt 0 -and $Pdf.Cur) { $Pdf.Cur.W } else { $Pdf.W })
    $ctx = @{ Pdf = $Pdf; Cols = @($Columns); L = $Margin; R = $pageW - $Margin; Size = $Size; Top = $Top; Heading = $Heading; Title = $Title; Subtitle = $Subtitle; W = @() }
    # Spaltenbreiten
    $wid = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $ctx.Cols.Count; $i++) {
        $c = $ctx.Cols[$i]
        $fix = [double](Get-HUPdfProp $c 'Width' 0)
        if ($fix -gt 0) { $wid.Add($fix); continue }
        $m = (Get-HUPdfTextWidth "$(Get-HUPdfProp $c 'Name' '')" $Size -Bold) + 10
        foreach ($r in $rowsAll) { $tw = (Get-HUPdfTextWidth (Get-HUPdfCell $r $c $i) $Size) + 10; if ($tw -gt $m) { $m = $tw } }
        if ($null -ne $SumRow) { $tw = (Get-HUPdfTextWidth (Get-HUPdfCell $SumRow $c $i) $Size -Bold) + 10; if ($tw -gt $m) { $m = $tw } }
        $wid.Add([Math]::Min($MaxColumnWidth, $m))
    }
    $sum = 0.0; foreach ($w in $wid) { $sum += $w }
    $f = ($ctx.R - $ctx.L) / [Math]::Max(1, $sum)
    $ctx.W = @(foreach ($w in $wid) { if ($f -lt 1) { $w * $f } else { $w } })
    # Kopf
    $newPage = -not ($Y -gt 0 -and $Pdf.Cur)
    $yy = Add-HUPdfTableHead $ctx $newPage $Y
    if (-not $rowsAll.Count -and $EmptyText) { Add-HUPdfText $Pdf $ctx.L $yy $EmptyText 10 -Gray 0.35; $yy += $RowHeight }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rowsAll) { $list.Add($r) }
    if ($null -ne $SumRow) { $list.Add($SumRow) }
    for ($n = 0; $n -lt $list.Count; $n++) {
        $row = $list[$n]
        $isSum = ($null -ne $SumRow -and $n -eq $list.Count - 1)
        if ($yy -gt $Pdf.Cur.H - $BottomMargin) { $yy = Add-HUPdfTableHead $ctx $true 0 }
        if ($isSum) { Add-HUPdfLine $Pdf $ctx.L ($yy - 10) $ctx.R ($yy - 10) 0.8 }
        Add-HUPdfTableRow $ctx $yy $row $isSum
        $yy += $RowHeight
        if (-not $isSum) { Add-HUPdfLine $Pdf $ctx.L ($yy - 9.5) $ctx.R ($yy - 9.5) 0.2 0.8 }
    }
    return [double]$yy
}

<#
.SYNOPSIS
    Fusszeile auf allen Seiten: -Text links (z. B. "Erstellt am ..."), rechts "Seite x von n" (-PageFormat, {0} = Seite, {1} = Anzahl;
    leer = ohne Seitenzahl). Am Ende ist wieder die letzte Seite aktuell. Erst aufrufen, wenn alle Seiten angelegt sind.
#>
function Add-HUPdfPageFooter {
    param([Parameter(Mandatory)]$Pdf, [string]$Text = '', [string]$PageFormat = 'Seite {0} von {1}', [double]$Margin = 40, [double]$Bottom = 24, [double]$Size = 7.5, [double]$Gray = 0.5)
    $n = $Pdf.Pages.Count
    for ($i = 0; $i -lt $n; $i++) {
        $Pdf.Cur = $Pdf.Pages[$i]
        if ($Text) { Add-HUPdfText $Pdf $Margin ($Pdf.Cur.H - $Bottom) $Text $Size -Gray $Gray }
        if ($PageFormat) { Add-HUPdfText $Pdf ($Pdf.Cur.W - $Margin) ($Pdf.Cur.H - $Bottom) ($PageFormat -f ($i + 1), $n) $Size -Gray $Gray -Align right }
    }
}

Export-ModuleMember -Function ConvertTo-HUPdfBytes, Get-HUPdfTextWidth, Get-HUPdfFitText, Split-HUPdfTextLines,
    New-HUPdf, Add-HUPdfPage, Add-HUPdfText, Add-HUPdfLine, Add-HUPdfRect, ConvertTo-HUPdfJpeg, Add-HUPdfImage,
    Get-HUPdfBytes, Add-HUPdfTable, Add-HUPdfPageFooter
