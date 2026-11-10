#Requires -Version 5.1
<#
.SYNOPSIS
    HU.Qr - QR-Codes fuer Module: Daten verpacken, in ein PDF (HU.Pdf) zeichnen, aus Bildern und gescannten PDFs lesen.
.DESCRIPTION
    Bibliothek: ZXing.Net 0.16.11 (Apache-2.0), Datei Lib\ZXing.Net\zxing.dll (Build net48, laeuft unter Windows PowerShell 5.1
    und PowerShell 7). Wird nur aus dem Programmordner geladen, als Bytes nach Pruefung der SHA256-Pruefsumme (nie aus dem Netz).
    Neue Fassungen kommen nur mit einem getesteten Release; der Waechter (.github\waechter.json) meldet neue Versionen und Luecken.

    Daten verpacken:  New-HUQrPayload -Prefix 'HUR1' -Data $obj  ->  'HUR1:' + Base64(Deflate(JSON, UTF-8))
                      Read-HUQrPayload -Text $t -Prefix 'HUR1'    ->  Objekt oder $null
    Zeichnen:         Add-HUPdfQrCode -Pdf $pdf -Text $t -X -Y -Size (Punkt, Y von oben wie HU.Pdf)
    Lesen:            Read-HUQrCodes -Path <Bild oder PDF>        ->  alle gefundenen Texte
      Gescannte PDFs: Bilder werden direkt aus dem PDF geholt (JPEG/DCT, Flate mit/ohne PNG-Vorhersage, CCITT G3/G4 ueber
      eine TIFF-Huelle). JPEG 2000 und JBIG2 kann Windows nicht lesen: dann kommt eine klare Meldung.
.NOTES
    Zielmaschine: der PC, auf dem die App laeuft. PowerShell 5.1 und 7. Laeuft auch in einem Hintergrund-Runspace.
    Add-HUPdfQrCode braucht HU.Pdf (vorher importiert).
#>
Set-StrictMode -Version 1.0

$script:HUQrLibVersion = '0.16.11'
$script:HUQrLibSha256 = '52F00A43F9574A411009899A6236F3EFB5B597FBAE0F3163DFEC9CBB0C1F4D38'
$script:HUQrLibPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Lib\ZXing.Net\zxing.dll'
$script:HUQrMaxText = 2300   # QR Version 40, Fehlerkorrektur M: hoechstens 2331 Bytes

# ---------------------------------------------------------------- Bibliothek ----------------------------------------------------------------
function Initialize-HUQr {
    if ('ZXing.QrCode.QRCodeWriter' -as [type]) { return }
    if (-not (Test-Path -LiteralPath $script:HUQrLibPath)) { throw "QR-Bibliothek fehlt: $($script:HUQrLibPath). Bitte HU-Amt aktualisieren." }
    $bytes = [IO.File]::ReadAllBytes($script:HUQrLibPath)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $h = ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '') } finally { $sha.Dispose() }
    if ($h -ne $script:HUQrLibSha256) { throw 'QR-Bibliothek ist verändert (Prüfsumme falsch) und wird nicht geladen. Bitte HU-Amt aktualisieren.' }
    [void][Reflection.Assembly]::Load($bytes)
    foreach ($a in 'PresentationCore', 'WindowsBase') { try { Add-Type -AssemblyName $a -ErrorAction Stop } catch { } }
}

# ---------------------------------------------------------------- Daten verpacken ----------------------------------------------------------------
function New-HUQrPayload {
    param([Parameter(Mandatory)][string]$Prefix, [Parameter(Mandatory)]$Data)
    if ($Prefix -cnotmatch '^[A-Z][A-Z0-9]{1,9}$') { throw "Ungültiges Kürzel für den QR-Inhalt: '$Prefix'" }
    $json = ConvertTo-Json -InputObject $Data -Depth 10 -Compress
    $raw = [Text.Encoding]::UTF8.GetBytes($json)
    $ms = New-Object IO.MemoryStream
    $z = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionLevel]::Optimal)
    try { $z.Write($raw, 0, $raw.Length) } finally { $z.Dispose() }
    $txt = "$($Prefix):" + [Convert]::ToBase64String($ms.ToArray())
    if ($txt.Length -gt $script:HUQrMaxText) { throw "Zu viele Daten für einen QR-Code ($($txt.Length) Zeichen, höchstens $($script:HUQrMaxText))." }
    return $txt
}

function Read-HUQrPayload {
    param([string]$Text, [Parameter(Mandatory)][string]$Prefix)
    if (-not "$Text".StartsWith("$($Prefix):", [StringComparison]::Ordinal)) { return $null }
    try {
        $bytes = [Convert]::FromBase64String("$Text".Substring($Prefix.Length + 1))
        $z = New-Object IO.Compression.DeflateStream((New-Object IO.MemoryStream(, $bytes)), [IO.Compression.CompressionMode]::Decompress)
        $r = New-Object IO.StreamReader($z, [Text.Encoding]::UTF8)
        try { $json = $r.ReadToEnd() } finally { $r.Dispose() }
        return ($json | ConvertFrom-Json -ErrorAction Stop)
    } catch { return $null }
}

# ---------------------------------------------------------------- Schreiben ----------------------------------------------------------------
# QR als Matrix: Size (Module inkl. Rand) und Rows (je Zeile ein Text aus 0/1, 1 = schwarz)
function New-HUQrMatrix {
    param([Parameter(Mandatory)][string]$Text, [ValidateSet('L', 'M', 'Q', 'H')][string]$Ecc = 'M', [int]$Margin = 4)
    Initialize-HUQr
    if ($Text.Length -gt $script:HUQrMaxText) { throw "Zu viele Daten für einen QR-Code ($($Text.Length) Zeichen)." }
    $hints = New-Object 'System.Collections.Generic.Dictionary[ZXing.EncodeHintType,object]'
    $hints[[ZXing.EncodeHintType]::ERROR_CORRECTION] = [ZXing.QrCode.Internal.ErrorCorrectionLevel]::$Ecc
    $hints[[ZXing.EncodeHintType]::MARGIN] = $Margin
    $hints[[ZXing.EncodeHintType]::CHARACTER_SET] = 'UTF-8'
    $m = (New-Object ZXing.QrCode.QRCodeWriter).encode($Text, [ZXing.BarcodeFormat]::QR_CODE, 0, 0, $hints)
    $rows = New-Object 'System.Collections.Generic.List[string]'
    $sb = New-Object Text.StringBuilder
    for ($y = 0; $y -lt $m.Height; $y++) {
        [void]$sb.Clear()
        for ($x = 0; $x -lt $m.Width; $x++) { [void]$sb.Append($(if ($m.Item($x, $y)) { '1' } else { '0' })) }
        $rows.Add($sb.ToString())
    }
    return [pscustomobject]@{ Size = [int]$m.Width; Rows = $rows.ToArray() }
}

# QR in ein HU.Pdf-Dokument zeichnen (weisser Grund, schwarze Module, je Zeile zusammenhaengende Strecken als ein Rechteck)
function Add-HUPdfQrCode {
    param([Parameter(Mandatory)]$Pdf, [Parameter(Mandatory)][string]$Text, [double]$X, [double]$Y, [double]$Size = 85, [ValidateSet('L', 'M', 'Q', 'H')][string]$Ecc = 'M')
    $q = New-HUQrMatrix -Text $Text -Ecc $Ecc
    $mod = $Size / $q.Size
    Add-HUPdfRect -Pdf $Pdf -X $X -Y $Y -W $Size -H $Size -Gray 1
    for ($r = 0; $r -lt $q.Size; $r++) {
        $row = $q.Rows[$r]; $c = 0
        while ($c -lt $q.Size) {
            if ($row[$c] -ne '1') { $c++; continue }
            $s = $c; while ($c -lt $q.Size -and $row[$c] -eq '1') { $c++ }
            # minimal ueberlappen, damit zwischen Modulen keine hellen Haarlinien entstehen
            Add-HUPdfRect -Pdf $Pdf -X ($X + $s * $mod) -Y ($Y + $r * $mod) -W (($c - $s) * $mod + 0.05) -H ($mod + 0.05) -Gray 0
        }
    }
    return $q
}

# ---------------------------------------------------------------- Lesen ----------------------------------------------------------------
# BitmapSource -> Texte aller QR-Codes darin (leer = keiner gefunden)
function Read-HUQrFromBitmap {
    param([Parameter(Mandatory)]$Bitmap)
    Initialize-HUQr
    $g = New-Object Windows.Media.Imaging.FormatConvertedBitmap($Bitmap, [Windows.Media.PixelFormats]::Gray8, $null, 0)
    $tries = @($g)
    # grosse Scans zusaetzlich verkleinert versuchen (schneller, oft sicherer bei Raster und Rauschen)
    if ([Math]::Max($g.PixelWidth, $g.PixelHeight) -gt 1800) {
        $f = 1600.0 / [Math]::Max($g.PixelWidth, $g.PixelHeight)
        $tries += , (New-Object Windows.Media.Imaging.TransformedBitmap($g, (New-Object Windows.Media.ScaleTransform($f, $f))))
    }
    foreach ($b in $tries) {
        $gb = $(if ($b.Format -eq [Windows.Media.PixelFormats]::Gray8) { $b } else { New-Object Windows.Media.Imaging.FormatConvertedBitmap($b, [Windows.Media.PixelFormats]::Gray8, $null, 0) })
        $w = [int]$gb.PixelWidth; $h = [int]$gb.PixelHeight
        $buf = New-Object byte[] ($w * $h); $gb.CopyPixels($buf, $w, 0)
        $src = New-Object ZXing.RGBLuminanceSource($buf, $w, $h, [ZXing.RGBLuminanceSource+BitmapFormat]::Gray8)
        $hints = New-Object 'System.Collections.Generic.Dictionary[ZXing.DecodeHintType,object]'
        $hints[[ZXing.DecodeHintType]::TRY_HARDER] = $true
        $hints[[ZXing.DecodeHintType]::CHARACTER_SET] = 'UTF-8'
        foreach ($bin in @((New-Object ZXing.Common.HybridBinarizer($src)), (New-Object ZXing.Common.GlobalHistogramBinarizer($src)))) {
            $res = $null
            try { $res = (New-Object ZXing.Multi.QrCode.QRCodeMultiReader).decodeMultiple((New-Object ZXing.BinaryBitmap($bin)), $hints) } catch { $res = $null }
            $texts = @(@($res) | Where-Object { $_ -and "$($_.Text)" } | ForEach-Object { "$($_.Text)" } | Select-Object -Unique)
            if ($texts.Count) { return $texts }
        }
    }
    return @()
}

function ConvertTo-HUBitmapFromBytes([byte[]]$Bytes) {
    $ms = New-Object IO.MemoryStream(, $Bytes)
    $dec = [Windows.Media.Imaging.BitmapDecoder]::Create($ms, [Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    return $dec.Frames[0]
}

# --- Bilder aus einem PDF holen (gescannte Seiten) ---
function Get-HUPdfDictValue([string]$Dict, [string]$Key) {
    $m = [regex]::Match($Dict, "/$Key(?![A-Za-z])\s*(\[[^\]]*\]|<<.*?>>|/[A-Za-z0-9]+|\d+\s+\d+\s+R|-?\d+(\.\d+)?|true|false)", 'Singleline')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return ''
}

function Resolve-HUPdfInt([string]$Text, [string]$Value) {
    if ($Value -match '^(\d+)\s+(\d+)\s+R$') {
        $m = [regex]::Match($Text, "(?<![0-9])$($Matches[1])\s+$($Matches[2])\s+obj\s*(\d+)\s*endobj")
        if ($m.Success) { return [int]$m.Groups[1].Value }
        return -1
    }
    if ($Value -match '^-?\d+$') { return [int]$Value }
    return -1
}

function ConvertFrom-HUPngPredictor([byte[]]$Data, [int]$Columns, [int]$Colors, [int]$Bpc) {
    $bpp = [Math]::Max(1, [int][Math]::Ceiling($Colors * $Bpc / 8.0))
    $rowLen = [int][Math]::Ceiling($Columns * $Colors * $Bpc / 8.0)
    $rows = [int][Math]::Floor($Data.Length / ($rowLen + 1))
    $out = New-Object byte[] ($rows * $rowLen)
    $prev = New-Object byte[] $rowLen
    for ($r = 0; $r -lt $rows; $r++) {
        $t = $Data[$r * ($rowLen + 1)]; $o = $r * ($rowLen + 1) + 1
        $cur = New-Object byte[] $rowLen
        for ($i = 0; $i -lt $rowLen; $i++) {
            $a = $(if ($i -ge $bpp) { [int]$cur[$i - $bpp] } else { 0 }); $b = [int]$prev[$i]; $c = $(if ($i -ge $bpp) { [int]$prev[$i - $bpp] } else { 0 })
            $x = [int]$Data[$o + $i]
            $v = switch ($t) {
                1 { $x + $a } 2 { $x + $b } 3 { $x + [Math]::Floor(($a + $b) / 2) }
                4 { $p = $a + $b - $c; $pa = [Math]::Abs($p - $a); $pb = [Math]::Abs($p - $b); $pc = [Math]::Abs($p - $c); $x + $(if ($pa -le $pb -and $pa -le $pc) { $a } elseif ($pb -le $pc) { $b } else { $c }) }
                default { $x }
            }
            $cur[$i] = [byte]($v -band 0xFF)
        }
        [Array]::Copy($cur, 0, $out, $r * $rowLen, $rowLen); $prev = $cur
    }
    return , $out
}

# CCITT-Daten in eine minimale TIFF-Datei packen (WPF liest CCITT G3/G4 aus TIFF)
function New-HUCcittTiff([byte[]]$Data, [int]$Width, [int]$Height, [int]$K, [bool]$BlackIs1) {
    $ms = New-Object IO.MemoryStream; $bw = New-Object IO.BinaryWriter($ms)
    $tags = @(
        @(256, 4, $Width), @(257, 4, $Height), @(258, 3, 1),
        @(259, 3, $(if ($K -lt 0) { 4 } else { 3 })),
        @(262, 3, $(if ($BlackIs1) { 0 } else { 1 })),   # PDF: BlackIs1=false heisst 0 = schwarz -> TIFF BlackIsZero (1)
        @(273, 4, 0), @(277, 3, 1), @(278, 4, $Height), @(279, 4, $Data.Length)
    )
    if ($K -gt 0) { $tags += , @(292, 4, 1) }   # T4Options: 2D-Kodierung
    $bw.Write([byte[]](0x49, 0x49, 42, 0)); $bw.Write([int]8)
    $bw.Write([uint16]$tags.Count)
    $dataOff = 8 + 2 + 12 * $tags.Count + 4
    foreach ($t in ($tags | Sort-Object { $_[0] })) {
        $bw.Write([uint16]$t[0]); $bw.Write([uint16]$t[1]); $bw.Write([int]1)
        $val = $(if ($t[0] -eq 273) { $dataOff } else { $t[2] })
        if ($t[1] -eq 3) { $bw.Write([uint16]$val); $bw.Write([uint16]0) } else { $bw.Write([int]$val) }
    }
    $bw.Write([int]0); $bw.Write($Data); $bw.Flush()
    return , $ms.ToArray()
}

<#
.SYNOPSIS
    Alle Bilder eines PDF als BitmapSource (gescannte Seiten). Rueckgabe: @{ Bitmaps = @(...); Skipped = @(Grund ...) }
#>
function Get-HUPdfScanImages {
    param([Parameter(Mandatory)][string]$Path)
    foreach ($a in 'PresentationCore', 'WindowsBase') { try { Add-Type -AssemblyName $a -ErrorAction Stop } catch { } }
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = [Text.Encoding]::GetEncoding(28591).GetString($bytes)   # Latin-1: 1 Zeichen = 1 Byte
    $out = New-Object 'System.Collections.Generic.List[object]'; $skip = @()
    foreach ($im in [regex]::Matches($text, '/Subtype\s*/Image')) {
        $objStart = $text.LastIndexOf(' obj', $im.Index); if ($objStart -lt 0) { continue }
        $st = [regex]::Match($text.Substring($im.Index), '>>\s*stream(\r\n|\n|\r)')
        if (-not $st.Success) { continue }
        $dict = $text.Substring($objStart, $im.Index - $objStart + $st.Index + 2)
        $dataStart = $im.Index + $st.Index + $st.Length
        $len = Resolve-HUPdfInt $text (Get-HUPdfDictValue $dict 'Length')
        if ($len -lt 0 -or $dataStart + $len -gt $bytes.Length) { $e = $text.IndexOf('endstream', $dataStart); if ($e -lt 0) { continue }; $len = $e - $dataStart }
        $data = New-Object byte[] $len; [Array]::Copy($bytes, $dataStart, $data, 0, $len)
        $filter = (Get-HUPdfDictValue $dict 'Filter') -replace '[\[\]]', ' '
        $w = Resolve-HUPdfInt $text (Get-HUPdfDictValue $dict 'Width'); $h = Resolve-HUPdfInt $text (Get-HUPdfDictValue $dict 'Height')
        $bpc = Resolve-HUPdfInt $text (Get-HUPdfDictValue $dict 'BitsPerComponent'); if ($bpc -le 0) { $bpc = 8 }
        $cs = Get-HUPdfDictValue $dict 'ColorSpace'
        $parms = Get-HUPdfDictValue $dict 'DecodeParms'
        try {
            if ($filter -match 'DCTDecode') {
                if ($filter -match 'FlateDecode') { $data = ConvertFrom-HUPdfFlate $data }
                $out.Add((ConvertTo-HUBitmapFromBytes $data))
            } elseif ($filter -match 'CCITTFaxDecode') {
                $k = 0; if ($parms -match '/K\s*(-?\d+)') { $k = [int]$Matches[1] }
                $cols = $w; if ($parms -match '/Columns\s*(\d+)') { $cols = [int]$Matches[1] }
                $rowsP = $h; if ($parms -match '/Rows\s*(\d+)') { $rowsP = [int]$Matches[1] }
                $b1 = ($parms -match '/BlackIs1\s*true')
                $out.Add((ConvertTo-HUBitmapFromBytes (New-HUCcittTiff $data $cols $rowsP $k $b1)))
            } elseif ($filter -match 'FlateDecode' -or -not $filter.Trim()) {
                if ($filter -match 'FlateDecode') { $data = ConvertFrom-HUPdfFlate $data }
                $colors = $(if ($cs -match 'DeviceRGB|CalRGB') { 3 } elseif ($cs -match 'DeviceCMYK') { 4 } elseif ($cs -match 'DeviceGray|CalGray') { 1 } else { 0 })
                if (-not $colors) { $skip += "Farbraum $cs nicht unterstützt"; continue }
                if ($parms -match '/Predictor\s*(\d+)' -and [int]$Matches[1] -ge 10) { $data = ConvertFrom-HUPngPredictor $data $w $colors $bpc }
                $fmt = switch ("$colors/$bpc") { '1/1' { [Windows.Media.PixelFormats]::BlackWhite } '1/8' { [Windows.Media.PixelFormats]::Gray8 } '3/8' { [Windows.Media.PixelFormats]::Rgb24 } '4/8' { [Windows.Media.PixelFormats]::Cmyk32 } default { $null } }
                if (-not $fmt) { $skip += "Bildformat $colors Farben / $bpc Bit nicht unterstützt"; continue }
                $stride = [int][Math]::Ceiling($w * $colors * $bpc / 8.0)
                if ($data.Length -lt $stride * $h) { $skip += 'Bilddaten unvollständig'; continue }
                $out.Add([Windows.Media.Imaging.BitmapSource]::Create($w, $h, 96, 96, $fmt, $null, $data, $stride))
            } elseif ($filter -match 'JPXDecode') { $skip += 'JPEG 2000 (von Windows nicht lesbar)' }
            elseif ($filter -match 'JBIG2Decode') { $skip += 'JBIG2 (von Windows nicht lesbar)' }
            else { $skip += "Bildfilter $filter nicht unterstützt" }
        } catch { $skip += "Bild nicht lesbar: $($_.Exception.Message)" }
    }
    return [pscustomobject]@{ Bitmaps = $out.ToArray(); Skipped = $skip }
}

function ConvertFrom-HUPdfFlate([byte[]]$Data) {
    # zlib-Kopf (2 Bytes) ueberspringen, Rest ist Deflate
    $ms = New-Object IO.MemoryStream -ArgumentList $Data, 2, ($Data.Length - 2)
    $z = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
    $o = New-Object IO.MemoryStream
    try { $z.CopyTo($o) } finally { $z.Dispose() }
    return , $o.ToArray()
}

<#
.SYNOPSIS
    QR-Codes aus einer Datei lesen: Bild (JPG, PNG, TIF, BMP, GIF) oder PDF (gescannte Seiten).
    Rueckgabe: @{ Texts = @(...); Hint = '' | Grund, warum nichts gelesen werden konnte }
#>
function Read-HUQrCodes {
    param([Parameter(Mandatory)][string]$Path)
    Initialize-HUQr
    $texts = New-Object 'System.Collections.Generic.List[string]'; $hint = ''
    if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -eq '.pdf') {
        $img = Get-HUPdfScanImages -Path $Path
        foreach ($b in @($img.Bitmaps)) { foreach ($t in @(Read-HUQrFromBitmap -Bitmap $b)) { if (-not $texts.Contains($t)) { $texts.Add($t) } } }
        if (-not $texts.Count) {
            $hint = $(if (-not @($img.Bitmaps).Count -and @($img.Skipped).Count) { "Der Scan ist in einem Format gespeichert, das Windows nicht lesen kann: $(@($img.Skipped | Select-Object -Unique) -join ', '). Am Scanner bitte JPEG oder Farbe/Graustufen wählen." }
                elseif (-not @($img.Bitmaps).Count) { 'Im PDF ist kein gescanntes Bild.' }
                else { 'Im Scan wurde kein QR-Code gefunden. Bitte mit mindestens 200 dpi scannen.' })
        }
    } else {
        $b = ConvertTo-HUBitmapFromBytes ([IO.File]::ReadAllBytes($Path))
        foreach ($t in @(Read-HUQrFromBitmap -Bitmap $b)) { $texts.Add($t) }
        if (-not $texts.Count) { $hint = 'Im Bild wurde kein QR-Code gefunden.' }
    }
    return [pscustomobject]@{ Texts = $texts.ToArray(); Hint = $hint }
}

Export-ModuleMember -Function Initialize-HUQr, New-HUQrPayload, Read-HUQrPayload, New-HUQrMatrix, Add-HUPdfQrCode, Read-HUQrFromBitmap, Get-HUPdfScanImages, Read-HUQrCodes
