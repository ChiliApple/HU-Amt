#Requires -Version 5.1
# Pester 5: QR-Codes (Core\HU.Qr.psm1): Daten verpacken, schreiben, aus Bildern und PDFs lesen (kuenstliche Scans unter TEMP).

BeforeAll {
    Set-StrictMode -Version 1.0
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Pdf.psm1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:Root 'Core\HU.Qr.psm1') -Force -DisableNameChecking
    Add-Type -AssemblyName PresentationCore, WindowsBase
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtQr-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](New-Item -ItemType Directory -Path $script:Tmp)

    # QR als Graustufen-Bild (wie ein Scan): Rand weiss, Module px gross, Seite mit weissem Umfeld
    function New-TestQrBitmap([string]$Text, [int]$Px = 4, [int]$Pad = 60) {
        $q = New-HUQrMatrix -Text $Text
        $size = $q.Size * $Px + 2 * $Pad
        $pix = New-Object byte[] ($size * $size)
        for ($i = 0; $i -lt $pix.Length; $i++) { $pix[$i] = 255 }
        for ($y = 0; $y -lt $q.Size * $Px; $y++) { for ($x = 0; $x -lt $q.Size * $Px; $x++) { if ($q.Rows[[int][Math]::Floor($y / $Px)][[int][Math]::Floor($x / $Px)] -eq '1') { $pix[($y + $Pad) * $size + $x + $Pad] = 0 } } }
        return [Windows.Media.Imaging.BitmapSource]::Create($size, $size, 300, 300, [Windows.Media.PixelFormats]::Gray8, $null, $pix, $size)
    }
    function ConvertTo-TestJpeg($Bitmap, [int]$Quality = 70) {
        $enc = New-Object Windows.Media.Imaging.JpegBitmapEncoder; $enc.QualityLevel = $Quality; $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($Bitmap))
        $ms = New-Object IO.MemoryStream; $enc.Save($ms); return , $ms.ToArray()
    }
    # minimales PDF mit einem Bild-Objekt (Filter frei waehlbar), wie es Scanner schreiben
    function New-TestImagePdf([string]$Path, [byte[]]$Data, [string]$Dict) {
        $l1 = [Text.Encoding]::GetEncoding(28591)
        $ms = New-Object IO.MemoryStream
        $w = { param($t) $b = $l1.GetBytes($t); $ms.Write($b, 0, $b.Length) }
        & $w "%PDF-1.4`n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj`n2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj`n"
        & $w "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /XObject << /Im1 4 0 R >> >> /Contents 5 0 R >> endobj`n"
        & $w "4 0 obj << /Type /XObject /Subtype /Image $Dict /Length 6 0 R >>`nstream`n"
        $ms.Write($Data, 0, $Data.Length)
        & $w "`nendstream`nendobj`n5 0 obj << /Length 31 >>`nstream`nq 595 0 0 842 0 0 cm /Im1 Do Q`nendstream`nendobj`n6 0 obj $($Data.Length) endobj`ntrailer << /Root 1 0 R >>`n%%EOF`n"
        [IO.File]::WriteAllBytes($Path, $ms.ToArray())
    }
    function ConvertTo-TestZlib([byte[]]$Raw) {
        $ms = New-Object IO.MemoryStream; $ms.WriteByte(0x78); $ms.WriteByte(0x9C)
        $z = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Compress, $true); $z.Write($Raw, 0, $Raw.Length); $z.Dispose()
        return , $ms.ToArray()
    }
    $script:Data = [ordered]@{ v = 1; id = 'a1b2c3'; name = 'Muster Maria'; reisen = @(1..5 | ForEach-Object { [ordered]@{ d = "2026-11-0$_"; nach = "Ort $_ (Ü)"; km = 12.5 * $_ } }) }
}
AfterAll { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }

Describe 'Bibliothek' {
    It 'laedt die gepruefte DLL aus Lib\ZXing.Net' {
        { Initialize-HUQr } | Should -Not -Throw
        ('ZXing.QrCode.QRCodeWriter' -as [type]) | Should -Not -BeNullOrEmpty
        (Get-FileHash -LiteralPath (Join-Path $script:Root 'Lib\ZXing.Net\zxing.dll') -Algorithm SHA256).Hash | Should -Be '52F00A43F9574A411009899A6236F3EFB5B597FBAE0F3163DFEC9CBB0C1F4D38'
        Test-Path -LiteralPath (Join-Path $script:Root 'Lib\ZXing.Net\LICENSE.txt') | Should -BeTrue
    }
}

Describe 'Daten verpacken' {
    It 'hin und zurueck, auch mit Umlauten' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        $t | Should -Match '^HUR1:[A-Za-z0-9+/=]+$'
        $o = Read-HUQrPayload -Text $t -Prefix 'HUR1'
        $o.name | Should -Be 'Muster Maria'
        @($o.reisen)[4].nach | Should -Be 'Ort 5 (Ü)'
        @($o.reisen).Count | Should -Be 5
    }
    It 'falsches Kuerzel oder kaputter Inhalt ergibt $null' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        Read-HUQrPayload -Text $t -Prefix 'XYZ1' | Should -BeNullOrEmpty
        Read-HUQrPayload -Text 'HUR1:kaputt!!' -Prefix 'HUR1' | Should -BeNullOrEmpty
        { New-HUQrPayload -Prefix 'hur' -Data 1 } | Should -Throw
    }
    It 'zu viele Daten werden abgelehnt' {
        $big = @(1..400 | ForEach-Object { [guid]::NewGuid().ToString() })
        { New-HUQrPayload -Prefix 'HUR1' -Data $big } | Should -Throw '*Zu viele Daten*'
    }
}

Describe 'Schreiben und Lesen' {
    It 'Matrix ist quadratisch, mit Rand' {
        $q = New-HUQrMatrix -Text 'Hallo'
        $q.Rows.Count | Should -Be $q.Size
        $q.Rows[0] | Should -Match '^0+$'
    }
    It 'Bild als JPEG (wie ein Scan) wird gelesen' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        $f = Join-Path $script:Tmp 'scan.jpg'
        [IO.File]::WriteAllBytes($f, (ConvertTo-TestJpeg (New-TestQrBitmap $t) 60))
        $r = Read-HUQrCodes -Path $f
        @($r.Texts) | Should -Be @($t)
        $r.Hint | Should -Be ''
    }
    It 'Scan-PDF mit JPEG (DCTDecode) wird gelesen' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        $bmp = New-TestQrBitmap $t
        $f = Join-Path $script:Tmp 'scan-dct.pdf'
        New-TestImagePdf $f (ConvertTo-TestJpeg $bmp 70) "/Width $($bmp.PixelWidth) /Height $($bmp.PixelHeight) /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /DCTDecode"
        @((Read-HUQrCodes -Path $f).Texts) | Should -Be @($t)
    }
    It 'Scan-PDF mit Flate (Graustufen, PNG-Vorhersage) wird gelesen' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        $bmp = New-TestQrBitmap $t
        $w = $bmp.PixelWidth; $h = $bmp.PixelHeight
        $pix = New-Object byte[] ($w * $h); $bmp.CopyPixels($pix, $w, 0)
        $raw = New-Object byte[] (($w + 1) * $h)   # PNG-Filter "Up" je Zeile
        for ($y = 0; $y -lt $h; $y++) { $raw[$y * ($w + 1)] = 2; for ($x = 0; $x -lt $w; $x++) { $up = $(if ($y) { $pix[($y - 1) * $w + $x] } else { 0 }); $raw[$y * ($w + 1) + 1 + $x] = [byte](($pix[$y * $w + $x] - $up) -band 0xFF) } }
        $f = Join-Path $script:Tmp 'scan-flate.pdf'
        New-TestImagePdf $f (ConvertTo-TestZlib $raw) "/Width $w /Height $h /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /DecodeParms << /Predictor 15 /Colors 1 /Columns $w >>"
        @((Read-HUQrCodes -Path $f).Texts) | Should -Be @($t)
    }
    It 'PDF mit gezeichnetem QR (HU.Pdf) und eingebettetem Scan-Bild: Bild wird gelesen' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        $pdf = New-HUPdf
        [void](Add-HUPdfPage $pdf)
        [void](Add-HUPdfQrCode -Pdf $pdf -Text $t -X 450 -Y 700 -Size 100)
        $jpg = Join-Path $script:Tmp 'q.jpg'; [IO.File]::WriteAllBytes($jpg, (ConvertTo-TestJpeg (New-TestQrBitmap $t) 80))
        Add-HUPdfImage -Pdf $pdf -Jpeg (ConvertTo-HUPdfJpeg -Path $jpg -MaxPx 2000) -X 50 -Y 50 -W 200 -H 200
        $f = Join-Path $script:Tmp 'hu.pdf'; [IO.File]::WriteAllBytes($f, (Get-HUPdfBytes $pdf))
        @((Read-HUQrCodes -Path $f).Texts) | Should -Be @($t)
    }
    It 'PDF ohne Bild und unlesbares Format: klarer Hinweis' {
        $pdf = New-HUPdf; [void](Add-HUPdfPage $pdf); Add-HUPdfText -Pdf $pdf -Text 'nur Text' -X 50 -Y 50
        $f = Join-Path $script:Tmp 'text.pdf'; [IO.File]::WriteAllBytes($f, (Get-HUPdfBytes $pdf))
        (Read-HUQrCodes -Path $f).Hint | Should -Match 'kein gescanntes Bild'
        $g = Join-Path $script:Tmp 'jbig2.pdf'
        New-TestImagePdf $g ([byte[]](1, 2, 3, 4)) '/Width 10 /Height 10 /ColorSpace /DeviceGray /BitsPerComponent 1 /Filter /JBIG2Decode'
        (Read-HUQrCodes -Path $g).Hint | Should -Match 'JBIG2'
    }
}

Describe 'Schutz gegen praeparierte Dateien' {
    It 'invertierter QR (hell auf dunkel) wird gelesen' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        $bmp = New-TestQrBitmap $t
        $w = $bmp.PixelWidth; $pix = New-Object byte[] ($w * $bmp.PixelHeight); $bmp.CopyPixels($pix, $w, 0)
        for ($i = 0; $i -lt $pix.Length; $i++) { $pix[$i] = [byte](255 - $pix[$i]) }
        $inv = [Windows.Media.Imaging.BitmapSource]::Create($w, $bmp.PixelHeight, 300, 300, [Windows.Media.PixelFormats]::Gray8, $null, $pix, $w)
        @(Read-HUQrFromBitmap -Bitmap $inv) | Should -Be @($t)
    }
    It 'Entpacken ist begrenzt' {
        $raw = New-Object byte[] 200000
        $z = ConvertTo-TestZlib $raw
        { & (Get-Module HU.Qr) { param($d) ConvertFrom-HUPdfFlate $d 1000 } $z } | Should -Throw '*zu groß*'
    }
    It 'unsinnige Bildgroesse und kaputte Zahlen brechen nichts ab' {
        $t = New-HUQrPayload -Prefix 'HUR1' -Data $script:Data
        $bmp = New-TestQrBitmap $t
        $f = Join-Path $script:Tmp 'mix.pdf'
        # zwei Bilder: eines mit Riesenmass, eines richtig
        $jpg = ConvertTo-TestJpeg $bmp 70
        $l1 = [Text.Encoding]::GetEncoding(28591); $ms = New-Object IO.MemoryStream
        $w = { param($x) $b = $l1.GetBytes($x); $ms.Write($b, 0, $b.Length) }
        & $w "%PDF-1.4`n7 0 obj << /Type /XObject /Subtype /Image /Width 99999999999 /Height 5 /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /Length 3 >>`nstream`nabc`nendstream`nendobj`n"
        & $w "8 0 obj << /Type /XObject /Subtype /Image /Width $($bmp.PixelWidth) /Height $($bmp.PixelHeight) /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /DCTDecode /Length $($jpg.Length) >>`nstream`n"
        $ms.Write($jpg, 0, $jpg.Length); & $w "`nendstream`nendobj`n%%EOF`n"
        [IO.File]::WriteAllBytes($f, $ms.ToArray())
        $r = Read-HUQrCodes -Path $f
        @($r.Texts) | Should -Be @($t)
    }
    It 'nur bekannte Bildformate, falsche Kennung im JPEG-Zweig wird abgelehnt' {
        $f = Join-Path $script:Tmp 'x.jpg'; [IO.File]::WriteAllBytes($f, [byte[]](1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
        (Read-HUQrCodes -Path $f).Hint | Should -Match 'nicht unterstützt'
        $g = Join-Path $script:Tmp 'fake-dct.pdf'
        New-TestImagePdf $g ([byte[]](0x89, 0x50, 0x4E, 0x47, 1, 2, 3, 4)) '/Width 10 /Height 10 /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /DCTDecode'
        (Read-HUQrCodes -Path $g).Hint | Should -Match 'nicht lesen'
    }
    It 'zu langer QR-Text wird nicht gelesen' {
        Read-HUQrPayload -Text ('HUR1:' + ('A' * 5000)) -Prefix 'HUR1' | Should -BeNullOrEmpty
    }
}
