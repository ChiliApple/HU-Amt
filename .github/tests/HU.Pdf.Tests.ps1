#Requires -Version 5.1
# Pester 5: PDF-Schreiber (Core\HU.Pdf.psm1). Nur kuenstliche Texte und ein kuenstlich erzeugtes Bild, kein PDF wird geoeffnet.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:ModPath = Join-Path $script:Root 'Core\HU.Pdf.psm1'
    Import-Module $script:ModPath -Force -DisableNameChecking
    Add-Type -AssemblyName PresentationCore, WindowsBase
    $script:L1 = [Text.Encoding]::GetEncoding(28591)   # ISO-8859-1: ein Byte = ein Zeichen
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('HU.Pdf.Tests_' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $script:Tmp)

    # Kuenstliches JPEG (Farbverlauf) erzeugen, Rueckgabe: Bytes
    function New-TestJpeg([int]$W, [int]$H) {
        $stride = $W * 3
        $px = New-Object byte[] ($stride * $H)
        for ($y = 0; $y -lt $H; $y++) { for ($x = 0; $x -lt $W; $x++) { $o = $y * $stride + $x * 3; $px[$o] = [byte]($x % 256); $px[$o + 1] = [byte]($y % 256); $px[$o + 2] = 128 } }
        $src = [System.Windows.Media.Imaging.BitmapSource]::Create($W, $H, 96, 96, [System.Windows.Media.PixelFormats]::Bgr24, $null, $px, $stride)
        $enc = New-Object System.Windows.Media.Imaging.JpegBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($src))
        $ms = New-Object IO.MemoryStream
        $enc.Save($ms)
        return , $ms.ToArray()
    }
    # xref pruefen: jeder Eintrag 20 Byte, Offset zeigt genau auf "n 0 obj". Rueckgabe: Liste der Fehler (leer = gut)
    function Test-Xref([byte[]]$Bytes) {
        $s = $script:L1.GetString($Bytes)
        $err = New-Object System.Collections.Generic.List[string]
        $m = [regex]::Match($s, "startxref\n(\d+)\n%%EOF\n$")
        if (-not $m.Success) { $err.Add('startxref fehlt'); return , $err.ToArray() }
        $pos = [int]$m.Groups[1].Value
        if ($s.Substring($pos, 5) -ne "xref`n") { $err.Add("startxref zeigt nicht auf xref ($pos)"); return , $err.ToArray() }
        $hm = [regex]::Match($s.Substring($pos + 5), '^0 (\d+)\n')
        $n = [int]$hm.Groups[1].Value
        $p = $pos + 5 + $hm.Length
        $size = [int]([regex]::Match($s, '/Size (\d+)').Groups[1].Value)
        if ($size -ne $n) { $err.Add("Size $size <> $n") }
        for ($k = 0; $k -lt $n; $k++) {
            $e = $s.Substring($p + 20 * $k, 20)
            if ($e -notmatch '^\d{10} \d{5} [nf] \n$') { $err.Add("Eintrag $k falsch: [$e]"); continue }
            if ($k -eq 0) { continue }
            $off = [long]$e.Substring(0, 10)
            if ($s.Substring($off, "$k 0 obj`n".Length) -ne "$k 0 obj`n") { $err.Add("Objekt $k nicht bei $off") }
        }
        if ($s.Substring($p + 20 * $n, 8) -ne "trailer`n") { $err.Add('trailer nicht direkt nach xref') }
        return , $err.ToArray()
    }
    function Get-PageText($Page) { return (($Page.Ops | ForEach-Object { $script:L1.GetString($_) }) -join '') }

    $script:JpgPath = Join-Path $script:Tmp 'bild.jpg'
    [IO.File]::WriteAllBytes($script:JpgPath, (New-TestJpeg 300 150))
}

AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force }
}

Describe 'WinAnsi' {
    It 'Umlaute, Euro, Gedankenstrich, Anfuehrungszeichen und Auslassung' {
        $b = ConvertTo-HUPdfBytes ("Äöß " + [char]0x20AC + ' ' + [char]0x2013 + ' ' + [char]0x201E + 'x' + [char]0x201C + [char]0x2026)
        ($b | ForEach-Object { '{0:X2}' -f $_ }) -join ' ' | Should -Be 'C4 F6 DF 20 80 20 96 20 84 78 93 85'
    }
    It 'Ersatz, unbekannte Zeichen, Emoji und Steuerzeichen' {
        $b = ConvertTo-HUPdfBytes ("a" + [char]0x2192 + "b" + [char]0x03A9 + [char]0xD83D + [char]0xDE00 + "`tc")
        $script:L1.GetString($b) | Should -Be 'a->b?c'
    }
    It 'liefert immer byte[]' {
        (ConvertTo-HUPdfBytes 'A').GetType().Name | Should -Be 'Byte[]'
        (ConvertTo-HUPdfBytes '').Length | Should -Be 0
    }
}

Describe 'Breiten, Kuerzen, Umbruch' {
    It 'misst W breiter als i, fett breiter, linear zur Groesse' {
        (Get-HUPdfTextWidth 'W' 10) | Should -BeGreaterThan (Get-HUPdfTextWidth 'i' 10)
        (Get-HUPdfTextWidth 'Abcdef' 10 -Bold) | Should -BeGreaterThan (Get-HUPdfTextWidth 'Abcdef' 10)
        [Math]::Round((Get-HUPdfTextWidth 'Abc' 20) / (Get-HUPdfTextWidth 'Abc' 10), 6) | Should -Be 2
    }
    It 'Arial-Masse wie Helvetica (i = 0,222 em), wenn Arial vorhanden ist' {
        $w = Get-HUPdfTextWidth 'i' 1000
        if ([Math]::Abs($w - 556) -lt 0.01) { Set-ItResult -Skipped -Because 'Arial fehlt, Ersatzbreite aktiv'; return }
        [Math]::Round($w) | Should -Be 222
    }
    It 'kuerzt mit Auslassungszeichen und haelt die Breite ein' {
        $t = Get-HUPdfFitText 'Ein ziemlich langer Text für eine schmale Spalte' 60 9
        $t | Should -Match ([string][char]0x2026 + '$')
        (Get-HUPdfTextWidth $t 9) | Should -BeLessOrEqual 60
        Get-HUPdfFitText 'kurz' 60 9 | Should -Be 'kurz'
    }
    It 'bricht an Leerzeichen und Zeilenumbruch um' {
        $txt = "Erste Zeile mit einigen Wörtern zum Umbrechen`nZweiter Absatz"
        $l = Split-HUPdfTextLines $txt 80 9
        $l.Count | Should -BeGreaterThan 2
        foreach ($z in $l) { if ($z -match ' ') { (Get-HUPdfTextWidth $z 9) | Should -BeLessOrEqual 80 } }
        $l[-1] | Should -Be 'Zweiter Absatz'
        (@(Split-HUPdfTextLines 'eins' 200 9)).Count | Should -Be 1
        (Split-HUPdfTextLines 'eins' 200 9).GetType().Name | Should -Be 'String[]'
    }
}

Describe 'PDF-Aufbau' {
    BeforeAll {
        $pdf = New-HUPdf
        [void](Add-HUPdfPage $pdf)
        Add-HUPdfText $pdf 56 64 ('Größe ' + [char]0x20AC + ' 12,50 (Prüfung) a\b') 12 -Bold
        Add-HUPdfText $pdf 539 80 'rechts' 9 -Align right
        Add-HUPdfLine $pdf 56 90 539 90 0.8
        Add-HUPdfRect $pdf 56 100 200 20 0.9
        [void](Add-HUPdfPage $pdf)
        Add-HUPdfText $pdf 56 64 'Seite zwei' 9
        $script:Doc = $pdf
        $script:Bytes = Get-HUPdfBytes $pdf ('Übersicht (Test) ' + [char]0x2013 + ' März') -Producer 'HU-Amt Prüfmodul'
        $script:S = $script:L1.GetString($script:Bytes)
    }
    It 'liefert byte[] mit gueltigem Kopf und Ende' {
        $script:Bytes.GetType().Name | Should -Be 'Byte[]'
        $script:S.Substring(0, 9) | Should -Be "%PDF-1.4`n"
        $script:Bytes[10] | Should -BeGreaterThan 127   # Binaer-Kommentar
        $script:S | Should -Match "%%EOF`n$"
    }
    It 'xref-Offsets stimmen byte-genau' {
        @(Test-Xref $script:Bytes) | Should -BeNullOrEmpty
    }
    It 'Seitenbaum, Schriften WinAnsi, MediaBox A4' {
        $script:S | Should -Match '/Type /Pages /Kids \[5 0 R 7 0 R\] /Count 2'
        $script:S | Should -Match '/BaseFont /Helvetica /Encoding /WinAnsiEncoding'
        $script:S | Should -Match '/BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding'
        $script:S | Should -Match '/MediaBox \[0 0 595.28 841.89\]'
    }
    It 'Umlaute und Euro stehen als WinAnsi-Bytes im Inhalt, Klammern und Backslash maskiert' {
        $exp = $script:L1.GetString([byte[]](0x47, 0x72, 0xF6, 0xDF, 0x65, 0x20, 0x80)) + ' 12,50 \(Pr' + [char]0xFC + 'fung\) a\\b) Tj'
        $script:S.Contains($exp) | Should -BeTrue
        $script:S.Contains('/F2 12 Tf 56 777.89 Td') | Should -BeTrue   # Y von oben: 841.89 - 64
    }
    It 'Content-Laengen stimmen' {
        foreach ($m in [regex]::Matches($script:S, "<< /Length (\d+) >>\nstream\n")) {
            $start = $m.Index + $m.Length
            $script:S.Substring($start + [int]$m.Groups[1].Value, 10) | Should -Be "`nendstream"
        }
    }
    It 'Titel und Producer als Parameter, WinAnsi und maskiert' {
        $script:S.Contains('/Title (' + [char]0xDC + 'bersicht \(Test\) ' + [char]0x96 + ' M' + [char]0xE4 + 'rz)') | Should -BeTrue
        $script:S.Contains('/Producer (HU-Amt Pr' + [char]0xFC + 'fmodul)') | Should -BeTrue
        $s2 = $script:L1.GetString((Get-HUPdfBytes $script:Doc))
        $s2.Contains('/Producer (HU-Amt)') | Should -BeTrue
    }
    It 'rechtsbuendiger Text endet am rechten Rand' {
        $t = Get-PageText $script:Doc.Pages[0]
        $x = 539 - (Get-HUPdfTextWidth 'rechts' 9)
        $t.Contains("Tf $($x.ToString('0.##', [Globalization.CultureInfo]::InvariantCulture)) ") | Should -BeTrue
    }
    It 'Dokument ohne Seite ergibt eine leere Seite und gueltige xref' {
        $b = Get-HUPdfBytes (New-HUPdf)
        @(Test-Xref $b) | Should -BeNullOrEmpty
        $script:L1.GetString($b) | Should -Match '/Count 1'
    }
}

Describe 'Bilder' {
    It 'verkleinert auf MaxPx und liefert JPEG' {
        $j = ConvertTo-HUPdfJpeg -Path $script:JpgPath -MaxPx 100
        $j.PxW | Should -Be 100
        $j.PxH | Should -Be 50
        $j.Bytes[0] | Should -Be 0xFF
        $j.Bytes[1] | Should -Be 0xD8
    }
    It 'nimmt auch Bytes, ohne Verkleinerung bleibt die Groesse' {
        $j = ConvertTo-HUPdfJpeg -Bytes ([IO.File]::ReadAllBytes($script:JpgPath))
        $j.PxW | Should -Be 300
        $j.PxH | Should -Be 150
    }
    It 'bettet das Bild einmal ein, auch wenn es zweimal gesetzt wird' {
        $j = ConvertTo-HUPdfJpeg -Path $script:JpgPath -MaxPx 200
        $pdf = New-HUPdf
        [void](Add-HUPdfPage $pdf)
        Add-HUPdfText $pdf 56 64 'Beleg' 10
        Add-HUPdfImage $pdf $j 56 80 200 100
        Add-HUPdfImage $pdf $j 56 200 100 50
        [void](Add-HUPdfPage $pdf)
        Add-HUPdfImage $pdf $j 56 80 400 200
        $b = Get-HUPdfBytes $pdf 'Bild'
        $s = $script:L1.GetString($b)
        @(Test-Xref $b) | Should -BeNullOrEmpty
        $hdr = "5 0 obj`n<< /Type /XObject /Subtype /Image /Width 200 /Height 100 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length $($j.Bytes.Length) >>`nstream`n"
        $i = $s.IndexOf($hdr)
        $i | Should -BeGreaterThan 0
        $start = $i + $hdr.Length
        $emb = New-Object byte[] $j.Bytes.Length
        [Array]::Copy($b, $start, $emb, 0, $emb.Length)
        [Convert]::ToBase64String($emb) | Should -Be ([Convert]::ToBase64String($j.Bytes))
        $s.Substring($start + $j.Bytes.Length, 10) | Should -Be "`nendstream"
        ([regex]::Matches($s, '/Subtype /Image')).Count | Should -Be 1
        ([regex]::Matches($s, '/XObject << /Im1 5 0 R >>')).Count | Should -Be 2
        (Get-PageText $pdf.Pages[0]) | Should -Match 'q 200 0 0 100 56 661.89 cm /Im1 Do Q'
    }
}

Describe 'Tabelle ueber mehrere Seiten' {
    BeforeAll {
        $pdf = New-HUPdf -Landscape
        $cols = @(
            @{ Name = 'Nr.'; Key = 'nr'; Align = 'right' }
            @{ Name = 'Bezeichnung'; Key = 'text' }
            @{ Name = 'Betrag'; Key = 'betrag'; Align = 'right' }
        )
        $rows = @(for ($i = 1; $i -le 60; $i++) { @{ nr = "$i"; text = "Testposten $i Überprüfung"; betrag = ('{0},00 ' -f ($i * 10)) + [char]0x20AC } })
        $rows[3] = [pscustomobject]@{ nr = '4'; text = ('Sehr langer Text ' * 40).Trim(); betrag = '40,00' }
        $script:TabY = Add-HUPdfTable -Pdf $pdf -Columns $cols -Rows $rows -SumRow @{ nr = ''; text = 'Summe'; betrag = '18.300,00' } -Title 'Testauswertung' -Subtitle 'Künstliche Daten' -Heading 'Kopfzeile'
        Add-HUPdfPageFooter -Pdf $pdf -Text 'Erstellt am 01.01.2026 12:00'
        $script:Tab = $pdf
        $script:TabBytes = Get-HUPdfBytes $pdf 'Tabelle'
    }
    It 'ist im Querformat' {
        $script:Tab.W | Should -Be 841.89
        $script:L1.GetString($script:TabBytes) | Should -Match '/MediaBox \[0 0 841.89 595.28\]'
    }
    It 'laeuft ueber mindestens zwei Seiten mit gueltiger xref' {
        $script:Tab.Pages.Count | Should -BeGreaterOrEqual 2
        @(Test-Xref $script:TabBytes) | Should -BeNullOrEmpty
    }
    It 'wiederholt Kopfzeile und Ueberschriften auf jeder Seite' {
        foreach ($p in $script:Tab.Pages) {
            $t = Get-PageText $p
            $t | Should -Match '/F2 8.5 Tf [\d.]+ [\d.]+ Td \(Betrag\) Tj'
            $t | Should -Match '\(Bezeichnung\) Tj'
            $t | Should -Match '\(Testauswertung\) Tj'
            $t | Should -Match '0.88 g '
        }
    }
    It 'verteilt alle Zeilen, Summe fett auf der letzten Seite, Fusszeile mit Seitenzahl' {
        $all = (@($script:Tab.Pages | ForEach-Object { Get-PageText $_ }) -join '')
        foreach ($i in 1, 30, 60) { $all | Should -Match "\(Testposten $i " }
        $n = $script:Tab.Pages.Count
        $last = Get-PageText $script:Tab.Pages[$n - 1]
        $last | Should -Match '/F2 8.5 Tf [\d.]+ [\d.]+ Td \(Summe\) Tj'
        $last | Should -Match "\(Seite $n von $n\) Tj"
        (Get-PageText $script:Tab.Pages[0]) | Should -Match "\(Seite 1 von $n\) Tj"
        $script:Tab.Cur | Should -Be $script:Tab.Pages[$n - 1]
        $script:TabY | Should -BeLessThan $script:Tab.H
    }
    It 'kuerzt zu lange Zellen mit Auslassung (Spaltenbreite begrenzt)' {
        $t = Get-PageText $script:Tab.Pages[0]
        $t.Contains('Sehr langer Text Sehr') | Should -BeTrue
        $t.Contains([string][char]0x85 + ') Tj') | Should -BeTrue
    }
    It 'Arrays als Zeilen, leere Tabelle mit Hinweis, Fortsetzung auf der aktuellen Seite' {
        $pdf = New-HUPdf
        [void](Add-HUPdfPage $pdf)
        Add-HUPdfText $pdf 40 60 'Vorspann' 10
        $y = Add-HUPdfTable -Pdf $pdf -Columns @(@{ Name = 'A' }, @{ Name = 'B'; Width = 100 }) -Rows @(, @('x1', 'y1')) -Y 90
        $pdf.Pages.Count | Should -Be 1
        $y | Should -Be 119
        $t = Get-PageText $pdf.Pages[0]
        $t | Should -Match '\(x1\) Tj'
        $t | Should -Match '\(y1\) Tj'
        $y2 = Add-HUPdfTable -Pdf $pdf -Columns @(@{ Name = 'A' }) -Rows @() -Title 'Leer'
        $pdf.Pages.Count | Should -Be 2
        (Get-PageText $pdf.Pages[1]) | Should -Match ('\(Keine Eintr' + [char]0xE4 + 'ge.\) Tj')
        $y2 | Should -BeGreaterThan 0
    }
}

Describe 'Nackter Runspace' {
    It 'laeuft ohne Fenster und ohne App-Variablen im MTA-Runspace (Messen und Bild)' {
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = 'MTA'
        $rs.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $rs
        try {
            [void]$ps.AddScript({
                    param($Mod, $Jpg)
                    Set-StrictMode -Version 1.0
                    Import-Module $Mod -DisableNameChecking
                    $pdf = New-HUPdf
                    [void](Add-HUPdfPage $pdf)
                    Add-HUPdfText $pdf 56 64 'Rechts ausgerichtet' 10 -Align right
                    $j = ConvertTo-HUPdfJpeg -Path $Jpg -MaxPx 50
                    Add-HUPdfImage $pdf $j 56 80 50 25
                    [void](Add-HUPdfTable -Pdf $pdf -Columns @(@{ Name = 'A' }) -Rows @(@{ A = 'eins' }) -Y 140)
                    $b = Get-HUPdfBytes $pdf 'Runspace'
                    [pscustomobject]@{ Head = [Text.Encoding]::ASCII.GetString($b, 0, 8); Len = $b.Length; W = (Get-HUPdfTextWidth 'W' 10); PxW = $j.PxW; Apt = [Threading.Thread]::CurrentThread.GetApartmentState().ToString() }
                }).AddArgument($script:ModPath).AddArgument($script:JpgPath)
            $r = @($ps.Invoke())
            if ($ps.Streams.Error.Count) { throw $ps.Streams.Error[0].Exception }
        } finally { $ps.Dispose(); $rs.Dispose() }
        $r.Count | Should -Be 1
        $r[0].Head | Should -Be '%PDF-1.4'
        $r[0].Len | Should -BeGreaterThan 500
        $r[0].W | Should -BeGreaterThan 5
        $r[0].PxW | Should -Be 50
        $r[0].Apt | Should -Be 'MTA'
    }
}
