#Requires -Version 5.1
# Pester 5: Excel und CSV lesen, xlsx schreiben (Core\HU.Import.psm1). Nur kuenstliche Dateien unter TEMP.
# Vergleich mit den bisherigen Lesern der Module: Umgebungsvariable HU_IMPORT_REFERENZ = Pfade zu den Modul-Skripten
# (durch ; getrennt). Ohne die Variable (z. B. in der CI des oeffentlichen Repos) werden diese Vergleiche uebersprungen.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $script:ModPfad = Join-Path $script:Root 'Core\HU.Import.psm1'
    Import-Module $script:ModPfad -Force -DisableNameChecking
    Add-Type -AssemblyName System.IO.Compression

    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtImport-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    [void](New-Item -ItemType Directory -Path $script:Tmp)

    # Bisherige Leser der Module (nur wenn angegeben)
    foreach ($p in @("$env:HU_IMPORT_REFERENZ" -split ';' | Where-Object { $_ -and (Test-Path -LiteralPath $_) })) { . $p }
    $script:HatKz = [bool](Get-Command Read-HUKzXlsx -CommandType Function -ErrorAction SilentlyContinue)
    $script:HatGs = [bool](Get-Command Read-HUGsXlsx -CommandType Function -ErrorAction SilentlyContinue)

    $script:Ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $script:Kopf = '<?xml version="1.0" encoding="UTF-8"?>'
    function New-TestZip([string]$Path, $Teile) {
        $fs = [IO.File]::Open($Path, 'Create')
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
        foreach ($k in $Teile.Keys) {
            $e = $zip.CreateEntry($k); $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding($false))); $w.Write($Teile[$k]); $w.Dispose()
        }
        $zip.Dispose(); $fs.Dispose()
    }
    # xlsx mit einem Blatt "Liste", wahlweise sharedStrings
    function New-TestXlsx([string]$Path, [string]$SheetXml, [string]$Sst = '') {
        $ct = $script:Kopf + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' + $(if ($Sst) { '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' }) + '</Types>'
        $wbr = $script:Kopf + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' + $(if ($Sst) { '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' }) + '</Relationships>'
        $teile = [ordered]@{
            '[Content_Types].xml'        = $ct
            '_rels/.rels'                = $script:Kopf + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            'xl/workbook.xml'            = $script:Kopf + "<workbook $($script:Ns)><sheets><sheet name=`"Liste`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
            'xl/_rels/workbook.xml.rels' = $wbr
            'xl/worksheets/sheet1.xml'   = $script:Kopf + "<worksheet $($script:Ns)>$SheetXml</worksheet>"
        }
        if ($Sst) { $teile['xl/sharedStrings.xml'] = $script:Kopf + "<sst $($script:Ns)>$Sst</sst>" }
        New-TestZip $Path $teile
    }
    function S([string]$ref, [string]$t) { "<c r=`"$ref`" t=`"inlineStr`"><is><t>$t</t></is></c>" }
    function N([string]$ref, [string]$v) { "<c r=`"$ref`"><v>$v</v></c>" }

    # --- Grenzfaelle (gleiche Faelle wie der bisherige Vergleich mit SheetJS) ---
    $script:FallDir = Join-Path $script:Tmp 'faelle'; [void](New-Item -ItemType Directory -Path $script:FallDir)
    $f = $script:FallDir
    $kopf3 = '<row r="3">' + (S 'A3' 'Kostenstelle') + (S 'B3' 'Name') + (S 'C3' 'Anlagennummer') + (S 'D3' 'Haushaltsstelle') + '</row>'
    $daten3 = '<row r="4">' + (S 'A4' 'K1') + (S 'B4' 'Halle') + (N 'C4' '1001') + (S 'D4' '1/000') + '</row>' +
    '<row r="5">' + (N 'C5' '1002') + '</row>' +
    '<row r="6">' + (S 'A6' 'K2') + (S 'B6' 'Schule') + (N 'C6' '2001') + '</row>'
    New-TestXlsx "$f\01.xlsx" ('<dimension ref="A1:D6"/><sheetData>' + $kopf3 + $daten3 + '</sheetData>')
    New-TestXlsx "$f\02.xlsx" ('<sheetData>' + $kopf3 + $daten3 + '</sheetData>')
    New-TestXlsx "$f\03.xlsx" ('<dimension ref="A3:D6"/><sheetData>' + $kopf3 + $daten3 + '</sheetData>')
    New-TestXlsx "$f\04.xlsx" ('<sheetData><row r="1"><c r="A1" s="1"/></row>' + $kopf3 + $daten3 + '</sheetData>')
    New-TestXlsx "$f\05.xlsx" ('<dimension ref="A1"/><sheetData>' + $kopf3 + $daten3 + '</sheetData>')
    New-TestXlsx "$f\06.xlsx" ('<dimension ref="A3:B4"/><sheetData>' + $kopf3 + $daten3 + '</sheetData>')
    $k7 = '<row r="1">' + (S 'B1' 'Kostenstelle') + (S 'C1' 'Name') + (S 'D1' 'Anlagennummer') + '</row><row r="2">' + (S 'A2' 'x') + (S 'B2' 'K1') + (N 'D2' '1') + '</row><row r="3">' + (S 'B3' 'K2') + (N 'D3' '2') + '</row>'
    New-TestXlsx "$f\07.xlsx" ('<sheetData>' + $k7 + '</sheetData>')
    New-TestXlsx "$f\08.xlsx" ('<dimension ref="B1:D3"/><sheetData>' + $k7 + '</sheetData>')
    $k9 = '<row r="1">' + (S 'A1' 'Kostenstelle') + (S 'B1' 'Name') + (S 'C1' 'Kostenstelle') + (S 'E1' 'Name_1') + (S 'F1' 'Name') + (S 'G1' '__EMPTY') + (S 'H1' 'KST') + (S 'I1' 'Kst') + (S 'J1' 'Kostenstelle') + '</row>' +
    '<row r="2">' + (S 'A2' 'K1') + (S 'B2' 'n1') + (S 'C2' 'K9') + (S 'D2' 'd') + (S 'E2' 'e') + (S 'F2' 'f') + (S 'G2' 'g') + (S 'H2' 'h') + (S 'I2' 'i') + (S 'J2' 'j') + '</row>' +
    '<row r="3">' + (S 'C3' 'K8') + (N 'D3' '5') + '</row>'
    New-TestXlsx "$f\09.xlsx" ('<sheetData>' + $k9 + '</sheetData>')
    # Zahl in der Datei -> String(zahl) in JavaScript (Werte von Hand nach ECMAScript Number::toString)
    $script:Zahlen = [ordered]@{
        '1234567890123456789' = '1234567890123456800'; '-3.5e-7' = '-3.5e-7'; '1e21' = '1e+21'; '123456789012345680000' = '123456789012345680000'
        '0.1' = '0.1'; '1e-6' = '0.000001'; '1.5e-6' = '0.0000015'; '-0' = '0'; '100' = '100'; '2.5E+3' = '2500'; '123e-20' = '1.23e-18'
        '1E+15' = '1000000000000000'; '999999999999999' = '999999999999999'; '12345678901234567' = '12345678901234568'
        '0.30000000000000004' = '0.30000000000000004'; '5e-324' = '5e-324'; '1.7976931348623157e308' = '1.7976931348623157e+308'
        '306123456789' = '306123456789'; '-12.5' = '-12.5'; '1e-7' = '1e-7'; '0.0000001234' = '1.234e-7'; '4.35' = '4.35'
        '9007199254740993' = '9007199254740992'; '2.675' = '2.675'; '1e20' = '100000000000000000000'; '1.000000000000001' = '1.000000000000001'
    }
    $r = '<row r="1">' + (S 'A1' 'Nr') + (S 'B1' 'Wert') + '</row>'; $i = 2
    foreach ($z in $script:Zahlen.Keys) { $r += "<row r=`"$i`">" + (N "A$i" "$i") + (N "B$i" $z) + '</row>'; $i++ }
    New-TestXlsx "$f\10.xlsx" ('<sheetData>' + $r + '</sheetData>')
    $sst = '<si><t>Kostenstelle</t></si><si><t>Name</t></si><si><r><t>K</t></r><r><t>7</t></r></si><si><t></t></si><si><t xml:space="preserve"> Sp </t></si>'
    $k11 = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>' + (S 'C1' 'Wert') + '</row>' +
    '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="e"><v>#N/A</v></c><c r="C2" t="b"><v>1</v></c></row>' +
    '<row r="3"><c r="A3" t="str"><f>""</f><v></v></c></row>' +
    '<row r="5"><c r="A5" t="s"><v>3</v></c></row>' +
    '<row r="6"><c r="A6" t="e"><v>#NULL!</v></c><c r="B6" t="s"><v>4</v></c><c r="C6" t="b"><v>0</v></c></row>' +
    '<row r="7"><c r="A7" t="str"><f>1+1</f><v>2</v></c><c r="C7"><f>1+1</f><v>2</v></c></row>' +
    '<row r="8"><c r="A8" s="2"/><c r="B8"><v></v></c></row>' +
    '<row r="9"><c r="C9" t="n"><v>7</v></c></row>'
    New-TestXlsx "$f\11.xlsx" ('<sheetData>' + $k11 + '</sheetData>') $sst
    New-TestXlsx "$f\12.xlsx" ('<sheetData><row r="1">' + (S 'A1' 'Kostenstelle') + '</row></sheetData>')
    New-TestXlsx "$f\13.xlsx" '<sheetData/>'
    New-TestXlsx "$f\14.xlsx" ('<sheetData><row><c t="inlineStr"><is><t>Kostenstelle</t></is></c><c t="inlineStr"><is><t>Name</t></is></c></row><row><c t="inlineStr"><is><t>K1</t></is></c><c><v>3</v></c></row></sheetData>')
    $sst15 = '<si><t xml:space="preserve"> Name </t></si><si><t> ohne preserve </t></si><si><r><t xml:space="preserve"> a </t></r><r><t xml:space="preserve"> b </t></r></si><si><t>Zeile1&#10;Zeile2</t></si><si><t>&amp;&lt;x&gt; &#252;_x0041_</t></si>'
    $k15 = '<row r="1"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> Kostenstelle </t></is></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>W</t></is></c></row>' +
    '<row r="2"><c r="A2" t="inlineStr"><is><t xml:space="preserve">  K1  </t></is></c><c r="B2" t="s"><v>1</v></c><c r="C2" t="s"><v>2</v></c></row>' +
    '<row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3" t="s"><v>4</v></c><c r="C3" t="inlineStr"><is><r><t xml:space="preserve"> r1 </t></r><r><t>r2 </t></r></is></c></row>'
    New-TestXlsx "$f\15.xlsx" ('<sheetData>' + $k15 + '</sheetData>') $sst15
    $k16 = '<row r="1">' + (S 'A1' 'Kostenstelle') + (S 'B1' 'Name') + '</row>' +
    '<row r="2"><c r="A2" t="e"><v>#N/A</v></c></row>' +
    '<row r="3"><c r="B3" t="e"><v>#NULL!</v></c></row>' +
    '<row r="4">' + (S 'A4' 'K4') + '</row><row r="30"/>' +
    '<row r="5"><c r="F5" s="3"/></row>'
    New-TestXlsx "$f\16.xlsx" ('<sheetData>' + $k16 + '</sheetData>')
    New-TestXlsx "$f\17.xlsx" ('<sheetData><row r="1">' + (N 'A1' '2024') + (N 'B1' '0.5') + '<c r="C1" t="b"><v>1</v></c></row><row r="2">' + (S 'A2' 'a') + (S 'B2' 'b') + (S 'C2' 'c') + '</row></sheetData>')
    New-TestXlsx "$f\18.xlsx" ('<sheetData><row r="1">' + (S 'A1' 'Kostenstelle') + (S 'B1' 'Name') + '</row><row r="2"><c r="A2"><f>X1</f></c><c r="H2"><v></v></c></row><row r="3" spans="1:2" ht="20" customHeight="1"/><row r="4"><c r="B4" t="str"><v></v></c></row><row r="40"></row></sheetData>')
    New-TestXlsx "$f\19.xlsx" ('<sheetData><row r="1">' + (S 'A1' 'Kostenstelle') + '</row><row r="2">' + (S 'A2' 'K') + '</row><row r="9" ht="20"></row></sheetData>')
    New-TestXlsx "$f\20.xlsx" ('<sheetData><row r="1">' + (S 'A1' 'Kostenstelle') + '</row><row r="2">' + (S 'A2' 'K') + '<c r="E2" t="str"><v></v></c><c r="G2" t="e"><v>#N/A</v></c></row></sheetData>')

    # --- Von Hand gebaute Mappe wie aus Excel: sharedStrings, Datum, Zahl, Wahrheitswert, Rich Text, zwei Blaetter ---
    $script:Serial = [int]([datetime]'2026-02-10').ToOADate()
    $script:Hand = Join-Path $script:Tmp 'Hand.xlsx'
    $hsst = '<si><t>Bezeichnung</t></si><si><t>Datum</t></si><si><t>Betrag</t></si><si><t>Erledigt</t></si>' +
    '<si><r><rPr><b/></rPr><t>Fett</t></r><r><t xml:space="preserve"> und normal</t></r><rPh sb="0" eb="1"><t>ふりがな</t></rPh></si>' +
    '<si><t>Straße_x000A_Größe &amp; Co</t></si>'
    $hs1 = '<dimension ref="A1:D4"/><sheetData>' +
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>' +
    "<row r=`"2`"><c r=`"A2`" t=`"s`"><v>4</v></c><c r=`"B2`" s=`"1`"><v>$($script:Serial)</v></c><c r=`"C2`" s=`"2`"><v>1234.5</v></c><c r=`"D2`" t=`"b`"><v>1</v></c></row>" +
    '<row r="4"><c r="A4" t="s"><v>5</v></c><c r="B4" s="1"><v>46064.5</v></c><c r="C4"><v>-0.25</v></c><c r="D4" t="b"><v>0</v></c><c r="E4" t="e"><v>#DIV/0!</v></c></row>' +
    '</sheetData>'
    $hs2 = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Zweites Blatt</t></is></c></row></sheetData>'
    New-TestZip $script:Hand ([ordered]@{
            '[Content_Types].xml'        = $script:Kopf + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
            '_rels/.rels'                = $script:Kopf + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
            'xl/workbook.xml'            = $script:Kopf + "<workbook $($script:Ns)><sheets><sheet name=`"Übersicht`" sheetId=`"1`" r:id=`"rId1`"/><sheet name=`"Zwei`" sheetId=`"2`" r:id=`"rId2`"/></sheets></workbook>"
            'xl/_rels/workbook.xml.rels' = $script:Kopf + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
            'xl/styles.xml'              = $script:Kopf + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00\ &quot;€&quot;"/></numFmts><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>'
            'xl/sharedStrings.xml'       = $script:Kopf + "<sst $($script:Ns) count=`"6`" uniqueCount=`"6`">$hsst</sst>"
            'xl/worksheets/sheet1.xml'   = $script:Kopf + "<worksheet $($script:Ns)>$hs1</worksheet>"
            'xl/worksheets/sheet2.xml'   = $script:Kopf + "<worksheet $($script:Ns)>$hs2</worksheet>"
        })

    # --- Vergleichstexte ---
    function Get-RowsText($Bl) { $t = foreach ($z in $Bl.Rows) { (@($z) -join [char]31) }; return "$($Bl.Rows.Count)|" + (@($t) -join [char]30) }
    function Get-WerteText($Bl) {
        $t = foreach ($r in $Bl.Werte.Keys) { "${r}:" + ((@($Bl.Werte[$r].Keys | Sort-Object) | ForEach-Object { "$_=$($Bl.Werte[$r][$_])" }) -join [char]31) }
        return (@($t) -join [char]30)
    }
    function Get-AnzeigeText($Bl) { return ((@($Bl.Anzeige.Keys) | Sort-Object -CaseSensitive | ForEach-Object { "$_=$($Bl.Anzeige[$_])" }) -join [char]30) }
    function Get-ObjText($Liste) { return ((@($Liste | Where-Object { $null -ne $_ }) | ForEach-Object { $h = $_; (@($h.Keys) | Sort-Object -CaseSensitive | ForEach-Object { "$_=$($h[$_])" }) -join [char]31 }) -join [char]30) }
    function Get-ZipTexte([byte[]]$Bytes) {
        $ms = New-Object IO.MemoryStream(, $Bytes); $z = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Read)
        try {
            $h = [ordered]@{}
            foreach ($e in $z.Entries) { $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); $h[$e.FullName] = $sr.ReadToEnd(); $sr.Dispose() }
            return $h
        } finally { $z.Dispose(); $ms.Dispose() }
    }
    # Blaetter zum Schreiben (kuenstlich, mit Umlauten, XML-Sonderzeichen, Steuerzeichen, leeren Zellen)
    $script:Schreib = @(
        [pscustomobject]@{ Name = 'Zuordnungen'; Kopf = [string[]]@('Nummer', 'Kostenstelle', 'Name / Bezeichnung'); Zeilen = @([string[]]@('AT0000000000000000000000000000001', 'K1', 'Halle & <Bühne>'), [string[]]@('', 'K2', ('Größe' + [char]1 + 'x')), [string[]]@('1001', '', '')) }
        [pscustomobject]@{ Name = 'Leer'; Kopf = [string[]]@('A'); Zeilen = @() }
    )
}

AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Excel lesen (Standard)' {
    It 'von Hand gebaute Mappe: sharedStrings, Datum als Seriennummer, Zahl, Wahrheitswert, Rich Text, Fehler, Lücke' {
        $bl = @(Read-HUImportFile $script:Hand)
        $bl.Count | Should -Be 2
        $bl[0].Name | Should -Be 'Übersicht'
        $bl[0].Rows.Count | Should -Be 4
        ($bl[0].Rows[0] -join '|') | Should -Be 'Bezeichnung|Datum|Betrag|Erledigt'
        ($bl[0].Rows[1] -join '|') | Should -Be "Fett und normal|$($script:Serial)|1234.5|WAHR"
        @($bl[0].Rows[2]).Count | Should -Be 0
        ($bl[0].Rows[3] -join '|') | Should -Be "Straße`nGröße & Co|46064.5|-0.25|FALSCH|"
        $bl[1].Name | Should -Be 'Zwei'
        $bl[1].Rows[0][0] | Should -Be 'Zweites Blatt'
        @($bl[0].PSObject.Properties.Name) -join ',' | Should -Be 'Name,Rows'
    }
    It 'Datum aus der Seriennummer zurückrechnen' {
        $bl = @(Read-HUImportFile $script:Hand)[0]
        [datetime]::FromOADate([double]::Parse($bl.Rows[1][1], [Globalization.CultureInfo]::InvariantCulture)) | Should -Be ([datetime]'2026-02-10')
        [datetime]::FromOADate([double]::Parse($bl.Rows[3][1], [Globalization.CultureInfo]::InvariantCulture)) | Should -Be ([datetime]'2026-02-11 12:00')
    }
    It 'Zeilengrenze -MaxRows' {
        $bl = @(Read-HUXlsx ([IO.File]::ReadAllBytes("$($script:FallDir)\10.xlsx")) -MaxRows 5)[0]
        $bl.Rows.Count | Should -Be 5
        $bl = @(Read-HUXlsx ([IO.File]::ReadAllBytes("$($script:FallDir)\10.xlsx")))[0]
        $bl.Rows.Count | Should -Be ($script:Zahlen.Count + 1)
    }
    It 'keine Excel-Datei oder Datei mit DTD wird abgelehnt' {
        { Read-HUXlsx ([byte[]](1, 2, 3, 4)) } | Should -Throw
        $p = Join-Path $script:Tmp 'dtd.xlsx'
        New-TestXlsx $p '<sheetData/>'
        $b = [IO.File]::ReadAllBytes($p)
        $ms = New-Object IO.MemoryStream; $ms.Write($b, 0, $b.Length)
        $z = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Update, $true)
        $z.GetEntry('xl/workbook.xml').Delete()
        $e = $z.CreateEntry('xl/workbook.xml'); $w = New-Object IO.StreamWriter($e.Open())
        $w.Write('<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa">]><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>'); $w.Dispose(); $z.Dispose()
        { Read-HUXlsx $ms.ToArray() } | Should -Throw
    }
    It 'nicht unterstützter Dateityp' {
        { Read-HUImportFile (Join-Path $script:Tmp 'x.xls') } | Should -Throw '*Dateityp .xls wird nicht unterstützt*'
    }
}

Describe 'Excel lesen wie SheetJS (-SheetJs)' {
    It 'Zahlen wie String(zahl) in JavaScript' {
        $bl = @(Read-HUXlsx ([IO.File]::ReadAllBytes("$($script:FallDir)\10.xlsx")) -SheetJs)[0]
        $i = 1
        foreach ($k in $script:Zahlen.Keys) { "$k -> $($bl.Rows[$i][1])" | Should -Be "$k -> $($script:Zahlen[$k])"; $i++ }
        ConvertTo-HUJsNumberText ([double]::NaN) | Should -Be 'NaN'
        ConvertTo-HUJsNumberText ([double]::NegativeInfinity) | Should -Be '-Infinity'
    }
    It 'Typen, Bereich und Zeilen-Objekte der von Hand gebauten Mappe' {
        $bl = @(Read-HUXlsx ([IO.File]::ReadAllBytes($script:Hand)) -SheetJs)[0]
        ($bl.Ref -join ',') | Should -Be '0,0,3,3'
        ($bl.Rows[1] -join '|') | Should -Be "Fett und normal|$($script:Serial)|1234.5|true"
        ($bl.Rows[3] -join '|') | Should -Be "Straße`nGröße & Co|46064.5|-0.25|false|#DIV/0!"
        $bl.Anzeige['1,3'] | Should -Be 'TRUE'
        $o = @(ConvertTo-HUXlsxRowObject $bl)
        $o.Count | Should -Be 2
        $o[0]['Bezeichnung'] | Should -Be 'Fett und normal'
        $o[0]['Erledigt'] | Should -Be 'true'
        $o[1]['Betrag'] | Should -Be '-0.25'
        $o[1].ContainsKey('E') | Should -BeFalse   # Spalte E liegt ausserhalb von <dimension>
    }
    It 'doppelte und leere Spaltennamen wie sheet_to_json' {
        $bl = @(Read-HUXlsx ([IO.File]::ReadAllBytes("$($script:FallDir)\09.xlsx")) -SheetJs)[0]
        $o = @(ConvertTo-HUXlsxRowObject $bl)
        (@($o[0].Keys) | Sort-Object -CaseSensitive) -join ',' | Should -Be ((@('Kostenstelle', 'Name', 'Kostenstelle_1', '__EMPTY', 'Name_1', 'Name_2', '__EMPTY_1', 'KST', 'Kst', 'Kostenstelle_2') | Sort-Object -CaseSensitive) -join ',')
        $o[1]['Kostenstelle_1'] | Should -Be 'K8'
    }
    It 'Blatt ohne Bereich liefert keine Zeilen-Objekte' {
        $bl = @(Read-HUXlsx ([IO.File]::ReadAllBytes("$($script:FallDir)\13.xlsx")) -SheetJs)[0]
        $bl.Ref | Should -BeNullOrEmpty
        @(ConvertTo-HUXlsxRowObject $bl).Count | Should -Be 0
    }
}

Describe 'CSV lesen' {
    It 'UTF-8 mit BOM, Semikolon, Anführungszeichen mit Zeilenumbruch und ""' {
        $b = (New-Object Text.UTF8Encoding $true).GetPreamble() + [Text.Encoding]::UTF8.GetBytes("Name;Text;Wert`r`nMüller;""a;b`r`nc"";""5,00""`r`n""Zitat """"x"""""";;`r`n")
        $bl = @(Read-HUCsv $b)[0]
        $bl.Name | Should -Be 'CSV'
        $bl.Rows.Count | Should -Be 3
        $bl.Rows[1][0] | Should -Be 'Müller'
        $bl.Rows[1][1] | Should -Be "a;b`r`nc"   # CR in Anfuehrungszeichen bleibt wie bisher erhalten
        $bl.Rows[1][2] | Should -Be '5,00'
        $bl.Rows[2][0] | Should -Be 'Zitat "x"'
        @($bl.Rows[2]).Count | Should -Be 3
    }
    It 'Windows-1252, UTF-8 ohne BOM und UTF-16 mit BOM' {
        $t = "Ort,Straße`nMusterort,Hauptstraße 1`n"
        (@(Read-HUCsv ([Text.Encoding]::GetEncoding(1252).GetBytes($t)))[0].Rows[1][1]) | Should -Be 'Hauptstraße 1'
        (@(Read-HUCsv ([Text.Encoding]::UTF8.GetBytes($t)))[0].Rows[1][1]) | Should -Be 'Hauptstraße 1'
        $u = [Text.Encoding]::Unicode.GetPreamble() + [Text.Encoding]::Unicode.GetBytes("A`tB`r`nÄ`tÖ`r`n")
        $bl = @(Read-HUCsv $u)[0]
        ($bl.Rows[1] -join '|') | Should -Be 'Ä|Ö'
    }
    It 'Trennzeichen erkannt (Tabulator, Komma) oder fest vorgegeben' {
        (@(Read-HUCsv ([Text.Encoding]::UTF8.GetBytes("a`tb`tc`n1`t2`t3")))[0].Rows[1] -join '|') | Should -Be '1|2|3'
        (@(Read-HUCsv ([Text.Encoding]::UTF8.GetBytes("a,b;c`n1,2;3")))[0].Rows[1] -join '|') | Should -Be '1,2|3'
        (@(Read-HUCsv ([Text.Encoding]::UTF8.GetBytes("a,b;c`n1,2;3")) -Delimiter ',')[0].Rows[1] -join '|') | Should -Be '1|2;3'
    }
    It 'eine Zeile bleibt eine Zeile, leere Datei hat keine Zeilen' {
        $bl = @(Read-HUCsv ([Text.Encoding]::UTF8.GetBytes('a;b;c')))[0]
        $bl.Rows.Count | Should -Be 1
        ($bl.Rows[0] -join '|') | Should -Be 'a|b|c'
        @(Read-HUCsv ([byte[]]@()))[0].Rows.Count | Should -Be 0
    }
    It 'CSV-Datei über Read-HUImportFile (.csv und .txt)' {
        $p = Join-Path $script:Tmp 'Liste.csv'
        [IO.File]::WriteAllText($p, "Von;Bis`r`n6001;6010`r`n", (New-Object Text.UTF8Encoding $true))
        (@(Read-HUImportFile $p)[0].Rows[1] -join '|') | Should -Be '6001|6010'
        Copy-Item -LiteralPath $p -Destination (Join-Path $script:Tmp 'Liste.txt')
        @(Read-HUImportFile (Join-Path $script:Tmp 'Liste.txt'))[0].Rows.Count | Should -Be 2
    }
}

Describe 'Excel schreiben' {
    It 'xlsx bauen und wieder lesen (Umlaute, Sonderzeichen, Steuerzeichen entfernt, leere Zellen)' {
        $b = New-HUXlsxBytes $script:Schreib
        $b.GetType().Name | Should -Be 'Byte[]'
        $b[0] | Should -Be 0x50
        $bl = @(Read-HUXlsx $b)
        $bl.Count | Should -Be 2
        $bl[0].Name | Should -Be 'Zuordnungen'
        ($bl[0].Rows[0] -join '|') | Should -Be 'Nummer|Kostenstelle|Name / Bezeichnung'
        ($bl[0].Rows[1] -join '|') | Should -Be 'AT0000000000000000000000000000001|K1|Halle & <Bühne>'
        ($bl[0].Rows[2] -join '|') | Should -Be '|K2|Größex'
        ($bl[0].Rows[3] -join '|') | Should -Be '1001'
        $bl[1].Rows.Count | Should -Be 1
        $x = Get-ZipTexte $b
        $x['xl/worksheets/sheet1.xml'] | Should -Match '<autoFilter ref="A1:C4"/>'
        $x['xl/worksheets/sheet1.xml'] | Should -Match 'state="frozen"'
        $x['xl/worksheets/sheet2.xml'] | Should -Not -Match 'autoFilter'
    }
    It 'Parametername Blaetter wie bisher' {
        (New-HUXlsxBytes -Blaetter $script:Schreib).Length | Should -BeGreaterThan 500
    }
    It 'Spaltennamen und -index' {
        Get-HUXlsxColumnName 0 | Should -Be 'A'
        Get-HUXlsxColumnName 27 | Should -Be 'AB'
        Get-HUXlsxColumnName 16383 | Should -Be 'XFD'
        Get-HUXlsxColumnIndex 'XFD1' | Should -Be 16383
        (ConvertFrom-HUXlsxRange '$B$2:D6') -join ',' | Should -Be '1,1,5,3'
        ConvertFrom-HUXlsxRange 'A1' | Should -BeNullOrEmpty
    }
}

Describe 'Gleichheit mit den bisherigen Lesern der Module' {
    It 'Leser mit SheetJS-Werten: Rows, Bereich, Werte, Anzeige, Zeilen-Objekte gleich (alle Grenzfälle)' {
        if (-not $script:HatKz) { Set-ItResult -Skipped -Because 'HU_IMPORT_REFERENZ ohne bisherigen Leser mit SheetJS-Werten'; return }
        $dateien = @(Get-ChildItem -LiteralPath $script:FallDir -Filter *.xlsx | Sort-Object Name) + @(Get-Item -LiteralPath $script:Hand)
        foreach ($d in $dateien) {
            $b = [IO.File]::ReadAllBytes($d.FullName)
            $alt = @(Read-HUKzXlsx $b); $neu = @(Read-HUXlsx $b -SheetJs)
            "$($d.Name) Blätter $($neu.Count)" | Should -Be "$($d.Name) Blätter $($alt.Count)"
            for ($i = 0; $i -lt $alt.Count; $i++) {
                "$($d.Name) Name $($neu[$i].Name)" | Should -Be "$($d.Name) Name $($alt[$i].Name)"
                "$($d.Name) Ref $($neu[$i].Ref -join ',')" | Should -Be "$($d.Name) Ref $($alt[$i].Ref -join ',')"
                "$($d.Name) $(Get-RowsText $neu[$i])" | Should -BeExactly "$($d.Name) $(Get-RowsText $alt[$i])"
                "$($d.Name) $(Get-WerteText $neu[$i])" | Should -BeExactly "$($d.Name) $(Get-WerteText $alt[$i])"
                "$($d.Name) $(Get-AnzeigeText $neu[$i])" | Should -BeExactly "$($d.Name) $(Get-AnzeigeText $alt[$i])"
                "$($d.Name) $(Get-ObjText (ConvertTo-HUXlsxRowObject $neu[$i]))" | Should -BeExactly "$($d.Name) $(Get-ObjText (ConvertTo-HUKzRowObjects $alt[$i]))"
                $neu[$i].Werte.GetType().FullName | Should -Be $alt[$i].Werte.GetType().FullName
            }
        }
    }
    It 'Leser mit SheetJS-Werten: Zahlen als Text gleich' {
        if (-not $script:HatKz) { Set-ItResult -Skipped -Because 'HU_IMPORT_REFERENZ ohne bisherigen Leser mit SheetJS-Werten'; return }
        $rnd = New-Object Random 4711
        $werte = @(0.1, 1e21, 1e-7, 5e-324, 123456.789, -2.675) + @(1..300 | ForEach-Object { [BitConverter]::Int64BitsToDouble([long]($rnd.NextDouble() * [long]::MaxValue)) })
        foreach ($w in $werte) { ConvertTo-HUJsNumberText $w | Should -BeExactly (ConvertTo-HUKzJsZahl $w) }
    }
    It 'Excel schreiben: geschriebene xlsx gleich (XML-Teile)' {
        if (-not $script:HatKz) { Set-ItResult -Skipped -Because 'HU_IMPORT_REFERENZ ohne bisherigen Leser mit SheetJS-Werten'; return }
        $alt = Get-ZipTexte (New-HUKzXlsxBytes $script:Schreib); $neu = Get-ZipTexte (New-HUXlsxBytes $script:Schreib)
        (@($neu.Keys) -join ',') | Should -Be (@($alt.Keys) -join ',')
        foreach ($k in $alt.Keys) { "$k $($neu[$k])" | Should -BeExactly "$k $($alt[$k])" }
    }
    It 'Leser Standard: Rows gleich (alle Grenzfälle und Hand-Mappe)' {
        if (-not $script:HatGs) { Set-ItResult -Skipped -Because 'HU_IMPORT_REFERENZ ohne bisherigen Standard-Leser'; return }
        $dateien = @(Get-ChildItem -LiteralPath $script:FallDir -Filter *.xlsx | Sort-Object Name) + @(Get-Item -LiteralPath $script:Hand)
        foreach ($d in $dateien) {
            $b = [IO.File]::ReadAllBytes($d.FullName)
            $alt = @(Read-HUGsXlsx $b); $neu = @(Read-HUXlsx $b)
            "$($d.Name) Blätter $($neu.Count)" | Should -Be "$($d.Name) Blätter $($alt.Count)"
            for ($i = 0; $i -lt $alt.Count; $i++) {
                "$($d.Name) Name $($neu[$i].Name)" | Should -Be "$($d.Name) Name $($alt[$i].Name)"
                "$($d.Name) $(Get-RowsText $neu[$i])" | Should -BeExactly "$($d.Name) $(Get-RowsText $alt[$i])"
            }
        }
    }
    It 'Leser Standard: CSV gleich (mehrzeilige Dateien)' {
        if (-not $script:HatGs) { Set-ItResult -Skipped -Because 'HU_IMPORT_REFERENZ ohne bisherigen Standard-Leser'; return }
        $faelle = @(
            [byte[]]((New-Object Text.UTF8Encoding $true).GetPreamble() + [Text.Encoding]::UTF8.GetBytes("Von;Bis;Datum;Wert`r`n6001;6010;01.03.2026;""5,00""`r`n6020;;02.03.2026;`r`n")),
            [byte[]][Text.Encoding]::GetEncoding(1252).GetBytes("Nr,Empfänger`n1,Müller`n2,""Huber, Anna""`n"),
            [byte[]][Text.Encoding]::UTF8.GetBytes("a`tb`n""x`ny""`tz`n`n")
        )
        $faelle.Count | Should -Be 3
        foreach ($b in $faelle) {
            $alt = @(Read-HUGsCsv ([byte[]]$b))[0]; $neu = @(Read-HUCsv ([byte[]]$b))[0]
            Get-RowsText $neu | Should -BeExactly (Get-RowsText $alt)
        }
    }
}

Describe 'Betrieb' {
    It 'läuft im nackten Runspace ohne App (nur das Kern-Modul)' {
        $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
        $ps = [powershell]::Create(); $ps.Runspace = $rs
        try {
            [void]$ps.AddScript('param($m, $p) Set-StrictMode -Version 1.0; Import-Module $m -DisableNameChecking; $b = @(Read-HUXlsx ([IO.File]::ReadAllBytes($p)) -SheetJs); "$($b[0].Name)|$($b[0].Rows[1][0])|$(@(ConvertTo-HUXlsxRowObject $b[0]).Count)|$((New-HUXlsxBytes @([pscustomobject]@{ Name = ''x''; Kopf = [string[]]@(''a''); Zeilen = @() })).Length -gt 0)"').AddArgument($script:ModPfad).AddArgument($script:Hand)
            $r = $ps.Invoke()
            $ps.HadErrors | Should -BeFalse
            "$($r[0])" | Should -Be 'Übersicht|Fett und normal|2|True'
        } finally { $ps.Dispose(); $rs.Dispose() }
    }
    It 'Moduldatei UTF-8 mit BOM und CRLF' {
        $b = [IO.File]::ReadAllBytes($script:ModPfad)
        ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) | Should -BeTrue
        $t = [Text.Encoding]::UTF8.GetString($b)
        ([regex]::Matches($t, '(?<!\r)\n')).Count | Should -Be 0
        $t.TrimEnd() | Should -Match 'ConvertFrom-HUXlsxRange$'
    }
}
