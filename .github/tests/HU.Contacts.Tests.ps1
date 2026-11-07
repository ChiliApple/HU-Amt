#Requires -Version 5.1
# Pester 5: Kontakte (Core\HU.Contacts.psm1, Outlook-Kontakte in Core\HU.Office.psm1). Nur kuenstliche Kontakte, kein Outlook.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Contacts.psm1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:Root 'Core\HU.Office.psm1') -Force -DisableNameChecking
}

Describe 'vCard lesen' {
    It 'vCard 4.0 mit Faltung, Maskierung und Adresse' {
        $t = "BEGIN:VCARD`r`nVERSION:4.0`r`nFN:Musikverein Beispiel`r`nORG:Musikverein Beispiel\, Ortsgruppe;Kassa`r`nEMAIL;TYPE=work;PREF=1:kassa@`r`n beispiel.invalid`r`nTEL;TYPE=cell:+43 660 0000000`r`nTEL;TYPE=work,voice:+43 3800 0000`r`nADR;TYPE=work:;;Hauptstraße 1;Musterort;Stmk.;8000;Österreich`r`nEND:VCARD`r`n"
        $k = @(ConvertFrom-HUVCard -Text $t)
        $k.Count | Should -Be 1
        $k[0].Name | Should -Be 'Musikverein Beispiel'
        $k[0].Firma | Should -Be 'Musikverein Beispiel, Ortsgruppe'
        $k[0].Mail | Should -Be 'kassa@beispiel.invalid'
        $k[0].Mobil | Should -Be '+43 660 0000000'
        $k[0].Telefon | Should -Be '+43 3800 0000'
        $k[0].Strasse | Should -Be 'Hauptstraße 1'
        $k[0].Plz | Should -Be '8000'
        $k[0].Ort | Should -Be 'Musterort'
        $k[0].Land | Should -Be 'Österreich'
    }
    It 'vCard 2.1 wie aus Outlook: Quoted-Printable mit Zeichensatz und weichem Umbruch' {
        $t = "BEGIN:VCARD`r`nVERSION:2.1`r`nN;CHARSET=Windows-1252;ENCODING=QUOTED-PRINTABLE:M=FCller;J=FCrgen`r`nFN;CHARSET=utf-8;ENCODING=QUOTED-PRINTABLE:J=C3=BCrgen M=C3=BCller`r`nTEL;WORK;VOICE:03800 1111`r`nTEL;CELL;VOICE:0660 2222`r`nADR;WORK;ENCODING=QUOTED-PRINTABLE:;;Gro=DFe Gasse 2=`r`n;Musterort;;8000;`r`nEMAIL;PREF;INTERNET:juergen@beispiel.invalid`r`nEND:VCARD`r`n"
        $k = @(ConvertFrom-HUVCard -Text $t)
        $k[0].Name | Should -Be 'Jürgen Müller'
        $k[0].Telefon | Should -Be '03800 1111'
        $k[0].Mobil | Should -Be '0660 2222'
        $k[0].Strasse | Should -Be 'Große Gasse 2'
        $k[0].Ort | Should -Be 'Musterort'
        $k[0].Mail | Should -Be 'juergen@beispiel.invalid'
    }
    It 'ohne FN: Name aus N, gruppierte Eigenschaften (item1.EMAIL), mehrere Karten' {
        $t = "BEGIN:VCARD`nVERSION:3.0`nN:Huber;Anna;;;`nitem1.EMAIL;type=INTERNET:anna@beispiel.invalid`nEND:VCARD`nBEGIN:VCARD`nVERSION:3.0`nFN:Sportverein`nEND:VCARD`n"
        $k = @(ConvertFrom-HUVCard -Text $t)
        $k.Count | Should -Be 2
        $k[0].Name | Should -Be 'Anna Huber'
        $k[0].Mail | Should -Be 'anna@beispiel.invalid'
        $k[1].Name | Should -Be 'Sportverein'
    }
    It 'Datei in Windows-1252 und in UTF-8 mit BOM' {
        $dir = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtVcf-' + [guid]::NewGuid().ToString('N').Substring(0, 8)); [void](New-Item -ItemType Directory -Path $dir)
        try {
            $txt = "BEGIN:VCARD`r`nVERSION:3.0`r`nFN:Bäckerei Löffler`r`nEND:VCARD`r`n"
            [IO.File]::WriteAllBytes((Join-Path $dir 'a.vcf'), [Text.Encoding]::GetEncoding(1252).GetBytes($txt))
            [IO.File]::WriteAllText((Join-Path $dir 'b.vcf'), $txt, (New-Object Text.UTF8Encoding $true))
            (Import-HUVCardFile (Join-Path $dir 'a.vcf'))[0].Name | Should -Be 'Bäckerei Löffler'
            (Import-HUVCardFile (Join-Path $dir 'b.vcf'))[0].Name | Should -Be 'Bäckerei Löffler'
            (Import-HUVCardFile (Join-Path $dir 'b.vcf'))[0].Quelle | Should -Be 'vCard b.vcf'
        } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    It 'leere oder kaputte Datei liefert keine Kontakte und wirft nicht' {
        @(ConvertFrom-HUVCard -Text '').Count | Should -Be 0
        @(ConvertFrom-HUVCard -Text "BEGIN:VCARD`nVERSION:3.0`nEND:VCARD").Count | Should -Be 0
        @(ConvertFrom-HUVCard -Text "irgendwas`nohne:karte").Count | Should -Be 0
    }
    It 'Suche nach Name, Firma, Mail und Ort' {
        $k = @((New-HUContact -Name 'Anna Huber' -Ort 'Musterort'), (New-HUContact -Firma 'Sportverein' -Mail 'sv@x.invalid'))
        @(Find-HUContact $k 'sport').Count | Should -Be 1
        @(Find-HUContact $k 'muster').Count | Should -Be 1
        @(Find-HUContact $k '').Count | Should -Be 2
        (New-HUContact -Firma 'Nur Firma').Name | Should -Be 'Nur Firma'
    }
}

Describe 'Outlook-Kontakte (Stellvertreter)' {
    It 'nimmt nur Kontakte (Class 40), Adresse und Telefon mit Rueckfall, filtert nach Suchtext' {
        $mk = { param($h) $o = [pscustomobject]$h; foreach ($p in 'FullName', 'CompanyName', 'Email1Address', 'BusinessTelephoneNumber', 'HomeTelephoneNumber', 'MobileTelephoneNumber', 'MailingAddressStreet', 'MailingAddressPostalCode', 'MailingAddressCity', 'MailingAddressCountry', 'BusinessAddressStreet', 'BusinessAddressPostalCode', 'BusinessAddressCity', 'BusinessAddressCountry', 'HomeAddressStreet', 'HomeAddressPostalCode', 'HomeAddressCity', 'HomeAddressCountry') { if (-not $o.PSObject.Properties[$p]) { $o | Add-Member NoteProperty $p '' } }; $o }
        $items = @(
            (& $mk @{ Class = 40; FullName = 'Anna Huber'; Email1Address = 'anna@x.invalid'; HomeTelephoneNumber = '0123'; HomeAddressStreet = 'Weg 1'; HomeAddressCity = 'Musterort'; HomeAddressPostalCode = '8000' })
            (& $mk @{ Class = 69; FullName = 'Verteiler' })
            (& $mk @{ Class = 40; FullName = ''; CompanyName = 'Sportverein'; BusinessTelephoneNumber = '0456'; MailingAddressStreet = 'Platz 2'; MailingAddressCity = 'Anderort' })
        )
        $folder = [pscustomobject]@{ Items = $items }
        $ns = [pscustomobject]@{ F = $folder }; $ns | Add-Member ScriptMethod GetDefaultFolder { param($n) if ($n -eq 10) { $this.F } }
        $sess = [pscustomobject]@{ Namespace = $ns }
        $k = @(Get-HUOutlookContacts -Session $sess)
        $k.Count | Should -Be 2
        $k[0].Name | Should -Be 'Anna Huber'; $k[0].Telefon | Should -Be '0123'; $k[0].Strasse | Should -Be 'Weg 1'; $k[0].Quelle | Should -Be 'Outlook'
        $k[1].Name | Should -Be 'Sportverein'; $k[1].Ort | Should -Be 'Anderort'
        @(Get-HUOutlookContacts -Session $sess -Search 'sport').Count | Should -Be 1
    }
}

Describe 'Fenster Kontakt uebernehmen' {
    It 'XAML ist gueltig' {
        if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { Set-ItResult -Skipped -Because 'XAML braucht STA'; return }
        Add-Type -AssemblyName PresentationFramework
        $src = [IO.File]::ReadAllText((Join-Path $script:Root 'Functions\UI-Contacts.ps1'), [Text.Encoding]::UTF8)
        $x = [regex]::Match($src, "(?s)@'\r?\n(<Window.*?</Window>)\r?\n'@").Groups[1].Value
        $theme = [IO.File]::ReadAllText((Join-Path $script:Root 'XAML\Theme.xaml'), [Text.Encoding]::UTF8)
        $x = $x.Replace('<!--HU:THEME-->', [regex]::Match($theme, '(?s)<ResourceDictionary[^>]*>(.*)</ResourceDictionary>').Groups[1].Value)
        $w = [Windows.Markup.XamlReader]::Parse($x)
        $w.FindName('grd') | Should -Not -BeNullOrEmpty
        $w.FindName('btnVcf') | Should -Not -BeNullOrEmpty
    }
}

Describe 'vCard: Base64 mit = am Zeilenende (Pruefung 086, KERN-B2)' {
    It 'PHOTO <Photo> laesst END:VCARD und die naechste Karte stehen' -TestCases @(@{ Photo = 'abcde=' }, @{ Photo = 'abcde==' }) {
        param($Photo)
        foreach ($v in '3.0', '4.0') {
            $t = "BEGIN:VCARD`r`nVERSION:$v`r`nFN:Erste Person`r`nPHOTO;ENCODING=b;TYPE=JPEG:$Photo`r`nEND:VCARD`r`nBEGIN:VCARD`r`nVERSION:$v`r`nFN:Zweite Person`r`nEMAIL:zwei@beispiel.invalid`r`nEND:VCARD`r`n"
            $k = @(ConvertFrom-HUVCard -Text $t)
            $k.Count | Should -Be 2
            $k[0].Name | Should -Be 'Erste Person'
            $k[1].Mail | Should -Be 'zwei@beispiel.invalid'
        }
    }
    It 'gefaltetes Base64 mit = vor der Faltung bleibt eine Eigenschaft' {
        $t = "BEGIN:VCARD`r`nVERSION:3.0`r`nFN:Gefaltet`r`nPHOTO;ENCODING=b:abc=`r`n def==`r`nEMAIL:f@beispiel.invalid`r`nEND:VCARD`r`n"
        $k = @(ConvertFrom-HUVCard -Text $t)
        $k.Count | Should -Be 1
        $k[0].Mail | Should -Be 'f@beispiel.invalid'
    }
    It 'Quoted-Printable mit weichem Umbruch wird weiter zusammengefuegt' {
        $t = "BEGIN:VCARD`r`nVERSION:2.1`r`nN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:M=C3=BCller;J=`r`n=C3=BCrgen`r`nNOTE;QUOTED-PRINTABLE:a=`r`nb`r`nEND:VCARD`r`n"
        $k = @(ConvertFrom-HUVCard -Text $t)
        $k.Count | Should -Be 1
        $k[0].Name | Should -Be 'Jürgen Müller'
    }
}

Describe 'Suche mit Sonderzeichen (Pruefung 086, KERN-B7)' {
    It 'eckige Klammer und Stern sind kein Platzhalter' {
        $c = @((New-HUContact -Name 'Verein [Kassa]' -Mail 'k@beispiel.invalid'), (New-HUContact -Name 'Anderer Verein' -Mail 'a@beispiel.invalid'))
        @(Find-HUContact -Contacts $c -Search '[').Count | Should -Be 1
        @(Find-HUContact -Contacts $c -Search '[kassa]')[0].Name | Should -Be 'Verein [Kassa]'
        @(Find-HUContact -Contacts $c -Search '*').Count | Should -Be 0
        @(Find-HUContact -Contacts $c -Search 'VEREIN').Count | Should -Be 2
    }
    It 'Outlook-Kontaktsuche nutzt IndexOf statt -like' {
        $src = [IO.File]::ReadAllText((Join-Path $script:Root 'Core\HU.Office.psm1'), [Text.Encoding]::UTF8)
        $src | Should -Not -Match '-notlike "\*\$q\*"'
        $src | Should -Match '\$city"\.IndexOf\(\$q, \[StringComparison\]::OrdinalIgnoreCase\)'
    }
}
