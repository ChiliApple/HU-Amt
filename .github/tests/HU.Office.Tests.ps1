#Requires -Version 5.1
# Pester 5: Outlook und Word (Core\HU.Office.psm1). Ohne echtes Outlook oder Word: COM wird durch Stellvertreter ersetzt,
# es oeffnet sich kein Fenster und kein Office-Prozess startet.

BeforeAll {
    Set-StrictMode -Version 1.0   # wie Main.ps1: nicht gesetzte Variablen sind Fehler
    $script:Root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    Import-Module (Join-Path $script:Root 'Core\HU.Office.psm1') -Force -DisableNameChecking

    # Stellvertreter fuer ein Outlook-Element mit Benutzereigenschaft
    function New-FakeItem([string]$Id, [string]$Key) {
        $props = [pscustomobject]@{ Map = @{} }
        if ($Key) { $props.Map['HUAmtSchluessel'] = [pscustomobject]@{ Value = $Key } }
        $props | Add-Member ScriptMethod Find { param($n) $this.Map[$n] }
        $props | Add-Member ScriptMethod Add { param($n, $t, $f) $p = [pscustomobject]@{ Value = '' }; $this.Map[$n] = $p; $p }
        return [pscustomobject]@{ EntryID = $Id; UserProperties = $props }
    }
}

Describe 'Schluessel' {
    It 'erlaubt Modul-Kennung mit Nummer' {
        Test-HUOfficeKey 'Fristen:F-2026-0012' | Should -BeTrue
        Test-HUOfficeKey 'Buchung:2026.17_a' | Should -BeTrue
    }
    It 'lehnt Zeichen ab, die einen Filter brechen koennten' {
        Test-HUOfficeKey "a'b" | Should -BeFalse
        Test-HUOfficeKey 'a b' | Should -BeFalse
        Test-HUOfficeKey '' | Should -BeFalse
        Test-HUOfficeKey ('a' * 121) | Should -BeFalse
        { Find-HUOutlookItem -Session ([pscustomobject]@{}) -FolderObject ([pscustomobject]@{}) -Key "x' or 1=1" } | Should -Throw
    }
}

Describe 'Ordnerangabe' {
    It 'leer ergibt den Standardordner des Typs' {
        $s = ConvertTo-HUOutlookFolderSpec $null -Type Tasks
        $s.Type | Should -Be 'Tasks'
        $s.EntryId | Should -Be ''
    }
    It 'uebernimmt gespeicherte Werte' {
        $s = ConvertTo-HUOutlookFolderSpec ([pscustomobject]@{ Name = 'Saal'; EntryId = 'E1'; StoreId = 'S1'; Type = 'Calendar' })
        $s.Name | Should -Be 'Saal'; $s.EntryId | Should -Be 'E1'; $s.StoreId | Should -Be 'S1'; $s.Type | Should -Be 'Calendar'
    }
    It 'Standardordner ueber GetDefaultFolder (Kalender 9, Aufgaben 13), sonst GetFolderFromID' {
        $ns = [pscustomobject]@{}
        $ns | Add-Member ScriptMethod GetDefaultFolder { param($n) "default-$n" }
        $ns | Add-Member ScriptMethod GetFolderFromID { param($e, $s) "id-$e-$s" }
        $sess = [pscustomobject]@{ Namespace = $ns }
        Get-HUOutlookFolder -Session $sess -Type Calendar | Should -Be 'default-9'
        Get-HUOutlookFolder -Session $sess -Type Tasks | Should -Be 'default-13'
        Get-HUOutlookFolder -Session $sess -Folder ([pscustomobject]@{ EntryId = 'E'; StoreId = 'S' }) | Should -Be 'id-E-S'
    }
}

Describe 'Eintrag finden' {
    BeforeEach {
        $script:A = New-FakeItem 'A' 'Fristen:1'
        $script:B = New-FakeItem 'B' 'Fristen:2'
        $script:C = New-FakeItem 'C' ''
        $all = @($script:A, $script:B, $script:C)
        $script:Ns = [pscustomobject]@{ Items = $all }
        $script:Ns | Add-Member ScriptMethod GetItemFromID { param($e, $s) $this.Items | Where-Object { $_.EntryID -eq $e } | Select-Object -First 1 }
        $script:Sess = [pscustomobject]@{ Namespace = $script:Ns }
        $script:RestrictWorks = $true
        $items = New-Object System.Collections.ArrayList
        [void]$items.AddRange($all)
        $items | Add-Member ScriptMethod Restrict {
            param($f)
            if (-not $script:RestrictWorks) { throw 'Feld nicht im Ordner' }
            $k = [regex]::Match($f, "'(.*)'").Groups[1].Value
            , @($this | Where-Object { (Get-HUOutlookItemKey $_) -eq $k })
        }
        $script:Folder = [pscustomobject]@{ StoreID = 'S'; Items = $items }
    }
    It 'findet ueber die EntryID, wenn der Schluessel passt' {
        (Find-HUOutlookItem -Session $script:Sess -FolderObject $script:Folder -Key 'Fristen:2' -EntryId 'B').EntryID | Should -Be 'B'
    }
    It 'EntryID mit fremdem Schluessel zaehlt nicht, dann Suche ueber den Schluessel' {
        (Find-HUOutlookItem -Session $script:Sess -FolderObject $script:Folder -Key 'Fristen:2' -EntryId 'A').EntryID | Should -Be 'B'
    }
    It 'findet ueber Restrict' {
        (Find-HUOutlookItem -Session $script:Sess -FolderObject $script:Folder -Key 'Fristen:1').EntryID | Should -Be 'A'
    }
    It 'wirft Restrict (Feld fehlt im Ordner), wird durchsucht' {
        $script:RestrictWorks = $false
        (Find-HUOutlookItem -Session $script:Sess -FolderObject $script:Folder -Key 'Fristen:2').EntryID | Should -Be 'B'
    }
    It 'unbekannter Schluessel: $null' {
        Find-HUOutlookItem -Session $script:Sess -FolderObject $script:Folder -Key 'Fristen:9' | Should -BeNullOrEmpty
    }
}

Describe 'Termin und Aufgabe setzen' {
    BeforeEach {
        # Ordner mit leerer Items-Auflistung (eigenes Objekt: bei einer ArrayList gewinnt deren .NET-Add gegen eine Skriptmethode)
        $script:Saved = 0
        $items = [pscustomobject]@{ List = (New-Object System.Collections.ArrayList) }
        $items | Add-Member ScriptMethod Restrict {
            param($f)
            $k = [regex]::Match($f, "'(.*)'").Groups[1].Value
            , @($this.List | Where-Object { (Get-HUOutlookItemKey $_) -eq $k })
        }
        $items | Add-Member ScriptMethod Add {
            param($t)
            $it = New-FakeItem ('N' + $this.List.Count) ''
            foreach ($p in 'Subject', 'AllDayEvent', 'Start', 'End', 'Location', 'Body', 'BusyStatus', 'Categories', 'ReminderSet', 'ReminderMinutesBeforeStart', 'StartDate', 'DueDate', 'ReminderTime') { $it | Add-Member NoteProperty $p $null }
            $it | Add-Member NoteProperty ItemType $t
            $it | Add-Member NoteProperty Parent ([pscustomobject]@{ StoreID = 'S' })
            $it | Add-Member ScriptMethod Save { $script:Saved++ }
            [void]$this.List.Add($it); $it
        }
        $script:Folder = [pscustomobject]@{ StoreID = 'S'; Items = $items }
        $ns = [pscustomobject]@{ Folder = $script:Folder }
        $ns | Add-Member ScriptMethod GetDefaultFolder { param($n) $this.Folder }
        $ns | Add-Member ScriptMethod GetItemFromID { param($e, $s) $null }
        $script:Sess = [pscustomobject]@{ Namespace = $ns }
    }
    It 'Termin: legt an, setzt Zeiten, Erinnerung und Schluessel' {
        $r = Set-HUOutlookAppointment -Session $script:Sess -Key 'Buchung:1' -Subject 'Saal' -Start ([datetime]'2026-12-04 18:00') -End ([datetime]'2026-12-04 23:00') -ReminderMinutes 60 -Body 'Text' -Location 'Ort' -Categories 'Kat'
        $r.Created | Should -BeTrue
        $it = $script:Folder.Items.List[0]
        $it.ItemType | Should -Be 1
        $it.End | Should -Be ([datetime]'2026-12-04 23:00')
        $it.ReminderSet | Should -BeTrue; $it.ReminderMinutesBeforeStart | Should -Be 60
        $it.BusyStatus | Should -Be 2
        $it.Body | Should -Be 'Text'; $it.Location | Should -Be 'Ort'; $it.Categories | Should -Be 'Kat'
        Get-HUOutlookItemKey $it | Should -Be 'Buchung:1'
        $script:Saved | Should -Be 1
    }
    It 'Termin: zweiter Aufruf aendert denselben Eintrag' {
        [void](Set-HUOutlookAppointment -Session $script:Sess -Key 'Buchung:1' -Subject 'Saal' -Start ([datetime]'2026-12-04 18:00'))
        $r = Set-HUOutlookAppointment -Session $script:Sess -Key 'Buchung:1' -Subject 'Saal neu' -Start ([datetime]'2026-12-05 18:00')
        $r.Created | Should -BeFalse
        $script:Folder.Items.List.Count | Should -Be 1
        $script:Folder.Items.List[0].Subject | Should -Be 'Saal neu'
        $script:Folder.Items.List[0].End | Should -Be ([datetime]'2026-12-05 19:00')
    }
    It 'Termin ganztaegig: Ende ist der Tag nach dem letzten Tag' {
        [void](Set-HUOutlookAppointment -Session $script:Sess -Key 'Buchung:2' -Subject 'Saal' -Start ([datetime]'2026-12-02') -End ([datetime]'2026-12-08') -AllDay)
        $it = $script:Folder.Items.List[0]
        $it.AllDayEvent | Should -BeTrue
        $it.End | Should -Be ([datetime]'2026-12-09')
        $it.ReminderSet | Should -BeFalse
    }
    It 'Aufgabe: Faelligkeit, Beginn und Erinnerung' {
        [void](Set-HUOutlookTask -Session $script:Sess -Key 'Fristen:7' -Subject 'Spielplatz' -DueDate ([datetime]'2027-03-31 15:00') -StartDate ([datetime]'2027-03-01') -ReminderTime ([datetime]'2027-03-24 08:00') -Body 'Notiz')
        $it = $script:Folder.Items.List[0]
        $it.ItemType | Should -Be 3
        $it.DueDate | Should -Be ([datetime]'2027-03-31')
        $it.StartDate | Should -Be ([datetime]'2027-03-01')
        $it.ReminderTime | Should -Be ([datetime]'2027-03-24 08:00')
        $it.Body | Should -Be 'Notiz'
    }
}

Describe 'Outlook starten mit Wiederholung' {
    BeforeAll {
        Mock -ModuleName HU.Office Test-HUOutlookClassic { $true }
        InModuleScope HU.Office { $script:HUOfRetryDelayMs = 10 }
        function New-RpcError([int]$HResult) { $e = New-Object System.Runtime.InteropServices.COMException('abgelehnt', $HResult); return $e }
    }
    AfterAll { InModuleScope HU.Office { $script:HUOfRetryDelayMs = 2000; $script:HUOfCreateApp = { New-Object -ComObject Outlook.Application -ErrorAction Stop } } }
    BeforeEach {
        $script:Calls = 0
        $script:FakeApp = [pscustomobject]@{}
        $script:FakeApp | Add-Member ScriptMethod GetNamespace { param($n) "ns-$n" }
    }
    It 'RPC_E_CALL_REJECTED beim kalten Start: wiederholt und oeffnet dann' {
        $err = New-RpcError -2147418111
        InModuleScope HU.Office -Parameters @{ Err = $err; App = $script:FakeApp } {
            param($Err, $App)
            $script:N = 0
            $script:TErr = $Err; $script:TApp = $App
            $script:HUOfCreateApp = { $script:N++; if ($script:N -lt 3) { throw $script:TErr }; $script:TApp }
        }
        $s = Open-HUOutlook -TimeoutSeconds 5
        $s.Namespace | Should -Be 'ns-MAPI'
        InModuleScope HU.Office { $script:N } | Should -Be 3
    }
    It 'RPC-Server weg (0x800706BA) zaehlt ebenfalls als kurz nicht erreichbar' {
        $err = New-RpcError -2147023174
        InModuleScope HU.Office -Parameters @{ Err = $err; App = $script:FakeApp } {
            param($Err, $App)
            $script:N = 0
            $script:TErr = $Err; $script:TApp = $App
            $script:HUOfCreateApp = { $script:N++; if ($script:N -lt 2) { throw $script:TErr }; $script:TApp }
        }
        (Open-HUOutlook -TimeoutSeconds 5).Namespace | Should -Be 'ns-MAPI'
    }
    It 'andere Fehler werden nicht wiederholt' {
        InModuleScope HU.Office { $script:N = 0; $script:HUOfCreateApp = { $script:N++; throw 'Klasse nicht registriert' } }
        { Open-HUOutlook -TimeoutSeconds 5 } | Should -Throw '*nicht registriert*'
        InModuleScope HU.Office { $script:N } | Should -Be 1
    }
    It 'nach Ablauf der Zeit wird der letzte Fehler gemeldet' {
        $err = New-RpcError -2147418111
        InModuleScope HU.Office -Parameters @{ Err = $err } { param($Err); $script:TErr = $Err; $script:HUOfCreateApp = { throw $script:TErr } }
        { Open-HUOutlook -TimeoutSeconds 1 } | Should -Throw '*abgelehnt*'
    }
}

Describe 'Outlook schliessen' {
    BeforeEach {
        $script:Quit = 0
        $script:App = [pscustomobject]@{ Explorers = [pscustomobject]@{ Count = 0 }; Inspectors = [pscustomobject]@{ Count = 0 } }
        $script:App | Add-Member ScriptMethod Quit { $script:Quit++ }
    }
    It 'lief Outlook vorher nicht und ist kein Fenster offen: beenden' {
        Close-HUOutlook ([pscustomobject]@{ Application = $script:App; Namespace = $null; WasRunning = $false })
        $script:Quit | Should -Be 1
    }
    It 'lief Outlook schon: nicht beenden' {
        Close-HUOutlook ([pscustomobject]@{ Application = $script:App; Namespace = $null; WasRunning = $true })
        $script:Quit | Should -Be 0
    }
    It 'ist eine angezeigte Mail offen: nicht beenden' {
        $script:App.Inspectors.Count = 1
        Close-HUOutlook ([pscustomobject]@{ Application = $script:App; Namespace = $null; WasRunning = $false })
        $script:Quit | Should -Be 0
    }
}

Describe 'Mail' {
    It 'mailto mit kodiertem Betreff und Text' {
        $u = New-HUMailtoUri -To 'a@b.at' -Subject 'Bestätigung 1' -Body "Zeile 1`r`nZeile 2"
        $u | Should -Match '^mailto:a@b\.at\?subject=Best%C3%A4tigung%201&body=Zeile%201%0D%0AZeile%202$'
    }
    It 'zu langer Text wird gekuerzt und bekommt den Hinweis' {
        $u = New-HUMailtoUri -Subject 'X' -Body ('a' * 5000) -MaxLength 300 -Note 'gekuerzt'
        $u.Length | Should -BeLessOrEqual 300
        $u | Should -Match 'gekuerzt$'
    }
}

Describe '.ics' {
    It 'hat CRLF, DTSTAMP, UID und Zeiten in UTC' {
        $c = New-HUIcsContent -Uid 'Buchung:7' -Subject 'Saal' -Start ([datetime]'2026-12-04 18:00') -End ([datetime]'2026-12-04 23:00') -Now ([datetime]'2026-11-07 10:00')
        $c | Should -Match "^BEGIN:VCALENDAR`r`n"
        $c | Should -Match "`r`nEND:VCALENDAR`r`n$"
        ($c -replace "`r`n", '') | Should -Not -Match "`n"
        $c | Should -Match "`r`nUID:Buchung:7`r`n"
        $c | Should -Match "`r`nDTSTAMP:\d{8}T\d{6}Z`r`n"
        $c | Should -Match ("`r`nDTSTART:" + ([datetime]'2026-12-04 18:00').ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'") + "`r`n")
    }
    It 'ganztaegig: Datum, Ende exklusiv' {
        $c = New-HUIcsContent -Uid 'F:1' -Subject 'Frist' -Start ([datetime]'2026-12-31') -AllDay
        $c | Should -Match "`r`nDTSTART;VALUE=DATE:20261231`r`n"
        $c | Should -Match "`r`nDTEND;VALUE=DATE:20270101`r`n"
    }
    It 'Erinnerung als VALARM mit TRIGGER' {
        $c = New-HUIcsContent -Uid 'F:1' -Subject 'Frist' -Start ([datetime]'2026-12-31 08:00') -AlarmMinutes 15
        $c | Should -Match "`r`nBEGIN:VALARM`r`nACTION:DISPLAY`r`n"
        $c | Should -Match "`r`nTRIGGER:-PT15M`r`n"
    }
    It 'maskiert Komma, Semikolon, Backslash und Zeilenwechsel' {
        ConvertTo-HUIcsText "a,b;c\d`r`ne" | Should -Be 'a\,b\;c\\d\ne'
    }
    It 'faltet lange Zeilen bei 75 Oktetten, ohne Umlaute zu zerteilen' {
        $line = 'DESCRIPTION:' + ('ä' * 100)
        $f = Format-HUIcsLine $line
        $parts = $f -split "`r`n"
        $enc = New-Object System.Text.UTF8Encoding $false
        foreach ($p in $parts) { $enc.GetByteCount($p) | Should -BeLessOrEqual 75 }
        ($parts[1..($parts.Count - 1)] | ForEach-Object { $_.StartsWith(' ') }) | Should -Not -Contain $false
        (($parts | ForEach-Object { $_ -replace '^ ', '' }) -join '') | Should -Be $line
    }
    It 'lehnt ein Ende vor dem Beginn ab' {
        { New-HUIcsContent -Uid 'x' -Subject 's' -Start ([datetime]'2026-12-04 18:00') -End ([datetime]'2026-12-04 17:00') } | Should -Throw
    }
    It 'schreibt die Datei ohne BOM' {
        $p = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtIcs-' + [guid]::NewGuid().ToString('N') + '.ics')
        try {
            New-HUIcsFile -Path $p -Uid 'x' -Subject 'Test' -Start ([datetime]'2026-12-04 18:00') | Should -Be $p
            $b = [IO.File]::ReadAllBytes($p)
            $b[0] | Should -Be ([byte][char]'B')
        } finally { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Word' {
    It 'Zeilenwechsel werden manuelle Zeilenumbrueche, Steuerzeichen fallen weg' {
        ConvertTo-HUWordText "a`r`nb`nc$([char]7)" | Should -Be ("a$([char]11)b$([char]11)c")
    }
    It 'fehlende Vorlage oder fehlendes Ziel wird gemeldet, bevor Word startet' {
        { Invoke-HUWordTemplate -Template 'C:\gibt-es-nicht\x.docx' -PdfPath 'x.pdf' } | Should -Throw '*Vorlage nicht gefunden*'
        $t = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtTpl-' + [guid]::NewGuid().ToString('N') + '.docx')
        try {
            [IO.File]::WriteAllText($t, 'x')
            { Invoke-HUWordTemplate -Template $t } | Should -Throw '*Kein Ziel*'
        } finally { Remove-Item -LiteralPath $t -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Aufpasser fuer eigene Prozesse' {
    BeforeEach { $script:P = Start-Process -FilePath "$env:SystemRoot\System32\PING.EXE" -ArgumentList '-n', '60', '127.0.0.1' -WindowStyle Hidden -PassThru }
    AfterEach { Stop-Process -Id $script:P.Id -Force -ErrorAction SilentlyContinue }
    It 'beendet den Prozess nach Ablauf der Zeit' {
        $g = InModuleScope HU.Office -Parameters @{ Id = $script:P.Id } { param($Id) Start-HUProcessGuard -ProcessId $Id -Seconds 1 }
        $script:P.WaitForExit(8000) | Should -BeTrue
        $g.Fired() | Should -BeTrue
        $g.Stop()
    }
    It 'Stop vor Ablauf laesst den Prozess in Ruhe' {
        $g = InModuleScope HU.Office -Parameters @{ Id = $script:P.Id } { param($Id) Start-HUProcessGuard -ProcessId $Id -Seconds 2 }
        $g.Stop()
        $script:P.WaitForExit(3500) | Should -BeFalse
        $g.Fired() | Should -BeFalse
    }
    It 'Nachlauf beendet einen Prozess, der nach der Frist noch laeuft' {
        InModuleScope HU.Office -Parameters @{ Id = $script:P.Id } { param($Id) Stop-HUOwnProcess -ProcessId $Id -GraceSeconds 1 }
        $script:P.WaitForExit(8000) | Should -BeTrue
    }
}

Describe 'Testmodus sperrt Outlook, Word und Mail' {
    BeforeAll { $script:OldEnv = $env:HUAMT_TESTMODE; $env:HUAMT_TESTMODE = '1' }
    AfterAll { if ($null -eq $script:OldEnv) { Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue } else { $env:HUAMT_TESTMODE = $script:OldEnv } }
    It 'Test-HUOfficeTestMode erkennt den Testmodus' { Test-HUOfficeTestMode | Should -BeTrue }
    It 'Outlook wird nicht geoeffnet' { { Open-HUOutlook } | Should -Throw '*Testmodus*' }
    It 'Termin ohne Sitzung wird nicht angelegt' { { Set-HUOutlookAppointment -Key 'Test:1' -Subject 's' -Start (Get-Date) } | Should -Throw '*Testmodus*' }
    It 'Word wird nicht gestartet' { { Invoke-HUWordTemplate -Template 'C:\x.docx' -PdfPath 'x.pdf' } | Should -Throw '*Testmodus*' }
    It 'Mail oeffnet nichts und meldet den Testmodus' { (Show-HUMail -To 'a@b.at' -Subject 'x').Way | Should -Be 'testmodus' }
    It '.ics-Datei geht weiter (nur eine Datei)' {
        $p = Join-Path ([IO.Path]::GetTempPath()) ('HUAmtIcsT-' + [guid]::NewGuid().ToString('N') + '.ics')
        try { New-HUIcsFile -Path $p -Uid 'x' -Subject 'T' -Start ([datetime]'2030-01-01 10:00') | Should -Be $p } finally { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Erkennung' {
    It 'liefert Wahr/Falsch und wirft nie' {
        Test-HUOutlookClassic | Should -BeOfType [bool]
        Test-HUWord | Should -BeOfType [bool]
    }
}

Describe 'Word-Vorlage ohne Makros (Pruefung 086, KERN-B3)' {
    It 'lehnt <Ext> mit klarer Meldung ab, bevor Word startet' -TestCases @(@{ Ext = '.docm' }, @{ Ext = '.doc' }, @{ Ext = '.dotm' }, @{ Ext = '.rtf' }, @{ Ext = '' }) {
        param($Ext)
        # Vorlage muss nicht existieren: die Endung wird zuerst geprueft (auch vor Testmodus-Sperre und Word-Suche)
        { Invoke-HUWordTemplate -Template "C:\gibt-es-nicht\vorlage$Ext" -PdfPath 'C:\gibt-es-nicht\x.pdf' } | Should -Throw '*.docx- oder .dotx-Datei*'
    }
    It 'laesst .docx und .dotx durch (dann kommt die naechste Pruefung)' -TestCases @(@{ Ext = '.docx' }, @{ Ext = '.DOTX' }) {
        param($Ext)
        $env:HUAMT_TESTMODE = '1'
        try { { Invoke-HUWordTemplate -Template "C:\gibt-es-nicht\vorlage$Ext" -PdfPath 'C:\gibt-es-nicht\x.pdf' } | Should -Throw '*Testmodus*' }
        finally { Remove-Item -LiteralPath Env:\HUAMT_TESTMODE -ErrorAction SilentlyContinue }
    }
    It 'setzt AutomationSecurity = 3 vor Documents.Open' {
        $src = [IO.File]::ReadAllText((Join-Path $script:Root 'Core\HU.Office.psm1'), [Text.Encoding]::UTF8)
        $body = [regex]::Match($src, '(?s)function Invoke-HUWordTemplate \{.*?\n\}').Value
        $body | Should -Not -BeNullOrEmpty
        $iSec = $body.IndexOf('$word.AutomationSecurity = 3')
        $iOpen = $body.IndexOf('$docs.Open(')
        $iSec | Should -BeGreaterThan 0
        $iOpen | Should -BeGreaterThan $iSec
        $body.IndexOf('Assert-HUWordTemplateExtension') | Should -BeLessThan $body.IndexOf('Assert-HUOfficeLive')
    }
}

Describe 'Anderes Postfach (Get-HUOutlookMailboxInbox, ab 1.0.4)' {
    BeforeEach {
        # Profil mit einem zweiten Konto; freigegebenes Postfach nur ueber CreateRecipient
        $store = [pscustomobject]@{}; $store | Add-Member ScriptMethod GetDefaultFolder { param($n) "konto-inbox-$n" }
        $acc = [pscustomobject]@{ SmtpAddress = 'zweit@example.org'; DeliveryStore = $store }
        $script:Shared = @{}
        $ns = [pscustomobject]@{ Accounts = @($acc); Known = @('amt@example.org', 'ohnerecht@example.org') }
        $ns | Add-Member ScriptMethod CreateRecipient {
            param($a) $r = [pscustomobject]@{ Address = $a; Resolved = $false; Known = $this.Known }
            $r | Add-Member ScriptMethod Resolve { $this.Resolved = ($this.Known -contains $this.Address); $this.Resolved }
            $r
        }
        $ns | Add-Member ScriptMethod GetSharedDefaultFolder { param($r, $n) if ($r.Address -eq 'ohnerecht@example.org') { throw 'Zugriff verweigert' }; "geteilt-$($r.Address)-$n" }
        $script:Sess = [pscustomobject]@{ Namespace = $ns }
    }
    It 'Konto im Profil: Posteingang seines Postfachs' {
        Get-HUOutlookMailboxInbox -Session $script:Sess -Mailbox ' Zweit@Example.org ' | Should -Be 'konto-inbox-6'
    }
    It 'freigegebenes Postfach: GetSharedDefaultFolder mit Posteingang' {
        Get-HUOutlookMailboxInbox -Session $script:Sess -Mailbox 'amt@example.org' | Should -Be 'geteilt-amt@example.org-6'
    }
    It 'unbekannte Adresse, fehlendes Recht und ungueltige Eingabe: verstaendliche Meldung' {
        { Get-HUOutlookMailboxInbox -Session $script:Sess -Mailbox 'wer@example.org' } | Should -Throw '*nicht bekannt*'
        { Get-HUOutlookMailboxInbox -Session $script:Sess -Mailbox 'ohnerecht@example.org' } | Should -Throw '*kein Leserecht*'
        { Get-HUOutlookMailboxInbox -Session $script:Sess -Mailbox 'kein mail' } | Should -Throw '*keine gültige Mailadresse*'
    }
    It 'Get-HUOutlookAttachments liest mit -Mailbox den Posteingang dieses Postfachs' {
        $mail = [pscustomobject]@{ Class = 43; ReceivedTime = (Get-Date); Attachments = [pscustomobject]@{ Count = 0 }; Subject = 'Beleg' }
        $items = [pscustomobject]@{ List = @($mail); I = 0 }
        $items | Add-Member ScriptMethod Sort { param($p, $d) }
        $items | Add-Member ScriptMethod GetFirst { $this.I = 0; $this.List[0] }
        $items | Add-Member ScriptMethod GetNext { $this.I++; if ($this.I -lt $this.List.Count) { $this.List[$this.I] } else { $null } }
        $fld = [pscustomobject]@{ Items = $items; StoreID = 'S-AMT' }
        $script:Sess.Namespace | Add-Member ScriptMethod GetDefaultFolder { param($n) throw 'eigener Posteingang darf nicht gelesen werden' } -Force
        $script:Sess.Namespace | Add-Member ScriptMethod GetSharedDefaultFolder { param($r, $n) $script:SharedHit = "$($r.Address)"; $fld } -Force
        $script:SharedHit = ''
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ("hu-mb-" + [guid]::NewGuid().ToString('N'))
        try {
            $r = @(Get-HUOutlookAttachments -Session $script:Sess -Mailbox 'amt@example.org' -TargetDir $tmp -Days 5)
            $r.Count | Should -Be 0
            $script:SharedHit | Should -Be 'amt@example.org'
        } finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

Describe 'Hineinziehen aus Outlook (ab 1.0.4)' {
    BeforeAll {
        Add-Type -AssemblyName PresentationCore
        # FileGroupDescriptorW bauen: Anzahl, dann je Datei 592 Byte (Flags, Groesse ab 64/68, Name ab 72)
        function New-FakeDescriptor([object[]]$Files) {
            $b = New-Object byte[] (4 + 592 * $Files.Count)
            [BitConverter]::GetBytes([uint32]$Files.Count).CopyTo($b, 0)
            for ($i = 0; $i -lt $Files.Count; $i++) {
                $o = 4 + 592 * $i
                [BitConverter]::GetBytes([uint32]0x40).CopyTo($b, $o)
                [BitConverter]::GetBytes([uint32]0).CopyTo($b, $o + 64)
                [BitConverter]::GetBytes([uint32]$Files[$i].Size).CopyTo($b, $o + 68)
                [Text.Encoding]::Unicode.GetBytes($Files[$i].Name).CopyTo($b, $o + 72)
            }
            return $b
        }
        function New-FakeData([hashtable]$Formats) {
            $d = [pscustomobject]@{ F = $Formats }
            $d | Add-Member ScriptMethod GetDataPresent { param($n) $this.F.ContainsKey("$n") }
            $d | Add-Member ScriptMethod GetData { param($n) $this.F["$n"] }
            return $d
        }
    }
    It 'liest Namen und Groessen aus FileGroupDescriptorW, nur den Dateinamen' {
        $r = @(Read-HUFileGroupDescriptor (New-FakeDescriptor @(@{ Name = 'Scan 1.pdf'; Size = 5 }, @{ Name = '..\..\boese.jpg'; Size = 7 })))
        $r.Count | Should -Be 2
        $r[0].Name | Should -Be 'Scan 1.pdf'; $r[0].Size | Should -Be 5; $r[0].Index | Should -Be 0
        $r[1].Name | Should -Be 'boese.jpg'; $r[1].Index | Should -Be 1
        @(Read-HUFileGroupDescriptor ([byte[]](1, 0))).Count | Should -Be 0
    }
    It 'Explorer: liefert vorhandene Dateien, nichts wird kopiert' {
        $f = Join-Path ([IO.Path]::GetTempPath()) ("hu-drop-" + [guid]::NewGuid().ToString('N') + '.pdf')
        [IO.File]::WriteAllText($f, 'x')
        try {
            $d = New-FakeData @{ ([System.Windows.DataFormats]::FileDrop) = [string[]]@($f, 'C:\gibt\es\nicht.pdf') }
            @(Get-HUDropNames $d).Count | Should -Be 2
            $r = @(Save-HUDroppedFiles -Data $d -TargetDir (Join-Path ([IO.Path]::GetTempPath()) 'hu-drop-nie'))
            $r | Should -Be @($f)
            Test-Path (Join-Path ([IO.Path]::GetTempPath()) 'hu-drop-nie') | Should -BeFalse
        } finally { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
    }
    It 'Outlook: schreibt die Anhaenge in den Zielordner, eindeutig, gekuerzt auf die Groesse, Fehler je Datei' {
        $desc = New-FakeDescriptor @(@{ Name = 'Abrechnung.pdf'; Size = 3 }, @{ Name = 'Abrechnung.pdf'; Size = 2 }, @{ Name = 'kaputt.png'; Size = 1 })
        $d = New-FakeData @{ FileGroupDescriptorW = (New-Object IO.MemoryStream (, $desc)) }
        Mock -ModuleName HU.Office Get-HUDropContent {
            if ($Index -eq 2) { throw 'nicht lesbar' }
            return , ([byte[]](65, 66, 67, 68, 69))
        }
        $dir = Join-Path ([IO.Path]::GetTempPath()) ("hu-drop-" + [guid]::NewGuid().ToString('N'))
        $prob = New-Object 'System.Collections.Generic.List[string]'
        try {
            @(Get-HUDropNames $d) | Should -Be @('Abrechnung.pdf', 'Abrechnung.pdf', 'kaputt.png')
            Test-HUDropHasFiles $d | Should -BeTrue
            $r = @(Save-HUDroppedFiles -Data $d -TargetDir $dir -Problems $prob)
            $r.Count | Should -Be 2
            [IO.Path]::GetFileName($r[0]) | Should -Be 'Abrechnung.pdf'
            [IO.Path]::GetFileName($r[1]) | Should -Be 'Abrechnung_2.pdf'
            [IO.File]::ReadAllText($r[0]) | Should -Be 'ABC'
            [IO.File]::ReadAllText($r[1]) | Should -Be 'AB'
            $prob.Count | Should -Be 1
            $prob[0] | Should -BeLike 'kaputt.png:*nicht lesbar*'
        } finally { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    It 'ohne Dateien: nichts' {
        $d = New-FakeData @{ Text = 'hallo' }
        Test-HUDropHasFiles $d | Should -BeFalse
        @(Save-HUDroppedFiles -Data $d -TargetDir 'C:\nie').Count | Should -Be 0
    }
    It 'C#-Teil fuer FileContents laesst sich unter dieser PowerShell uebersetzen' {
        { & (Get-Module HU.Office) { Initialize-HUDropNative } } | Should -Not -Throw
        ('HUDrop.Contents' -as [type]) | Should -Not -BeNullOrEmpty
    }
    It 'ohne Mock: Format-ID ab 0xC000 wird umgerechnet, kein Int16-Fehler (Pruefung 092, Kern-B3)' {
        & (Get-Module HU.Office) { Initialize-HUDropNative }
        [System.Windows.DataFormats]::GetDataFormat('FileContents').Id | Should -BeGreaterThan 32767
        [HUDrop.Contents]::ToFormat(0xC0A5) | Should -Be ([int16]-16219)
        [HUDrop.Contents]::ToFormat(15) | Should -Be ([int16]15)
        # echtes WPF-Datenobjekt ohne FileContents: Windows meldet "Format fehlt", nicht der Cast
        $d = New-Object System.Windows.DataObject
        $err = $null
        try { [void](Get-HUDropContent -Data $d -Index 0) } catch { $err = "$($_.Exception.Message) $($_.Exception.InnerException)" }
        $err | Should -Not -BeNullOrEmpty
        $err | Should -Not -Match 'Int16'
    }
}

Describe 'Anderes Postfach: Grenzfaelle (Pruefung 092)' {
    It 'Resolve wirft (offline): verstaendliche Meldung, Adresse ohne Punkt in der Domain ist ungueltig' {
        $ns = [pscustomobject]@{ Accounts = @() }
        $ns | Add-Member ScriptMethod CreateRecipient { param($a) $r = [pscustomobject]@{ Resolved = $false }; $r | Add-Member ScriptMethod Resolve { throw 'COM 0x80040115' }; $r }
        $sess = [pscustomobject]@{ Namespace = $ns }
        { Get-HUOutlookMailboxInbox -Session $sess -Mailbox 'amt@example.org' } | Should -Throw '*lässt sich gerade nicht prüfen*'
        { Get-HUOutlookMailboxInbox -Session $sess -Mailbox 'amt@localhost' } | Should -Throw '*keine gültige Mailadresse*'
    }
}
