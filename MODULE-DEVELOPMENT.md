# Module entwickeln

Ein Modul ist ein Ordner mit dem Namen seiner Kennung und wird als eigener Reiter eingebaut. [Zurück zum README](README.md)

## Aufbau

```
<Kennung>\
  addon.json   Beschreibung des Moduls
  Tab.xaml     Oberfläche, Wurzel <TabItem>
  Addon.ps1    Funktionen des Moduls
  Hilfe.html   optional, erscheint mit F1 im Reiter
```

`addon.json`:

```json
{
  "id": "Beispiel",
  "title": "Beispiel",
  "version": "1.0.0",
  "minApp": "0.1.0",
  "order": 10,
  "tab": { "header": "Beispiel", "xaml": "Tab.xaml", "script": "Addon.ps1", "init": "Initialize-HUAddonBeispiel" },
  "drop": [".pdf", ".xlsx"]
}
```

- **Kennung:** Buchstaben und Ziffern, beginnt mit einem Buchstaben. Reserviert sind `Benutzer`, `Laufend`, `Docs`, `Postfach` und `Wiederherstellung`.
- **`minApp`:** Ältere Programmversionen laden das Modul nicht.
- **`drop`:** Endungen, die der Reiter beim Hineinziehen annimmt. Ohne `drop` nimmt er keine Dateien.
- **Farbschema:** In `Tab.xaml` übernimmt `<TabItem.Resources><!--HU:THEME--></TabItem.Resources>` das dunkle Schema des Programms.

Funktionen in `Addon.ps1` (Konvention statt Schnittstelle):

| Funktion | Wann |
|---|---|
| `Initialize-HUAddon<Id>` | Pflicht, nach dem Einbau des Reiters |
| `Stop-HUAddon<Id>` | nur wenn etwas im Hintergrund läuft, beim Beenden |
| `Invoke-HUAddonDrop<Id> -Paths` | nur mit `drop`, beim Hineinziehen ins Fenster |

## Herkunft

| Ordner | Bedeutung |
|---|---|
| `Modules\<Id>\` | mitgeliefert, kommt mit dem Programm-Update |
| `Addons\<Id>\` | aus einem eigenen Repository, nur signiert. Neue Fassungen warten in `Addons\.pending\<Id>\` und werden beim nächsten Start übernommen |
| `AddonsDev\<Id>\` | Entwicklung, ohne Signaturprüfung, ersetzt ein Modul gleicher Kennung |

- **Daten:** `Data\<Id>\`, gemeinsam für alle Benutzer und in der Sicherung enthalten (`Get-HUAddonDataDir`).
- **Zugang zu privaten Repositories:** Token je Windows-Benutzer, verschlüsselt mit DPAPI.
- **Protokolle:** eigene Dateien `*.log` in `Data\Benutzer\<Name>\Logs\`. Der Kern löscht beim Start alle `*.log` dort, die älter als `logging.retentionDays` sind (Standard 30 Tage).
- **Kennung ändern:** Die Kennung legt Ordner und Daten fest. Ändert sie sich, zuerst ein signiertes Release mit der neuen Kennung veröffentlichen und erst danach die Modulquelle neu hinterlegen. Sonst holt das Update die alte Fassung, und es gibt zwei Reiter.

## Releases

- Ein Release enthält `HU-Amt-files.sha256` (Prüfsummen aller Dateien, erstellt von den automatischen Tests auf GitHub) und `HU-Amt-files.sha256.p7s` (Signatur des Herausgebers).
- Kanal **Stabil** nimmt nur freigegebene Releases, Kanal **Test** auch Vorab-Releases. In beiden Fällen nur signierte.

## Datei-Austausch mit anderen Programmen

Module schreiben keine eigenen CSV-Dateien, sondern nutzen den Kern (`Core\HU.Exchange.psm1`):

- **`Export-HUTable` / `Import-HUTable` mit Profil:** Trennzeichen, Kodierung (`utf8bom`, `utf8`, `windows-1252`), Dezimalzeichen, Datumsformat und Spalten (`Name` in der Datei, `Source` im Programm, `Type` = `text`, `number` oder `date`).
  - Vorgaben: `;`, Dezimalkomma, `dd.MM.yyyy`, UTF-8 mit BOM. Unter PowerShell 5.1 und 7 entstehen dieselben Bytes.
  - Profile liegen als JSON in `Data\<Modul>\Profile\` (`Get-HUTableProfileFolder`, `Get-HUTableProfiles`). Ein neues Zielformat ist nur eine Datei.
- **Übergabeordner:** `Export-HUHandoverTable` schreibt in `Data\<Modul>\Ausgang\` oder in den eingestellten Ordner (`Set-HUHandoverFolder`, auch eine Freigabe).
  - Die Datei entsteht unter einem Temp-Namen und wird erst fertig umbenannt; der Name trägt einen Zeitstempel, nichts wird überschrieben.

Beispiel:

```powershell
$p = New-HUTableProfile -Columns @(@{ Name = 'Beleg'; Source = 'Nr' }, @{ Name = 'Betrag'; Type = 'number' }, @{ Name = 'Datum'; Type = 'date' })
Export-HUHandoverTable -Rows $rows -TableProfile $p -DataRoot $script:DataRoot -ModuleId 'Beispiel' -BaseName 'Liste'
```

## Dateien an andere Module weitergeben

Der Kern hat ein Postfach je Modul im gemeinsamen Datenordner (`Data\Postfach\<Modul>\`). Nur einbauen, wenn die Funktionen vorhanden sind (`Get-Command`).

- `Get-HUAddonTargets -Extension '.pdf' [-Exclude <Id>]`: Module, die diese Endung annehmen (`Id`, `Title`).
- `Send-HUAddonFiles -To <Id> -Paths <...> -From <Id> [-Open]`: gibt Dateien weiter, Rückgabe ist die Anzahl.
  - Ohne `-Open` kommt eine Kopie ins Postfach. Das klappt auch zwischen PCs.
  - Mit `-Open` wechselt die App zum Reiter und übergibt die Dateien sofort.
- `Get-HUAddonInbox <Id>`: wartende Einträge (`Key`, `Path`, `Name`, `From`, `Pc`, `User`, `Time`), nur mit den Endungen aus `drop`.
- `Remove-HUAddonInboxItem <Eintrag>`: Eintrag samt Kopie entfernen, erst nach gelungener Übernahme.
- `Update-HUAddonInboxBadge [<Id>]`: Zahl am Reiter neu zählen. Der Kern zählt zusätzlich jede Minute.

## Outlook und Word

Module bauen keine eigenen COM-Aufrufe, sondern nutzen den Kern (`Core\HU.Office.psm1`, ab Version 0.1.7, also `minApp` 0.1.7). Grundsätze:

- Nur klassisches Outlook (`Test-HUOutlookClassic`), Word über `Test-HUWord`. Ohne Outlook: `.ics` bzw. `mailto:`.
- Mails werden nur angezeigt, nie gesendet.
- COM braucht einen STA-Thread: den Fenster-Thread oder einen Hintergrund-Auftrag des Kerns. `Select-HUOutlookFolder` und `Show-HUMail` nur im Fenster-Thread.
- Jeder Termin und jede Aufgabe hat einen **Schlüssel** (`<Modul>:<Id>`, nur Buchstaben, Ziffern und `: . _ -`). Er steht in der Benutzereigenschaft `HUAmtSchluessel`. Das Modul speichert zusätzlich `EntryId` und `StoreId` aus dem Ergebnis, das macht das Wiederfinden schneller.

**Outlook**

- `Open-HUOutlook` / `Close-HUOutlook $s` oder `Invoke-HUOutlook { param($s) ... }`: eine Sitzung für viele Einträge, z. B. beim Abgleich am Start. Alle folgenden Funktionen nehmen `-Session`; ohne sie öffnen und schließen sie selbst.
  - Outlook wird nur beendet, wenn es vorher nicht lief und kein Fenster offen ist.
  - Lehnt Outlook beim kalten Start oder kurz nach dem Beenden Aufrufe ab (`RPC_E_CALL_REJECTED` u. a.), wiederholt `Open-HUOutlook` bis `-TimeoutSeconds` (Standard 30). Ab Kern 0.1.8.
- `Select-HUOutlookFolder -Type Calendar|Tasks`: Ordner im Outlook-Dialog wählen (ab 0.1.10 kommt der Dialog nach vorne), auch einen gemeinsamen Kalender. Rückgabe `Name`, `Path`, `EntryId`, `StoreId`, `Type` oder `$null` bei Abbruch. Diese Angabe speichert das Modul in seinen Einstellungen und gibt sie als `-Folder` weiter. Leer = Standardordner des Benutzers.
- `Set-HUOutlookAppointment -Folder -Key [-EntryId] -Subject -Start [-End] [-AllDay] [-Location] [-Body] [-ReminderMinutes] [-BusyStatus] [-Categories]`: anlegen oder denselben Termin ändern. Rückgabe `EntryId`, `StoreId`, `Created`.
- `Set-HUOutlookTask -Folder -Key [-EntryId] -Subject -DueDate [-StartDate] [-ReminderTime] [-Body] [-Categories]`: dasselbe für Aufgaben.
- `Complete-HUOutlookItem -Type Task|Appointment -Key`: Aufgabe erledigt (`MarkComplete`), Termin ohne Erinnerung und mit „Erledigt: “ im Betreff.
- `Remove-HUOutlookItem -Type Task|Appointment -Key`: löschen, z. B. bei Storno.
- `Show-HUModuleMail -ModuleId -Sender <Angabe> -To -Cc -Subject -Body -Attachments` (ab 0.1.10): Mail anzeigen mit Absender. **Module nehmen diese Funktion.**
  - `-Sender`: Absender aus den Einstellungen des Moduls, `@{ Address; Mode }` mit `Mode` = `konto` (Konto im Outlook-Profil), `imAuftrag` (gemeinsames Postfach) oder `auto`. Leer: Bei einem Konto wird es genommen, bei mehreren fragt HU-Amt und merkt die Wahl auf Wunsch je Benutzer.
  - Fehlt das eingestellte Konto auf dem PC, fragt HU-Amt nach. Bei Abbruch ist `Way` = `abgebrochen`.
  - Für die Einstellungsseite: `Select-HUMailSender -Current <Angabe>` öffnet die Auswahl und gibt die Angabe zum Speichern zurück (`$null` bei Abbruch, `Address` leer = jeder wählt selbst).
- `Show-HUMail -From -FromMode`: die Grundfunktion darunter. Rückgabe `Way` (`outlook`, `mailto`, `fehler`, `testmodus`), `Message` und `From`. Mit festem Absender weicht sie nie auf `mailto:` aus.
- `Get-HUOutlookAccounts`: Mailkonten im Outlook-Profil (`Address`, `Name`, `Exchange`). Gemeinsame Postfächer ohne eigenes Konto stehen dort nicht.
- `Get-HUOutlookAttachments -Days 30 -Extensions '.pdf' -TargetDir <Ordner> [-Folder <Angabe>]` (ab 1.0.2): Anhänge aus dem Posteingang speichern, neueste zuerst. Die Mails bleiben unverändert (nicht verschoben, nicht als gelesen markiert). Rückgabe je Anhang `EntryId`, `StoreId`, `Subject`, `Sender`, `SenderAddress`, `Received`, `FileName`, `Path`. Im Testmodus gesperrt.
  - `-Mailbox <Adresse>` (ab 1.0.4): Posteingang eines anderen Postfachs statt des eigenen. Zuerst ein Konto im Outlook-Profil mit dieser Adresse, sonst ein freigegebenes Postfach (Exchange, Leserecht nötig). Unbekannte Adresse oder fehlendes Recht: Fehler mit verständlicher Meldung. Mehrere Postfächer: je Adresse einmal aufrufen.
- `Save-HUDroppedFiles -Data <e.Data> -TargetDir <Ordner> [-Problems <Liste>]` (ab 1.0.4): gezogene Dateien als Pfade, für eigene Fenster eines Moduls. Aus dem Explorer die vorhandenen Dateien, aus Outlook (Anhang direkt aus der Mail, kommt ohne `FileDrop`) die Anhänge, geschrieben nach `-TargetDir`. Fürs Ziehen selbst: `Test-HUDropHasFiles` und `Get-HUDropNames` (Namen schon während des Ziehens). Auf den Reiter gezogene Outlook-Anhänge löst der Kern selbst, das Modul bekommt in `Invoke-HUAddonDrop<Id>` wie immer Pfade.

**Ohne Outlook**

- `New-HUIcsFile -Path -Uid -Subject -Start [-End] [-AllDay] [-Location] [-Description] [-AlarmMinutes] [-Sequence]`: Kalenderdatei nach RFC 5545. Gleiche `Uid` für denselben Termin, `Sequence` bei Änderungen hochzählen.
- `New-HUMailtoUri`: Mail-Adresse für das Standard-Mailprogramm, zu langer Text wird gekürzt.

**Word**

- `Invoke-HUWordTemplate -Template <.docx> -PdfPath [-DocxPath] -Values @{ '{Mieter}' = '...' } -Tables @{ '{Positionen}' = @(@('Position', 'Preis'), @('Saal', '270,00')) }`
  - Platzhalter werden im Text, in Kopf- und Fußzeilen ersetzt, ohne Längengrenze. Ein Tabellen-Platzhalter steht allein in seinem Absatz, die erste Zeile wird Kopfzeile.
  - Die Vorlage bleibt unverändert, Word arbeitet an einer Kopie im Temp-Ordner, unsichtbar, und wird immer beendet.
  - Antwortet Word `-TimeoutSeconds` lang nicht (Standard 120, z. B. wegen eines unsichtbaren Dialogs), beendet der Kern genau diese Word-Instanz und meldet einen Fehler. Ein von Hand geöffnetes Word bleibt unberührt.
  - Am besten in einem Hintergrund-Auftrag aufrufen, damit das Fenster nicht stehen bleibt.
  - Rückgabe `Pdf`, `Docx` und `Missing` (Platzhalter, die in der Vorlage fehlen).
  - Ab 0.1.12 nur `.docx` und `.dotx` (sonst eine klare Fehlermeldung, bevor Word startet). Makros bleiben immer aus: Der Kern setzt vor dem Öffnen `AutomationSecurity = 3` (msoAutomationSecurityForceDisable). Ein per Programm gestartetes Word hätte sonst alle Makros an (Microsoft Learn, `MsoAutomationSecurity`).

Beispiel:

```powershell
$r = Set-HUOutlookAppointment -Folder $einst.Kalender -Key "Buchung:$($b.Id)" -EntryId $b.OutlookId `
    -Subject "Saal: $($b.Titel)" -Start $b.Beginn -End $b.Ende -ReminderMinutes 1440
$b.OutlookId = $r.EntryId
```

## Tests im strengen Modus

`Main.ps1` läuft mit `Set-StrictMode -Version 1.0`: eine Variable, die gelesen wird, bevor sie gesetzt ist, ist ein Fehler. Die Tests eines Moduls laufen deshalb ebenfalls so (`Set-StrictMode -Version 1.0` am Anfang von `BeforeAll` bzw. im Modultest). Variablen im Skriptbereich (`$script:…`) oben in der Datei mit einem Startwert setzen.

## Oberfläche: was der Kern für jedes Modul erledigt

- **Verschiebbare Bereiche (GridSplitter):** Der Kern merkt sich in jedem Reiter und in jedem Fenster mit `Restore-HUDialogState`/`Save-HUDialogState` die Zeilen- und Spaltengrößen aller Grids mit Splitter (je Benutzer, ab 0.1.10). **Module bauen dafür nichts Eigenes.** Ein Grid bekommt am besten ein `x:Name`, dann bleibt die Zuordnung auch bei Umbauten stabil.
- **Fenstergröße und -position:** `Restore-HUDialogState $w '<Name>'` beim Öffnen, `Save-HUDialogState $w '<Name>'` beim Schließen.
- **Tabellen:** Spaltenbreite nach Inhalt, Knöpfe wachsen mit der Beschriftung, Anfasser an jedem größenveränderbaren Fenster (über `New-HUWindow`).
- **Listen: Anlegen und Entfernen (ab 1.0.1):** Jede Liste, in der man Einträge anlegen kann, hat auch „Entfernen“.
  - Noch nicht gespeicherte Einträge: immer entfernbar, ohne Rückfrage.
  - Gespeicherte Einträge: nur entfernbar, wenn kein Datensatz darauf verweist (per ID prüfen). Sonst eine Meldung, wo der Eintrag verwendet wird, und der Hinweis auf „inaktiv“ bzw. „ausblenden“. Nichts Verwendetes löschen, damit alte Vorgänge vollständig bleiben.
  - Vor dem Entfernen eines gespeicherten Eintrags Rückfrage, das Entfernen mit Benutzer und Zeit festhalten.
- **Farbschema Dunkel/Hell (ab 1.0.0):** Jeder Benutzer wählt es unter Einstellungen › Allgemein. Module gestalten weiter dunkel wie bisher, der Kern setzt die Farben für Hell beim Laden um.
  - XAML: automatisch über `New-HUWindow`, `Get-HUThemedXaml`, den Reiter des Moduls und `Show-HUToast`. Nichts zu tun.
  - Farben im Code: immer `Get-HUBrush '<#RRGGBB>' '<Rolle>'`, nie `BrushConverter`, `ColorConverter` oder `[Windows.Media.Brushes]::…`. Rollen: `Fg` Schrift, `Bg` Fläche, `Border` Rahmen, `Btn` graue Knopffläche mit weißer Schrift. Ohne Rolle bleiben Weiß und kräftige Farben unverändert.
  - Farben in Datenbindungen (z. B. `Farbe = '#9E9E9E'` in Zeilen): `ConvertTo-HUThemeColor '<#RRGGBB>' 'Fg'`.
  - Abfrage: `Test-HULightTheme` (z. B. für eine eigene HTML-Seite in WebView2).
  - Sichtprüfung: Starttest mit `HUAMT_SMOKE_THEME=Hell` und `HUAMT_SMOKE_SHOTS=<Ordner>` legt je Reiter und Fenster ein Bild ab, ohne sichtbares Fenster.

## Kontakte übernehmen

Ab Kern 0.1.10. `Select-HUContact [-Multiple]` öffnet „Kontakt übernehmen“: Suche, Outlook-Kontakte (eigener Ordner Kontakte), vCard-Dateien öffnen oder hineinziehen. Rückgabe: die markierten Kontakte mit `Name`, `Firma`, `Mail`, `Telefon`, `Mobil`, `Strasse`, `Plz`, `Ort`, `Land`, `Quelle`; das Modul übernimmt die Felder selbst.

- Ohne Fenster: `Get-HUOutlookContacts [-Search] [-Max]`, `Import-HUVCardFile -Path`, `ConvertFrom-HUVCard -Text`, `Find-HUContact`.
- vCard 4.0, 3.0 und 2.1 (Outlook-Export mit Quoted-Printable). Im Testmodus geht nur vCard.

## Testmodus und geführter Test

Ab Kern 0.1.9 (`minApp` 0.1.9). Den Testmodus startet ein Verwalter in den Einstellungen; HU-Amt startet dafür neu.

- **Daten:** Im Testmodus ist der Datenstamm `Data\_Test\` statt `Data\`. Wer seine Daten über `Get-HUAddonDataDir` bzw. `$script:DataRoot` holt (Pflicht), arbeitet automatisch im Testbestand. Nie einen Pfad aus dem Programmordner selbst zusammensetzen.
- **Erkennen:** `Test-HUTestMode` (im Fenster) bzw. `Test-HUOfficeTestMode` (auch in Hintergrund-Aufträgen). Im Test zeigt das Modul statt Outlook/Word eine Simulation, z. B. „Termin würde eingetragen“.
- **Sperre:** `Open-HUOutlook`, alle Outlook-Funktionen ohne Sitzung und `Invoke-HUWordTemplate` werfen im Testmodus, `Show-HUMail` öffnet nichts (`Way` = `testmodus`). `.ics`-Dateien gehen weiter.
- **Zurücksetzen** macht der Kern (Knopf in der orangen Leiste): gelöscht wird beim nächsten Start alles unter `Data\_Test\`, nur wenn dort die Kennungsdatei `.hu-testbestand` liegt. Module löschen keine Testdaten selbst.

**Geführter Test:** Das Modul liefert nur Szenarien, Fenster und Protokoll macht der Kern.

```powershell
Show-HUTestAssistant -ModuleId 'Buchung' -ModuleVersion $ver -Scenarios @(
    @{ Name = 'Heimatsaal mit Teilfläche und Rabatt'; Steps = @(
        @{ Action = 'Neue Buchung: Saal 1/2, Samstag 18–23 Uhr'; Expected = 'Wird gespeichert'; Check = { @{ Ok = (Test-...); Text = '1 Buchung gefunden' } } }
        @{ Action = 'Zweite Buchung: Saal 2/3, gleiche Zeit'; Expected = 'Wird abgelehnt' }
    ) }
)
```

- `Check` ist optional und gibt `$true`/`$false` oder `@{ Ok; Text }` zurück. Ein Fehler darin zählt als „Fehler“.
- Automatische Ergebnisse und die Bestätigung des Menschen bleiben getrennt, auch im Protokoll (CSV und HTML unter `Data\_Test\Testprotokolle\`).
- Den Knopf „Test starten“ zeigt das Modul nur, wenn `Test-HUTestMode` wahr ist und `Test-HUDemoMode` nicht (siehe [Demo-Modus](#demo-modus)). Testdaten und Preise im Test als Beispielwerte kennzeichnen.

## Gemeinsame Bausteine des Kerns (ab 0.1.11)

Grundsatz: Das Hauptprogramm liefert alles, was mehr als ein Modul braucht. Ein Modul enthält nur seine Fachlogik und baut nichts nach, was der Kern hat. Fehlt etwas, das mehrere Module brauchen, kommt es in den Kern. Module, die diese Bausteine nutzen, setzen `minApp` auf 0.1.11.

Die Kern-Module (`Core\*.psm1`) laufen auch in einem Hintergrund-Runspace ohne Fenster. `Start-HUJob` lädt sie dort automatisch; ein eigener Runspace lädt sie mit `Import-Module (Join-Path $AppRoot 'Core\HU.Data.psm1')`.

| Baustein | Datei | Wichtigste Funktionen |
|---|---|---|
| Datenablage | `Core\HU.Data.psm1` | `Read-HUJson` (mit Wiederholung, `@{Data; Error}`), `Get-HUFileStamp`, `Save-HUDataFile -Build {…}` (Sperre, Warten, Stand-Prüfung, Kopie, atomar), `Wait-HUDataFileLock`, `Write-HUDataFile`, `Get-HUDataRecords` (Zwischenspeicher nach Dateistempel), `New-HUDataId`, `New-HUShortId`, `Get-HUNextNumber`, `Get-HUProp`, `ConvertFrom-HUJsonText` (ohne Datums-Umwandlung) |
| Zahlen, Euro, Datum | `Core\HU.Data.psm1` | `ConvertTo-HUDecimal`, `Test-HUAmbiguousNumber` (ab 0.1.12), `Format-HUEuro`, `Format-HUNumber`, `Format-HUKm`, `Get-HURound` (kaufmännisch), `ConvertFrom-HUDateText` (auch Excel-Seriennummer, `-DateOnly`), `ConvertFrom-HUClockText`, `Format-HUDate`, `Format-HUTime`. Tausenderpunkt fest, unter PS 5.1 und 7 gleich. |
| Verlauf | `Core\HU.Data.psm1` | `New-HUHistoryEntry`, `Add-HUHistoryEntry` (`zeit, benutzer, pc, aktion, text, felder`), `Compare-HURecordFields`, `Get-HUFieldChangeText`, `ConvertTo-HUHistoryEntry` (liest alte Felder `von`/`art`) |
| Anhänge | `Core\HU.Data.psm1` | `Test-HUAttachment`, `Add-HUAttachment` (überschreibt nie, `-TimePrefix`), `Get-HUAttachments`, `Remove-HUAttachment` (nach `_entfernt\`) |
| Gemeinsame Einstellungen | `Core\HU.Data.psm1` | `Read-HUSettings`, `Save-HUSettings` (`geaendert`, `geaendertVon`), `Merge-HUSettings`, `Test-HUPermittedUser $Liste $Benutzer -IsVerwalter`, `Get-HUUserDisplayName` |
| PDF erzeugen | `Core\HU.Pdf.psm1` | `New-HUPdf [-Landscape]`, `Add-HUPdfText`, `Add-HUPdfLine`, `Add-HUPdfRect`, `ConvertTo-HUPdfJpeg`, `Add-HUPdfImage`, `Add-HUPdfTable` (über mehrere Seiten mit Kopfzeile), `Add-HUPdfPageFooter`, `Get-HUPdfBytes -Producer '<Modulname>'` |
| QR-Codes (ab 1.0.2) | `Core\HU.Qr.psm1` | `New-HUQrPayload -Prefix -Data` (JSON, komprimiert, höchstens 2300 Zeichen), `Read-HUQrPayload -Text -Prefix`, `Add-HUPdfQrCode -Pdf -Text -X -Y -Size` (in ein HU.Pdf-Dokument), `Read-HUQrCodes -Path` (Bild oder gescanntes PDF, Rückgabe `Texts` und `Hint`). Bibliothek ZXing.Net im Kern, Module bringen dafür nichts mit. |
| Excel und CSV | `Core\HU.Import.psm1` | `Read-HUImportFile`, `Read-HUXlsx [-SheetJs]`, `Read-HUCsv` (Kodierung und Trennzeichen erkannt), `New-HUXlsxBytes` (Excel schreiben, ohne Excel) |
| Werkzeug-Aufrufe | `Core\HU.Tools.psm1` | `Invoke-HUNativeProcess` (Zeitlimit, Prozessbaum beenden), `ConvertTo-HUProcessArgument`, `Stop-HUProcessTree`, `ConvertTo-HUSafeFileName`, `Get-HUFreeFilePath`, `Get-HUFreeDirectoryPath` |
| Fenster und Steuerelemente | `Functions\UI-Controls.ps1` | `New-HUModuleWindow` (XAML aus dem Modulordner, Thema, Zustand), `New-HUComboItem`, `Set-HUComboTag`, `Get-HUComboTag`, `Add-HUCalendar` (dunkler Kalender, Umschalt+Pfeil = Jahr), `Initialize-HUTimeCombo`, `Format-HUNumberInput` |
| Dialoge | `Functions\UI-Controls.ps1` | `Read-HUText` (`$null` bei Abbruch, `-AllowEmpty`, `-Mehrzeilig`), `Read-HUChoice` (Index oder -1, ab 0.1.12 `-DefaultIndex`), `Read-HUNumberInput` (ab 0.1.12) |
| Hintergrund | `Functions\UI-Controls.ps1` | `Start-HUBackgroundJob` (Rückruf `Done($Ergebnis, $Fehler, $Kontext)`, `-Sta` für Outlook, Zeitlimit), `Stop-HUBackgroundJobs -Owner`, `Register-HUBackgroundLoad` (Laden mit Zeitgeber, Abgleich alle n Sekunden, Warteschlange, Punkt für „lädt“) |

Worauf beim Umstellen zu achten ist:
- `-GetArguments` von `Register-HUBackgroundLoad` gibt seine Argumente in die Pipeline. Ein Array als ein Argument braucht ein Komma davor (`,[string[]]$x`).
- Aufträge holt der Kern mit eigenem Zeitgeber ab; `Start-HUBackgroundJob` im Fenster-Thread aufrufen.
- `ConvertFrom-HUDateText` behält die Uhrzeit; nur das Datum mit `-DateOnly`.
- `HU.Data` lädt `HU.Paths` und `HU.Exchange` nur für sich selbst. Wer in einem eigenen Runspace oder Prozess `Unlock-HUDataFile`, `Get-HULockInfoText` o. ä. direkt aufruft, importiert `HU.Paths` selbst. In der App und in `Start-HUJob` ist es schon geladen.
- `Get-HUPdfBytes` ohne `-Producer` schreibt nur „HU-Amt“ in die Dokumenteigenschaften.
- `Get-HUFreeFilePath` legt die Datei nicht an. Schreiben mehrere PCs gleichzeitig in denselben Ordner, mit `CreateNew` anlegen und bei Fehler neu suchen.

### Geändertes Verhalten ab 0.1.12

- **Zahlen und Datum als Text in Tabellen** (`Export-HUTable`, Spaltentyp `number` bzw. `date`):
  - Zahl mit Komma: Komma = Dezimalzeichen, Punkt = Tausender in Dreiergruppen (`1,5`, `12,50`, `1.234,56`).
  - Zahl ohne Komma mit höchstens einem Punkt: Punkt = Dezimalzeichen (`1.5`, `1234.56`, z. B. aus Programmdaten), wie `ConvertTo-HUDecimal`. Mehrere Punkte in Dreiergruppen sind Tausender (`1.234.567`).
  - Englische Schreibweise mit Komma als Tausender (`1,234.56`) gilt nicht mehr als Zahl und bricht mit Zeile und Spalte ab. Bisher wurde `1,5` still zu 15.
  - Datum nur in festen Formaten: `d.M.yyyy` (auch mit `H:mm` bzw. `H:mm:ss`) und ISO `yyyy-MM-dd` (mit Uhrzeit, Sekundenbruchteil, Zone). Bisher wurde `05.11.2026` zum 11. Mai.
  - `[decimal]`, `[double]` und `[datetime]` bleiben, wie sie sind.
- **`Get-HUDataRecords` bei nicht erreichbarem Ordner:**
  - Fehlt die Wurzel (Laufwerk bzw. `\\server\freigabe`), oder fehlt der Elternordner und der Zwischenspeicher kennt dort schon Datensätze, kommt in `Errors` „Ordner nicht erreichbar: <Pfad>“.
  - Der Zwischenspeicher bleibt dann erhalten und seine Datensätze kommen als `Items`. Das Modul zeigt den alten Stand und meldet den Fehler, statt eine leere Liste zu speichern.
  - Leer ohne Fehler nur, wenn der Ordner wirklich fehlt: Elternordner da, oder dort waren noch nie Datensätze (Unterordner, die erst beim ersten Speichern entstehen).
- **`ConvertTo-HUSafeFileName`:** Gerätenamen auch mit Endung (`CON.txt` wird `_CON.txt`, `nul.json` wird `_nul.json`).
- **Abbrechen von Hintergrund-Aufträgen** (Zeitüberschreitung, `Stop-HUBackgroundJobs`, `Stop-HUBackgroundLoad`): kehrt sofort zurück. Der Kern gibt den Auftrag frei, sobald er wirklich steht.
- **`Show-HUMessage -Icon`:** `Information` geht als anderer Name für `Info`.
- **`Read-HUChoice -DefaultIndex <n>`:** Dieser Knopf reagiert auf Enter. Ohne Angabe hat wie bisher kein Knopf diese Rolle.

### Betragsfelder: zweideutige Eingabe 1.500

`ConvertTo-HUDecimal '1.500'` ergibt 1,5 (Punkt ohne Komma = Dezimalzeichen). Das bleibt so, Daten aus Dateien werden nicht anders gelesen. Tippt aber jemand in ein Betragsfeld `1.500`, kann 1500 gemeint sein. Dafür gibt es ab 0.1.12:

- `Test-HUAmbiguousNumber <Text>`: `$true` bei genau einem Punkt mit genau drei Ziffern danach und ohne Komma (`1.500`, `12.250`, `-3.000`).
- `Read-HUNumberInput <Text> [-Owner] [-Title]`: fragt bei solchen Eingaben einmal „Meinst du 1.500 (also 1500) oder 1,5?“ (Enter = 1500) und gibt die gewählte Zahl zurück. Sonst wie `ConvertTo-HUDecimal`. `$null` bei leerem Feld, keiner Zahl oder Abbruch.

Einsatz am Betragsfeld: beim Verlassen des Felds bzw. vor dem Speichern lesen und das Feld neu schreiben. Danach ist der Text eindeutig, es wird nicht noch einmal gefragt.

```powershell
$c.txtBetrag.Add_LostFocus({
        param($s, $e)
        if (-not "$($s.Text)".Trim()) { return }
        $v = Read-HUNumberInput $s.Text -Owner ([System.Windows.Window]::GetWindow($s))
        if ($null -ne $v) { $s.Text = Format-HUNumberInput $v } else { $s.Focus() | Out-Null }
    })
```

Nur für Eingaben im Fenster. Für Dateien und Hintergrund-Aufträge weiter `ConvertTo-HUDecimal` nehmen, die Rückfrage braucht den Fenster-Thread.

### Vorlagen für Modul-Repos

`HU-Amt\.github\modul-vorlage\` enthält `Format-Dateien.ps1`, `Test-Start.ps1` und `Deploy-Dev.ps1`. Module kopieren sie **unverändert** nach `.github\dev\`, wie die Wächter-Vorlage. Was nur ein Modul braucht, steht in `.github\dev\modul-dev.json`; Aufbau und Felder in der `LIESMICH.md` dort. `Deploy-Dev.ps1` bricht ab, wenn die Test-App läuft. Ab 0.1.12 gilt für die Modul-Id dieselbe Regel wie im Kern (Buchstabe, dann 1 bis 30 Buchstaben oder Ziffern), und Dateien, die kein gültiges UTF-8 sind, werden nicht umgeschrieben, sondern als Warnung gemeldet.

## Beispieldaten für den Testmodus

Ein Modul kann einen Ordner `Testdaten\` mitbringen. Er ist aufgebaut wie der Datenordner des Moduls (`Data\<Id>\`). Startet ein Verwalter den Testmodus, kopiert der Kern den Inhalt in den Testbestand `Data\_Test\<Id>\`, bevor das Modul lädt (ab 0.1.14). Das geschieht nur, wenn dort für das Modul noch nichts liegt, also beim ersten Start im Testmodus und nach „Test zurücksetzen“. Echte Daten werden nie berührt. In der Demo kommen dieselben Beispieldaten bei jedem Start neu (siehe [Demo-Modus](#demo-modus)).

Regeln:
- Nur künstliche Daten: ausgedachte Namen, Adressen, Nummern und Beträge, nichts aus einer echten Gemeinde.
- Dasselbe Format, das das Modul selbst schreibt, mit Feld `schema`. Ein Test im Modul liest jede Datei aus `Testdaten\` mit den eigenen Lesefunktionen.
- Klein halten: einige typische Fälle, dazu je ein Grenzfall (z. B. storniert, überfällig, ohne Kostenstelle).
- Persönliche Einstellungen (`Benutzer\`) gehören nicht hinein.

## Demo-Modus

Ab Kern 1.0.0. Die Demo startet jeder Benutzer unter Einstellungen › Allgemein › „Demo starten …“; HU-Amt startet dafür neu, mit blauer Leiste statt der orangen.

- **Eine Form des Testmodus:** In der Demo ist `Test-HUDemoMode` wahr und `Test-HUTestMode` ebenfalls. Was im Testmodus gesperrt ist (Outlook, Word, Mail), ist es auch in der Demo, ohne dass das Modul etwas tun muss.
- **Daten:** Der Datenstamm ist `Data\_Demo\<Benutzer>_<PC>\`, je Benutzer und PC ein eigener Bestand. Bei jedem Demo-Start leert der Kern ihn und legt die Beispieldaten aus `Testdaten\` neu hinein, bevor das Modul lädt. `Get-HUAddonDataDir` bzw. `$script:DataRoot` zeigen automatisch dorthin.
- **Beispieldaten auf heute beziehen:** Weil die Demo immer frisch ist, darf ein Modul seine Beispieldaten nach dem Laden an das heutige Datum anpassen, z. B. Termine und Fristen relativ zu heute verschieben. Nur im Demo-Bestand, nie in echten Daten.
- **Werkzeuge nur für Tests ausblenden:** Was nur zum Prüfen da ist (z. B. „Test starten“ für den geführten Test), zeigt das Modul in der Demo nicht: `if ((Test-HUTestMode) -and -not (Test-HUDemoMode)) { ... }`. Den geführten Test selbst lehnt der Kern in der Demo ab.
- Demo-Bestände kommen nie in die Sicherung und ziehen beim Wechsel des Datenordners nicht mit.

## Datenordner

Wo die Daten liegen, bestimmt der Kern: `<App>\Data` oder der Ordner aus `Config\datenort.json` (Einstellungen › Verwalter, ab 0.1.14). Module nehmen ihren Ordner deshalb immer über `Get-HUAddonDataDir` bzw. `$script:DataRoot` und bauen nie selbst einen Pfad aus dem Programmordner und `Data`. Dann ziehen sie beim Umzug ohne Änderung mit, Testmodus inklusive. Nur `Data\Laufend` (laufende Apps) bleibt immer beim Programm.

## Rechte auf der Freigabe

HU-Amt läuft meist aus einem Netzordner. Wer dort was darf, entscheiden Freigabe- und NTFS-Rechte, nicht die App. Die App kann keine Rechte erzwingen, sie zeigt nur, was ihr auffällt.

Empfehlung:

- **Freigabe (SMB):** „Ändern“ für die Benutzer der App (z. B. eine eigene Gruppe). Die strengeren NTFS-Rechte darunter bestimmen den Rest.
- **NTFS `Data`:** „Ändern“ für die Benutzer, aber **ohne „Vollzugriff“**.
  - Warum: Vollzugriff enthält „Unterordner und Dateien löschen“ und „Berechtigungen ändern“. Wer das am Ordner hat, kann jede Datei darin löschen und neu anlegen, auch wenn die Datei selbst geschützt ist.
- **NTFS `Data\Verwalter.txt`:** „Lesen“ für die Benutzer, „Ändern“ nur für die Verwalter, **Vererbung an dieser Datei unterbrochen**.
  - Warum: Wer die Liste ändern darf, macht sich selbst zum Verwalter und sieht Protokoll, Sicherung, Update-Einstellungen und Module. Ohne unterbrochene Vererbung bekäme die Datei das „Ändern“ des Ordners.

Was die App prüft (`Test-HUVerwalterAcl`, ab 0.1.12):

- Sie liest die NTFS-Rechte von `Data\Verwalter.txt` und `Data` (auch über UNC) und meldet, wenn breite Gruppen schreiben dürfen: Jeder, Authentifizierte Benutzer, Benutzer, Domänen-Benutzer. Erkannt wird über die SIDs (`S-1-1-0`, `S-1-5-11`, `S-1-5-32-545`, `…-513`), damit es auf deutschem und englischem Windows gleich geht. Geerbte Rechte zählen mit.
- An der Datei zählen Schreiben, Anhängen, Löschen, Berechtigungen ändern und Besitz übernehmen. Am Ordner zählen „Unterordner und Dateien löschen“, Berechtigungen ändern und Besitz übernehmen (also Vollzugriff, nicht „Ändern“). Fehlt die Liste, zählt am Ordner „Dateien erstellen“.
- „Verweigern“-Einträge werden nicht verrechnet. Die Prüfung kann deshalb warnen, obwohl ein Verweigern-Eintrag das Recht wieder nimmt.
- **Freigaberechte (SMB) sieht die App nicht**, sie liegen an der Freigabe und nicht in der NTFS-Liste. Sie werden nicht geprüft, auch nicht, wenn die Zeile „keine breiten Schreibrechte“ zeigt.
- Der Verwalter sieht das Ergebnis unter Einstellungen › Verwalter als Zeile mit Tooltip. Beim Start steht eine Warnung nur im Protokoll, normale Benutzer bekommen keine Meldung. Lassen sich die Rechte nicht lesen, steht dort still „nicht prüfbar“.

- `Config\datenort.json` prüft sie ab 0.1.14 mit: Wer diese Datei ändern darf, lenkt alle Daten samt Verwalter-Liste um. `Config` braucht wie das Programm nur Lesen für die Benutzer.

Wer unter `Data` schreiben darf, kann auch Modul-Einstellungen ändern, etwa Ordnerangaben der Texterkennung in `ordner.json` oder gemeinsame Verläufe. Module prüfen deshalb Pfade aus solchen Dateien selbst, bevor sie etwas öffnen oder starten. Die Texterkennung tut das seit 0.1.1. Die App ersetzt also keine sauber gesetzten Freigabe- und NTFS-Rechte.

## Externe Werkzeuge und Bibliotheken

Bringt ein Modul fremde Programme, DLLs oder Pakete mit, gilt:

- Feste Version mit SHA256, nie „latest“. Neue Stände kommen nur mit einem getesteten Release.
- Der Wächter meldet einmal pro Woche neuere Versionen, bekannte Lücken und das Ende der Pflege. Das Modul liefert dafür nur `.github\waechter.json` und eine unveränderte Kopie von [`.github/waechter/Vorlage-Modul.yml`](.github/waechter/Vorlage-Modul.yml) als `.github\workflows\waechter.yml`. Aufbau von `waechter.json`: [`.github/waechter/LIESMICH.md`](.github/waechter/LIESMICH.md).
