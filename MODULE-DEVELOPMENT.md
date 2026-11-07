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
- Den Knopf „Test starten“ zeigt das Modul nur, wenn `Test-HUTestMode` wahr ist. Testdaten und Preise im Test als Beispielwerte kennzeichnen.

## Externe Werkzeuge und Bibliotheken

Bringt ein Modul fremde Programme, DLLs oder Pakete mit, gilt:

- Feste Version mit SHA256, nie „latest“. Neue Stände kommen nur mit einem getesteten Release.
- Der Wächter meldet einmal pro Woche neuere Versionen, bekannte Lücken und das Ende der Pflege. Das Modul liefert dafür nur `.github\waechter.json` und eine unveränderte Kopie von [`.github/waechter/Vorlage-Modul.yml`](.github/waechter/Vorlage-Modul.yml) als `.github\workflows\waechter.yml`. Aufbau von `waechter.json`: [`.github/waechter/LIESMICH.md`](.github/waechter/LIESMICH.md).
