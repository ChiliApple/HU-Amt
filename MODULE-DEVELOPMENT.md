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
- `Select-HUOutlookFolder -Type Calendar|Tasks`: Ordner im Outlook-Dialog wählen, auch einen gemeinsamen Kalender. Rückgabe `Name`, `Path`, `EntryId`, `StoreId`, `Type` oder `$null` bei Abbruch. Diese Angabe speichert das Modul in seinen Einstellungen und gibt sie als `-Folder` weiter. Leer = Standardordner des Benutzers.
- `Set-HUOutlookAppointment -Folder -Key [-EntryId] -Subject -Start [-End] [-AllDay] [-Location] [-Body] [-ReminderMinutes] [-BusyStatus] [-Categories]`: anlegen oder denselben Termin ändern. Rückgabe `EntryId`, `StoreId`, `Created`.
- `Set-HUOutlookTask -Folder -Key [-EntryId] -Subject -DueDate [-StartDate] [-ReminderTime] [-Body] [-Categories]`: dasselbe für Aufgaben.
- `Complete-HUOutlookItem -Type Task|Appointment -Key`: Aufgabe erledigt (`MarkComplete`), Termin ohne Erinnerung und mit „Erledigt: “ im Betreff.
- `Remove-HUOutlookItem -Type Task|Appointment -Key`: löschen, z. B. bei Storno.
- `Show-HUMail -To -Cc -Subject -Body -Attachments`: Mail anzeigen. Rückgabe `Way` (`outlook`, `mailto`, `fehler`) und `Message`.

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

## Externe Werkzeuge und Bibliotheken

Bringt ein Modul fremde Programme, DLLs oder Pakete mit, gilt:

- Feste Version mit SHA256, nie „latest“. Neue Stände kommen nur mit einem getesteten Release.
- Der Wächter meldet einmal pro Woche neuere Versionen, bekannte Lücken und das Ende der Pflege. Das Modul liefert dafür nur `.github\waechter.json` und eine unveränderte Kopie von [`.github/waechter/Vorlage-Modul.yml`](.github/waechter/Vorlage-Modul.yml) als `.github\workflows\waechter.yml`. Aufbau von `waechter.json`: [`.github/waechter/LIESMICH.md`](.github/waechter/LIESMICH.md).
