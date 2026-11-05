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

- **Kennung:** Buchstaben und Ziffern, beginnt mit einem Buchstaben. Reserviert sind `Benutzer`, `Laufend`, `Docs` und `Postfach`.
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

## Externe Werkzeuge und Bibliotheken

Bringt ein Modul fremde Programme, DLLs oder Pakete mit, gilt:

- Feste Version mit SHA256, nie „latest“. Neue Stände kommen nur mit einem getesteten Release.
- Der Wächter meldet einmal pro Woche neuere Versionen, bekannte Lücken und das Ende der Pflege. Das Modul liefert dafür nur `.github\waechter.json` und eine unveränderte Kopie von [`.github/waechter/Vorlage-Modul.yml`](.github/waechter/Vorlage-Modul.yml) als `.github\workflows\waechter.yml`. Aufbau von `waechter.json`: [`.github/waechter/LIESMICH.md`](.github/waechter/LIESMICH.md).
