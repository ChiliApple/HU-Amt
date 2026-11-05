# HU-Amt

Werkzeuge für die Gemeindeverwaltung: ein Fenster mit Reitern, jeder Reiter ein Modul. Dazu Update per Knopf, Sicherung, Hilfe und dunkles Farbschema.

- Windows 10/11, Windows PowerShell 5.1 oder PowerShell 7
- ohne Installation, läuft aus einem Ordner am PC oder auf einer Freigabe

## Schnellstart

1. Ordner herunterladen und entpacken, z. B. nach `C:\Tools\HU-Amt`.
2. `Start-HU-Amt.cmd` doppelklicken.

## Für Module: Datei-Austausch mit anderen Programmen

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

## Für Module: Dateien an andere Module weitergeben

Der Kern hat ein Postfach je Modul im gemeinsamen Datenordner (`Data\Postfach\<Modul>\`). Nur einbauen, wenn die Funktionen vorhanden sind (`Get-Command`).

- `Get-HUAddonTargets -Extension '.pdf' [-Exclude <Id>]`: Module, die diese Endung annehmen (`Id`, `Title`).
- `Send-HUAddonFiles -To <Id> -Paths <...> -From <Id> [-Open]`: gibt Dateien weiter, Rückgabe ist die Anzahl.
  - Ohne `-Open` kommt eine Kopie ins Postfach. Das klappt auch zwischen PCs.
  - Mit `-Open` wechselt die App zum Reiter und übergibt die Dateien sofort.
- `Get-HUAddonInbox <Id>`: wartende Einträge (`Key`, `Path`, `Name`, `From`, `Pc`, `User`, `Time`), nur mit den Endungen aus `drop`.
- `Remove-HUAddonInboxItem <Eintrag>`: Eintrag samt Kopie entfernen, erst nach gelungener Übernahme.
- `Update-HUAddonInboxBadge [<Id>]`: Zahl am Reiter neu zählen. Der Kern zählt zusätzlich jede Minute.

## Lizenz

Nutzungslizenz, siehe [LICENSE](LICENSE).
