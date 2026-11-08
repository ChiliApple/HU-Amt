# Vorlagen für Modul-Repos

Werkzeuge zum Entwickeln eines Moduls auf dem eigenen PC. Sie liegen nur im Kern und gelten für alle Module. Ausgeliefert wird davon nichts, weder mit der App noch mit einem Modul.

## Was ein Modul übernimmt

Die drei Skripte kommen **unverändert** nach `.github/dev/` im Modul-Repository:

| Vorlage | im Modul | Zweck |
|---|---|---|
| `Format-Dateien.ps1` | `.github/dev/Format-Dateien.ps1` | Textdateien auf UTF-8 mit BOM und CRLF bringen (`.yml` und `.py` ohne BOM, mit LF). Vor jedem Commit. Beachtet `formatAusschluss` und `ohneFormat` aus `modul-dev.json` wie `Deploy-Dev.ps1`. |
| `Deploy-Dev.ps1` | `.github/dev/Deploy-Dev.ps1` | Modul in die Test-App kopieren (`<App>\AddonsDev\<Modul-Id>`). |
| `Test-Start.ps1` | `.github/dev/Test-Start.ps1` | Starttest der App mit nur diesem Modul in einer Temp-Kopie der Test-App, unter PowerShell 5.1 und 7. |

Ändert sich eine Vorlage im Kern, wird sie wieder unverändert kopiert. Was nur ein Modul braucht, steht in `modul-dev.json` (unten) oder in einem eigenen Skript daneben, nie in der Kopie.

Die Skripte gehen davon aus, dass das Modul der Hauptordner des Repositorys ist (wie in einem Release) und sie zwei Ebenen darunter liegen. Die Modul-Id kommt aus `addon.json` (`id`).

## Aufruf

Mit vollem Pfad, PowerShell startet sonst in `System32`:

```powershell
& 'C:\...\HU-Amt-<Name>\.github\dev\Format-Dateien.ps1'
& 'C:\...\HU-Amt-<Name>\.github\dev\Deploy-Dev.ps1'
& 'C:\...\HU-Amt-<Name>\.github\dev\Test-Start.ps1'
```

Gemeinsame Parameter von `Deploy-Dev.ps1` und `Test-Start.ps1`:

| Parameter | Standard | Bedeutung |
|---|---|---|
| `-ModulId` | `id` aus `addon.json` | Ordnername unter `AddonsDev` |
| `-Quelle` | Hauptordner des Repositorys | Modulordner |
| `-App` | `C:\Tools\HU-Amt` | Test-App (muss `Main.ps1` enthalten) |
| `-MinAppAnpassen` | aus | Nur Entwicklung: Ist die Test-App älter als `minApp` des Moduls, bekommt die Kopie die Version der Test-App als `minApp`. Die Quelle bleibt unverändert. |

Nur `Deploy-Dev.ps1`:

| Parameter | Standard | Bedeutung |
|---|---|---|
| `-VerboteneEndungen` | `.log .exe .dll .zip .pdf .png .jpg .jpeg .csv .pyc` | Dateien mit diesen Endungen bricht das Kopieren ab |
| `-VerboteneNamen` | `einstellungen.json __pycache__` | Datei- oder Ordnernamen, die nicht ins Modul gehören |
| `-ErlaubteNamen` | leer | Dateien, die trotz Endung erlaubt sind |
| `-FormatAusschluss` | `^\.github\\dev\\Werkzeuge\\` | Regex auf den relativen Pfad: nicht umformatieren |
| `-OhneFormat` | aus | Keine Textdatei umformatieren |

Der Ordner `Testdaten\` (Beispieldaten für den Testmodus) ist von beiden Prüfungen ausgenommen und darf Datendateien, PDFs und Bilder enthalten. Damit Git Beispiel-PDFs nicht umwandelt, gehört `*.pdf binary` (und Bilder) in `.gitattributes` des Modul-Repos.

Nur `Test-Start.ps1`:

| Parameter | Standard | Bedeutung |
|---|---|---|
| `-PowerShell` | `beide` | `5`, `7` oder `beide` |
| `-Erwartet` | leer | Zusätzliche Muster (Regex), die im Ergebnis stehen müssen, z. B. ein Schritt des Moduls |
| `-TestMode` | aus | Starttest im Testmodus: Beispieldaten aus `Testdaten\` landen im Testbestand der Kopie |
| `-TimeoutSec` | `180` | Zeitlimit je Lauf |

## Einstellungen je Modul: `modul-dev.json`

Optional in `.github/dev/` neben den Skripten. Damit laufen die Skripte ohne Parameter. Angegebene Parameter gehen vor.

```json
{
  "erwartet": ["Ziehen-Beispiel"],
  "verboteneEndungen": [".log", ".exe", ".dll", ".zip", ".pdf", ".png", ".jpg", ".jpeg", ".csv", ".ics"],
  "verboteneNamen": ["einstellungen.json", "zaehler.json"],
  "erlaubteNamen": ["Testseite.pdf"],
  "formatAusschluss": "^web\\\\lib\\\\",
  "ohneFormat": false,
  "minAppAnpassen": true
}
```

Jeder Eintrag ist optional. `verboteneEndungen` und `verboteneNamen` ersetzen die Standardliste, sie ergänzen sie nicht.

## Was die Skripte schützen

- `Deploy-Dev.ps1` bricht ab, wenn die Test-App läuft (Prozess mit `<App>\Main.ps1` in der Befehlszeile) oder eine Datei im Ziel gesperrt ist. Dann bleibt das Ziel unverändert.
- `Deploy-Dev.ps1` bricht ab, wenn im Modul Daten, Logs, Belege oder Programme liegen. Die gehören nach `Data` bzw. in das Werkzeugpaket.
- `Test-Start.ps1` liest die Test-App nur. Gestartet wird eine Kopie ohne `Data`, `Addons`, `AddonsDev` und `Logs` im Temp-Ordner, die danach gelöscht wird.
