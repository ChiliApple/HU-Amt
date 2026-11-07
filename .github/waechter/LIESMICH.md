# Wächter für externe Werkzeuge und Bibliotheken

Der Wächter prüft einmal pro Woche die Werkzeuge und Bibliotheken eines Moduls:

- neuere Versionen,
- bekannte Sicherheitslücken (osv.dev, NVD),
- Ende der Pflege.

Er ändert nichts. Befunde kommen in ein Issue im Repository des Moduls, bei Änderungen zusätzlich eine Teams-Nachricht.

Alles liegt nur im Kern und gilt für alle Module. Ausgeliefert wird davon nichts, weder mit der App noch mit einem Modul.

## Was ein Modul braucht

1. `.github/workflows/waechter.yml`: eine Kopie von `Vorlage-Modul.yml`, unverändert.
2. `.github/waechter.json`: die Liste seiner Werkzeuge und Bibliotheken, Aufbau unten.
3. Optional als Secret im Modul-Repository:
   - `TEAMS_WEBHOOK` für die Teams-Nachricht,
   - `NVD_API_KEY` für schnellere Abfragen bei der NVD.

Die Teams-Nachricht kommt nur aus diesem Lauf auf GitHub, nie von einer installierten App.

## Aufbau von `waechter.json`

```json
{
  "modul": "Name des Moduls",
  "werkzeugpaket": "werkzeuge.json",
  "pythonPakete": "requirements.txt",
  "eintraege": [
    { "id": "ghostscript",
      "neueste": { "art": "github", "repo": "ArtifexSoftware/ghostpdl-downloads", "tag": "^gs(\\d{2})(\\d{2})(\\d)$", "format": "{1}.{2:00}.{3}" },
      "cpe": "cpe:2.3:a:artifex:ghostscript" },
    { "id": "webview2", "version": "1.0.2903.40",
      "neueste": { "art": "nuget", "paket": "Microsoft.Web.WebView2" },
      "osv": { "oekosystem": "NuGet", "name": "Microsoft.Web.WebView2" } }
  ]
}
```

| Feld | Bedeutung |
|---|---|
| `modul` | Name im Issue und in der Teams-Nachricht |
| `werkzeugpaket` | optional. Die Versionen kommen dann aus dieser Datei (`dateien[].id`, `version`). Jede Datei darin braucht einen Eintrag. |
| `pythonPakete` | optional. Alle Pakete darin werden gegen osv.dev geprüft. |
| `eintraege[].id` | Kennung, bei einem Werkzeugpaket dieselbe wie dort |
| `version` | feste Version, wenn sie nicht aus dem Werkzeugpaket kommt |
| `pythonPaket` | Version aus `pythonPakete` übernehmen |
| `neueste.art` | `github` (Releases), `githubCommit` (feste Stände einer Datei), `pypi`, `nuget`, `npm`, `endoflife` |
| `cpe` | CPE-Name für die NVD (Programme ohne Paketquelle), `cpeTeile` kürzt die Version |
| `osv` | Ökosystem und Name bei osv.dev |
| `ignorieren` | geprüfte Fehlalarme: `[{ "id": "CVE-…", "grund": "…" }]` |
| `download` | Muster `…{version}…`, prüft, ob es die neue Version als fertiges Paket gibt |

Prüfen ohne Netz (Zielmaschine Entwicklungs-PC), vom Kern-Klon aus:

```
powershell.exe -File .github\waechter\Waechter.ps1 -Wurzel ..\<Modul-Klon> -Pruefen
```
