# Regeln für KI-gestützte Mitarbeit an Modulen

Diese Datei bekommt jede KI, die an einem Modul von HU-Amt mitarbeitet, als feste Vorgabe. Sie gilt ohne Ausnahme. Eine Kopie liegt unverändert in jedem Modul-Repository unter `.github/dev/KI-REGELN.md`.

## Arbeitsweise

- Sprache Deutsch, Du-Form, kurz und ohne Füllsätze, kein Gendern. Englische Fachbegriffe sind ok.
- Lesen und Analysieren selbstständig und vollständig. Änderungen schrittweise: ein Schritt, testen, dann der nächste.
- Erst die Ursache beweisen (Code lesen, nachstellen), dann Code ändern. Jeden Fix testen, bevor er vorgestellt wird.
- Nur die verlangte Änderung. Nichts Funktionierendes nebenbei umbauen.
- Unsicheres ausdrücklich als "unsicher" kennzeichnen. Behauptungen belegen (Datei:Zeile). Fehler direkt zugeben.
- Bei jedem Befehl die Zielmaschine nennen. Immer komplette, lauffähige Änderungen liefern.

## Technik

- PowerShell 5.1 ist Zielversion, alles muss zusätzlich unter PowerShell 7 laufen. Keine PS7-only-Syntax (kein ?:, ??, &&, ||).
- Set-StrictMode -Version 1.0 wie die App: jede Variable vor Gebrauch belegen.
- Dateien UTF-8 mit BOM und CRLF (.yml und .py ohne BOM, LF). Vor jedem Commit .github\dev\Format-Dateien.ps1.
- Keine typografischen Anführungszeichen („ “ ”) innerhalb von "..."-Zeichenketten: PowerShell behandelt sie wie " und schneidet den Text ab. Stattdessen $([char]0x201E) bzw. $([char]0x201C) oder einfache Anführungszeichen.
- $null nie an .NET-Parameter vom Typ string übergeben (wird zu ''), sondern [NullString]::Value.
- Nach jeder Dateibearbeitung Vollständigkeit prüfen (Dateiende, Klammern, schließende Tags).

## Aufbau

- Der Kern liefert alles Gemeinsame, das Modul enthält nur Fachlogik.
- Nie selbst bauen, was der Kern hat: Dialoge (Show-HUMessage, Confirm-HU, Read-HUText, Read-HUChoice, Read-HUNumberInput), Kalender, Splitter, Hintergrund-Aufträge (Start-HUBackgroundJob), Datenablage, Sperren, Verlauf, Anhänge, Zahlen und Datum, PDF, Excel/CSV, Kontakte.
- Outlook und Word nur über Core\HU.Office.psm1. Kein eigenes COM im Modul.
- Fehlt dem Kern etwas: melden, nicht im Modul nachbauen.
- Schnittstellen und Regeln stehen in MODULE-DEVELOPMENT.md im Kern-Repository. Diese Datei zuerst lesen.

## Daten

- Ablage nur im Datenordner des Moduls (Get-HUAddonDataDir bzw. $script:DataRoot), nie selbst aus Programmordner und "Data" zusammengesetzt.
- Eine JSON-Datei je Vorgang. Schreiben nur über die Get-/Save-Funktionen des Moduls, Write-HUFileAtomic und die Sperren des Kerns. Kein Lesen oder Schreiben von Datensätzen im Fenster-Code.
- Jeder Datensatz hat eine feste ID, die sich nie ändert (nicht aus Name, Nummer oder Datum gebildet). Verweise nur per ID, ein Name darf als Anzeigekopie daneben stehen.
- Jede gespeicherte Datei hat ein Feld "schema" (Formatversion). Fehlt es, gilt 0. Alte Stände müssen weiter lesbar sein.
- Datum als ISO-Text, Beträge und Mengen als JSON-Zahl, keine Einheiten im Wert.
- Formatänderungen nur mit automatischem Umbau beim Start (vorher Sicherungskopie), nie mit Handarbeit vor Ort.
- Keine Geheimnisse im Klartext in gemeinsamen Dateien.

## Demo und Testmodus

- Beispieldaten liegen im Ordner Testdaten\ des Moduls, aufgebaut wie der Datenordner, nur erfundene Inhalte.
- Test-HUTestMode: Outlook, Word und Mail sind gesperrt. Test-HUDemoMode: Demo für alle, bei jedem Start frisch; Test-HUTestMode ist dann ebenfalls wahr. Prüflisten und Verwalter-Testwerkzeuge in der Demo ausblenden. Aufrufe mit Get-Command absichern (ältere Kerne).
- In Demo und Test nie echte Ordner überwachen, verschieben oder in echte Benutzerordner schreiben.

## Oberfläche

- Dunkles Farbschema wie der Kern, Farben aus dem Theme des Kerns.
- Jedes Steuerelement hat einen Tooltip.
- Tabellen: Spalten etwas breiter als der Inhalt, nie abgeschnitten. Fenster so breit wie die Spalten.
- Größenveränderbare Fenster haben rechts unten einen sichtbaren Anfasser (ResizeMode CanResizeWithGrip).
- Texte aus Sicht der Anwender (Bedienstete ohne IT-Wissen), mit Umlauten, ohne interne Begriffe.
- Funktionen tragen das Präfix des Moduls (z. B. ...-HUBu... bei Buchungen).
- Hilfe.html des Moduls standardmäßig dunkel, jede Aussage muss zum Code passen.

## Tests

- Für jede Änderung ein Test im Testsystem des Moduls (.github\tests), mit Gegenprobe (der Test schlägt gegen den alten Stand fehl).
- Tests öffnen keine sichtbaren Fenster.
- .github\tests\Invoke-CITests.ps1 unter PS 5.1 und 7 grün, nacheinander, nicht parallel.
- Starttest: .github\dev\Test-Start.ps1 (auch -TestMode), Ablauftest des Moduls.
- Test-App nur über .github\dev\Deploy-Dev.ps1, nie während sie läuft.

## Git und Release

- Arbeiten nur auf einem Zweig release/x.y.z-test, Pull Request auf main. Nie direkt auf main, kein Force-Push, keine umgeschriebene Historie.
- Nicht mergen, nicht signieren, keine Releases anlegen. Das macht der Maintainer.
- Version in addon.json: kleine Schritte als Patch (1.0.1, 1.0.2 ...), ohne führende Nullen.
- CHANGELOG: je Version "## vX.Y.Z", gegliedert nach Bereich der App mit "### <Emoji> <Bereich>", jeder Punkt "- **Kurzer Titel:** Aussage", aus Sicht der Anwender. Am Ende nur wenn nötig: "### ⚠️ Nach dem Update", "### 🚧 Grenzen", "### ✅ Qualität" (nur Testzahl).
- Commit-Nachrichten Deutsch, sachlich, ohne Hinweis auf KI oder Werkzeuge.

## Was nie ins Repository darf

- Echte Daten: Namen, Adressen, Rechnungen, Buchungen, Excel-Inhalte. Nur erfundene Testdaten.
- Gemeinde-, Schul- und Personennamen, Domains, interne IP-Adressen.
- Geheimnisse: Tokens, Kennwörter, PINs, Webhook-Adressen, Zertifikatsschlüssel.
- Namen von KI-Werkzeugen und Hinweise wie "Generated with ...".
- Laufergebnisse und Protokolle (gehören in .gitignore).
