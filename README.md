<h1 align="center"><img src="Assets/logo64.png" width="44" alt="" align="absmiddle"/> HU-Amt</h1>

<p align="center"><b>Werkzeuge für die Gemeindeverwaltung – aus einem Fenster</b><br>
Jeder Reiter ist ein Modul: mitgeliefert oder signiert aus einem eigenen Repository. Dazu Update per Knopf, verschlüsselte Sicherung aller Einstellungen, Hilfe mit F1 und dunkles Farbschema.<br>
Läuft ohne Installation aus einem Ordner am PC oder auf einer Netzfreigabe, für mehrere Bedienstete gleichzeitig.</p>

<p align="center">
  <a href="https://github.com/ChiliApple/HU-Amt/releases/latest"><img src="https://img.shields.io/github/v/release/ChiliApple/HU-Amt?include_prereleases&label=Version&color=b9a88a" alt="Version"></a>
  <img src="https://img.shields.io/badge/PowerShell-5.1%20%7C%207-5391FE?logo=powershell&logoColor=white" alt="PowerShell 5.1 | 7">
  <img src="https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6" alt="Windows 10 | 11">
  <img src="https://img.shields.io/badge/Oberfl%C3%A4che-WPF-8839ef" alt="WPF">
  <img src="https://img.shields.io/badge/Updates-signiert-2E7D32" alt="Updates signiert">
  <a href="LICENSE"><img src="https://img.shields.io/badge/Lizenz-Nutzung%20frei-orange" alt="Lizenz"></a>
</p>

<p align="center">
  <a href="https://chiliapple.github.io/HU-Amt/Docs/Grundhilfe.html"><b>Hilfe</b></a> ·
  <a href="#installation">Installation</a> ·
  <a href="#auf-einer-netzfreigabe">Netzfreigabe</a> ·
  <a href="MODULE-DEVELOPMENT.md">Module entwickeln</a> ·
  <a href="CHANGELOG.md">Änderungen</a> ·
  <a href="LICENSE">Lizenz</a>
</p>

<p align="center">
  <a href="Docs/screenshot-energiegemeinschaft.png"><img src="Docs/screenshot-energiegemeinschaft.png" width="380" alt="Energiegemeinschaft"/></a>
  <a href="Docs/screenshot-uebersicht.png"><img src="Docs/screenshot-uebersicht.png" width="380" alt="Übersicht"/></a>
  <br>
  <a href="Docs/screenshot-pdf-werkzeuge.png"><img src="Docs/screenshot-pdf-werkzeuge.png" width="380" alt="PDF-Werkzeuge"/></a>
  <a href="Docs/screenshot-ideen.png"><img src="Docs/screenshot-ideen.png" width="380" alt="Ideen und Wünsche"/></a>
  <br><sub>Energiegemeinschaft · Übersicht · PDF-Werkzeuge · Ideen &amp; Wünsche – mit Testdaten, zum Vergrößern anklicken</sub>
</p>

---

## Was es kann

- **Module als Reiter:** Jedes Werkzeug ist ein eigener Reiter. Die Reiter lassen sich mit der Maus verschieben, die Reihenfolge bleibt gespeichert.
- **Dateien hineinziehen:** PDF- oder Excel-Dateien aufs Fenster ziehen, das passende Modul übernimmt sie. Module können Dateien auch an andere Module weitergeben.
- **Update per Knopf:** Programm und Module kommen als signierte Releases von GitHub. Jede Datei wird vor dem Ersetzen geprüft, bei einer Abweichung bleibt alles, wie es war.
- **Sicherung:** Einstellungen, Daten und Module werden verschlüsselt gesichert und lassen sich wiederherstellen.
- **Beim nächsten Start wie beim Schließen:** Fenster, Reiter und Einstellungen jedes Benutzers bleiben erhalten.
- **Ansichten:** Verwalter richtet ein, Bedienstete arbeiten.

## Installation

1. `Pull.ps1` in einen leeren Ordner legen, z. B. `C:\Tools\HU-Amt`, und ausführen:
   `powershell -ExecutionPolicy Bypass -File Pull.ps1`
2. **Start-HU-Amt.cmd** doppelklicken.
3. Beim ersten Start fragt das Programm, ob es eine Verknüpfung auf den Desktop oder in die Taskleiste legen soll.

Voraussetzungen: Windows 10 oder 11, Windows PowerShell 5.1 (vorhanden) oder PowerShell 7.

## Auf einer Netzfreigabe

- Den Ordner einmal auf die Freigabe legen. Jeder Benutzer startet `Start-HU-Amt.cmd` von dort und bekommt beim ersten Start die Frage nach der Verknüpfung.
- Daten liegen gemeinsam im Ordner `Data\`, Einstellungen und Zugangsdaten je Windows-Benutzer.
- Ein Update ersetzt nichts, solange jemand anderer das Programm offen hat.

## Module

- **Mitgeliefert** kommen sie mit dem Programm-Update.
- **Aus einem eigenen Repository** werden sie in der Modulverwaltung eingetragen, auch mehrere auf einmal über eine Zugangsdatei. Jede Fassung muss signiert sein.
- Wie man ein Modul baut: [MODULE-DEVELOPMENT.md](MODULE-DEVELOPMENT.md).

---

**Lizenz:** kostenlose Nutzung erlaubt, eigene Module und Einstellungen ausdrücklich auch – Details in [LICENSE](LICENSE).
