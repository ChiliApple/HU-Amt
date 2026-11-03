# Changelog

## v0.0.1 (in Arbeit)

### 🪟 Hauptfenster
- **Grundgerüst:** Kopfzeile mit Logo, Name und Version, Reiterleiste und Protokoll im dunklen Farbschema.
  - Ein Reiter „Start“ erscheint, solange kein Modul installiert ist.
- **Startbildschirm:** erscheint sofort beim Start und verabschiedet sich mit dem Saugroboter.
- **Fenster merkt sich Größe und Position**, je Windows-Benutzer.
- **Strg + Mausrad** vergrößert oder verkleinert die Oberfläche, Strg + 0 stellt zurück.
- **Protokoll in der Höhe verstellbar:** am Griff über dem Protokoll ziehen, die Höhe bleibt gemerkt.

### 🧩 Module
- **Ein Reiter je Modul:** Module kommen mit dem Programm, aus einem privaten Repository (nur signiert) oder zum Entwickeln aus `AddonsDev`.
  - Die Reihenfolge der Reiter legt das Modul fest; Reiter lassen sich mit der Maus verschieben.
- **Dateien ins Fenster ziehen:** Die Dateien gehen an den aktiven Reiter. Passt er nicht, aber genau ein anderer Reiter, wechselt die App dorthin.
- **Daten je Modul** liegen gemeinsam für alle in `Data\<Modul>`.

### 🔄 Update
- **Goldener Knopf „Update“:** Beim Start sucht die App im Hintergrund nach einer neuen Version; ist eine da, wird der Knopf gold. Ein Klick installiert sie.
  - Jede Version wird vor dem Ersetzen vollständig geladen und mit Prüfsumme und Signatur geprüft; bei einem Fehler bleibt alles unverändert.
- **Gemeinsamer Netzordner:** Das Update startet erst, wenn alle anderen die App geschlossen haben. Die App zeigt, wer sie noch offen hat.
- **Daten bleiben erhalten:** Der Ordner `Data` und lokale Einstellungen werden beim Update nie angefasst.
- **Verwaltung im Rechtsklick-Menü:** andere Version, Kanal Test, private Module, GitHub-Token; Release signieren und freigeben nur mit Herausgeber-Zertifikat.
  - Verwalter ist, wer in `Data\Verwalter.txt` steht.
- **Info-Fenster** mit Version, Kanal und Prüfung der Installation.

### 🖥️ Start
- **Desktop-Symbol:** Beim ersten Start fragt die App, ob ein Symbol auf den Desktop soll.
- **PowerShell 7 oder 5.1:** Die App startet mit PowerShell 7, wenn installiert, sonst mit Windows PowerShell 5.1.

### 📁 Daten
- **Ordner `Data`:** Persönliches (Fenster, Protokoll) liegt getrennt je Benutzer in `Data\Benutzer\<Name>`.
- **Gemeinsamer Netzordner:** Mehrere Personen können dieselbe Installation auf einer Freigabe nutzen.
  - Jede geöffnete App meldet sich in `Data\Laufend` an; nach einem Absturz gibt Windows die Anmeldung von selbst frei.
  - Gemeinsame Listen lassen sich nur von einer Person zugleich bearbeiten, die anderen sehen, wer gerade bearbeitet.
  - Vor dem Speichern einer gemeinsamen Liste entsteht eine Kopie mit Zeitstempel.
- **Klare Meldung**, wenn der Datenordner nicht beschreibbar ist.

### 🚧 Grenzen
- Einstellungen, Support und Hilfe sind noch ohne Funktion.

### ✅ Qualität
- 46 Tests.
