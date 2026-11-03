# Changelog

## v0.0.1 (in Arbeit)

### 🪟 Hauptfenster
- **Grundgerüst:** Kopfzeile mit Logo, Name und Version, Reiterleiste und Protokoll im dunklen Farbschema.
  - Ein Reiter „Start“ erscheint, solange kein Modul installiert ist.
- **Startbildschirm:** erscheint sofort beim Start und verabschiedet sich mit dem Saugroboter.
- **Fenster merkt sich Größe und Position**, je Windows-Benutzer.
- **Strg + Mausrad** vergrößert oder verkleinert die Oberfläche, Strg + 0 stellt zurück.

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
- Einstellungen, Support, Hilfe und Update sind noch ohne Funktion.

### ✅ Qualität
- 23 Tests.
