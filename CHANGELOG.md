# Changelog

## v0.1.6

### 🔐 Verwalter
- **Erster Verwalter bei der Installation:** Wer HU-Amt mit `Pull.ps1` neu installiert, ist automatisch Verwalter.
  - Bisher war nach einer neuen Installation niemand Verwalter, auch der Reiter zum Eintragen fehlte.
  - Ein Update ändert die Verwalter-Liste nie.
- **Frage beim ersten Start:** Gibt es noch keinen Verwalter, fragt HU-Amt einmal, ob du Verwalter werden willst.
  - Wer Nein sagt, wird nicht mehr gefragt. Sobald jemand eingetragen ist, kommt die Frage für niemanden mehr.

### 📖 Hilfe
- **Einrichten ohne Umweg:** Anleitung, Einstieg und README beschreiben den ersten Verwalter neu.

### ✅ Qualität
- 111 Tests.

## v0.1.5

### 💾 Speichern
- **Speichern auf Netzfreigaben repariert:** Auf manchen Freigaben kam beim Speichern „Der Zugriff auf den Pfad wurde verweigert“, obwohl der Benutzer Vollzugriff hat.
  - Betroffen waren alle Daten, die überschrieben werden, z. B. die Einstellungen im Reiter „🚗 Reisen“.
  - HU-Amt überschreibt die Datei jetzt auf normalem Weg, wenn die Freigabe das Ersetzen nicht erlaubt.

### ✅ Qualität
- 108 Tests.

## v0.1.4

### 📖 Hilfe
- **Neues Kapitel Reiseabrechnung:** Die Anleitung beschreibt den Reiter „🚗 Reisen“: einrichten, Reise erfassen, abrechnen und einreichen.
  - Einstieg und Funktionen zeigen das Modul ebenfalls.
- **Telefonrechnungen in der Kostenzuordnung:** Die Anleitung erklärt den Weg über die Texterkennung, die Zuordnung über die Rufnummer und was bei einem unlesbaren Betrag passiert.
  - Dazu: Der Excel-Import ersetzt nur die Kostenarten der Datei, die Spalte im Export heißt bei Telefon „Rufnummer“.

### 🎨 Oberfläche
- **Knöpfe ohne eigene Farbe gut lesbar:** Setzt ein Modul an einem Knopf keine Farbe, ist er jetzt dunkelgrau statt hellgrau mit weißer Schrift.

### ✅ Qualität
- 108 Tests.

## v0.1.3

### 📦 Installation
- **Installation in den eigenen Ordner:** Liegt `Pull.ps1` allein in einem neuen Ordner, kommt HU-Amt genau dorthin, auch auf eine Netzfreigabe.
  - Bisher landete die Erstinstallation immer in `Desktop\HU-Amt` des Benutzers.
  - In Desktop, Downloads, Dokumente und auf der Laufwerkswurzel bleibt es bei `Desktop\HU-Amt`.
- **Starten per Rechtsklick:** Rechtsklick auf `Pull.ps1` › „Mit PowerShell ausführen“ reicht, kein Befehl mehr nötig.
  - Anleitung, Einstieg und README nennen diesen Weg. Der bisherige Befehl ohne vollen Pfad fand die Datei nicht.

### ✅ Qualität
- 108 Tests.

## v0.1.2

### 🧩 Module
- **Alte Protokolle der Module werden gelöscht:** Beim Start räumt die App die Protokolle der Module nach derselben Frist auf wie ihre eigenen, nach 30 Tagen.
- **Reservierte Kennungen:** Ein Modul darf nicht `Postfach` oder `Wiederherstellung` heißen, weil die App diese Ordner selbst nutzt.

### 💾 Daten
- **Kopien vor dem Speichern:** Beim Aufräumen alter Kopien bleiben fremde Dateien im Ordner `_Kopien` stehen, auch wenn sie ähnlich heißen.

### 🔄 Update
- **Windows-Namen mit Punkt am Ende:** Das Update findet den gespeicherten Lese-Token jetzt auch bei solchen Benutzernamen.

### 📖 Hilfe
- **Neues Kapitel „Was passiert mit deinen Dateien“** in der Anleitung: wo was liegt, was mit Originalen passiert, wie lange etwas bleibt, was Löschen oder Umbenennen im Explorer bewirkt und was in der Sicherung ist.
- **Modul Kostenzuordnung:** Anleitung, Einstieg und Funktionen beschreiben den neuen Namen und die Kostenarten.

### ✅ Qualität
- 104 Tests.

## v0.1.1

### 🔄 Update
- **Release signieren friert nicht mehr ein:** Die Abfrage des GitHub-Tokens erscheint jetzt in einem eigenen Fenster.
  - Unter PowerShell 7 wartete die App vorher auf eine Eingabe im unsichtbaren Konsolenfenster und reagierte nicht mehr.
  - Gilt auch für „Release freigeben“ und das Hinterlegen des Lese-Tokens für Updates.

## v0.1.0

### 🧩 Module
- **Dateien zwischen Modulen weitergeben:** Ein Modul gibt Dateien an ein anderes weiter, z. B. die Texterkennung fertige Rechnungen an die Energiegemeinschaft.
  - Die Datei kommt als Kopie in ein Postfach im gemeinsamen Datenordner. So kommt sie auch an, wenn die Erkennung auf einem anderen PC läuft.
  - Am Reiter steht, wie viele Dateien warten, z. B. „Energiegemeinschaft (3)“. Die Zahl wird jede Minute aktualisiert.
  - Das Ziel-Modul nimmt nur Dateien mit seinen Endungen an.
  - Ob und wohin weitergegeben wird, legt der Anwender im Modul fest. Am Inhalt wird nichts erraten.
- **Wächter für externe Werkzeuge:** Werkzeuge und Bibliotheken der Module werden jede Woche auf neue Versionen, bekannte Lücken und Ende der Pflege geprüft.
  - Meldung als Issue im Repository des Moduls und auf Wunsch in Teams.

### 🪟 Hauptfenster
- **Reiter verschieben:** Reiter lassen sich wieder mit der Maus an eine andere Stelle ziehen. Die Reihenfolge bleibt beim nächsten Start erhalten.
  - Rechtsklick auf einen Reiter stellt die Reihenfolge zurück.

## v0.0.2

### 🧩 Module
- **Module per Datei eintragen:** Eine Modul-Datei (`*.humodule`) mit Repositories und Nur-Lese-Token ins Fenster ziehen, und alle Module sind auf einmal hinterlegt.
  - Nur für Verwalter; vorher kommt eine Rückfrage mit der Liste der Repositories.
  - Jeder Token wird vor dem Speichern geprüft, das Ergebnis erscheint je Repository.
  - Sind alle Einträge gespeichert, wird die Datei gelöscht. Sonst bleibt sie, und die Meldung nennt, was fehlt.
  - Geht jederzeit, auch später mit einer Datei nur für ein neues Modul.
  - Alternativ im Fenster „Module“ über „Datei einlesen …“.

### 🪟 Hauptfenster
- **Erster Start:** Die App fragt nach einem Symbol auf dem Desktop und zeigt auf Wunsch, wie sie an die Taskleiste kommt.

### 🚧 Grenzen
- **Taskleiste:** Windows erlaubt Programmen nicht, sich selbst anzuheften. Das Anheften macht der Anwender mit Rechtsklick auf das Symbol.
- **Token je Windows-Benutzer:** Jeder Verwalter, der Module aktualisiert, zieht die Modul-Datei selbst ein.

## v0.0.1

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
- **Austausch mit anderen Programmen:** Module geben Tabellen als CSV nach einem Profil aus (Trennzeichen, Kodierung auch Windows-1252, Dezimalzeichen, Datumsformat, Spalten).
  - Ein neues Zielformat ist nur eine Profil-Datei in `Data\<Modul>\Profile`.
  - Jedes Modul hat einen Übergabeordner (Standard `Data\<Modul>\Ausgang`, auch eine Freigabe). Dateien erscheinen erst fertig geschrieben, mit Zeitstempel im Namen und ohne zu überschreiben.

### ⚙️ Einstellungen
- **Einstellungsfenster** mit den Reitern Allgemein, Sicherung, Update, Module und Verwalter.
  - Allgemein: Verknüpfungen und Starter, Reiter beim Start, Oberflächengröße, Startanimation.
  - Update: Update-Prüfung beim Start für alle; Kanal, Signaturpflicht und Vorversion nur für Verwalter.
  - Nur für Verwalter: PowerShell beim Start, Protokoll-Stufe, Sicherung, geladene Module, Verwalter-Liste.
- **Verwalter-Ansicht:** Protokoll und Verwaltung sehen nur Verwalter; alle anderen bekommen eine aufgeräumte Oberfläche.
  - Verwalter lassen sich in den Einstellungen eintragen und entfernen (mit Rückfrage, vorher eine Kopie der Liste).

### ❓ Hilfe
- **F1 und „?“** öffnen die Anleitung zum aktuellen Reiter im Browser, Rechtsklick auf „?“ die Übersicht.
  - Ohne eigene Hilfe eines Moduls erscheint eine kurze Grundhilfe.

### 🔄 Update
- **Goldener Knopf „Update“:** Beim Start sucht die App im Hintergrund nach einer neuen Version; ist eine da, wird der Knopf gold. Ein Klick installiert sie.
  - Jede Version wird vor dem Ersetzen vollständig geladen und mit Prüfsumme und Signatur geprüft; bei einem Fehler bleibt alles unverändert.
- **Gemeinsamer Netzordner:** Das Update startet erst, wenn alle anderen die App geschlossen haben. Die App zeigt, wer sie noch offen hat.
- **Daten bleiben erhalten:** Der Ordner `Data` und lokale Einstellungen werden beim Update nie angefasst.
- **Verwaltung im Rechtsklick-Menü:** andere Version, Kanal Test, private Module, GitHub-Token; Release signieren und freigeben nur mit Herausgeber-Zertifikat.
  - Verwalter ist, wer in `Data\Verwalter.txt` steht.
- **Info-Fenster** mit Version, Kanal und Prüfung der Installation.

### 💾 Sicherung
- **Verschlüsselte Sicherung aller Daten** (`.husicherung`) in einen Ordner, am besten auf ein Netzlaufwerk.
  - Automatisch täglich oder wöchentlich; im gemeinsamen Netzordner sichert nur einer.
  - Ohne Kennwort lässt sich eine Sicherung nicht öffnen. Das Kennwort für die Automatik bleibt verschlüsselt beim eigenen Windows-Benutzer.
  - Jede Sicherung wird nach dem Schreiben zurückgelesen und geprüft.
- **Einstellen im Reiter Sicherung:** Ordner, Kennwort, Automatik, Aufbewahren, Jetzt sichern, Sicherung prüfen, Wiederherstellen.
- **Wiederherstellen nach Bereichen:** je Modul, Benutzer, Verwalter, Allgemein. Der vorherige Stand bleibt in `Data\Wiederherstellung`.
  - Erst, wenn niemand sonst die App offen hat; danach startet die App neu.
- **Diskette 💾 in der Kopfzeile** für Verwalter: Stand im Tooltip, während der Sicherung mit Prozent, bei einem Fehler mit „!“. Klick öffnet den Reiter Sicherung.

### 🛟 Support
- **Support anfragen:** Beschreibung, Protokolle und Einstellungen als ZIP auf dem Desktop, zum Ansehen oder per E-Mail.
  - Anonymisiert: E-Mail-Adressen, IDs, IP-Adressen, Programmordner, Benutzer und PC. Daten der Module, Kennwörter und Token sind nie dabei.
  - Verwalter können zusätzlich ein GitHub-Issue vorbereiten.

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

### ✅ Qualität
- 78 Tests.
