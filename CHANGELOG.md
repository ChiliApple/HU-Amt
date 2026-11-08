# Changelog

## v0.1.14

### 💾 Daten auf der Freigabe
- **Datenordner frei wählbar:** Verwalter legen die Daten aller Module unter Einstellungen › Verwalter › „Datenordner ändern …“ auf eine andere Freigabe, am besten als UNC-Pfad. Das Programm selbst kann bleiben, wo es ist.
  - Ist der neue Ordner leer, kopiert HU-Amt alle Daten dorthin, prüft jede Datei und stellt erst dann um. Klappt das nicht vollständig, bleibt alles beim Alten.
  - Liegen dort schon Daten von HU-Amt, werden sie ohne Kopieren übernommen. Ordner mit anderen Dateien lehnt HU-Amt ab.
  - Vorher müssen alle anderen HU-Amt-Fenster geschlossen sein, die Meldung nennt, wer es noch offen hat.
  - Der alte Ordner bleibt als Stand erhalten und bekommt den Hinweis UMGEZOGEN.txt.
  - Danach startet HU-Amt neu, auf den anderen PCs gilt der neue Ordner beim nächsten Start.
- **Kein stilles Ausweichen:** Ist der eingestellte Ordner nicht erreichbar, startet HU-Amt mit einer klaren Meldung nicht, statt still einen zweiten, leeren Datenbestand anzulegen.

### ❓ Hilfe
- **Wo liegen meine Daten?** Neue Hilfeseite mit Grafik: für jedes Modul, was mit hineingezogenen Dateien passiert, wo Ergebnisse landen, was gesichert wird und was ins Internet geht. Erreichbar über die Hilfe (F1) und die Anleitung.

### 🧪 Testmodus
- **Beispieldaten:** Bringt ein Modul Beispieldaten mit, legt HU-Amt sie beim ersten Start im Testmodus und nach „Test zurücksetzen“ in den Testbestand. So lässt sich jedes Modul gleich ausprobieren. Echte Daten bleiben unberührt.

### 💾 Sicherung
- **Warnung, wenn die Sicherung fehlt:** Ist keine Sicherung eingerichtet, die Automatik aus oder die letzte erfolgreiche Sicherung zu alt (täglich: 2 Tage, wöchentlich: 9 Tage), wird die Diskette oben orange. Verwalter bekommen einmal am Tag einen Hinweis.

### 🔐 Rechte
- **Nur Verwalter:** Den Datenordner ändern nur Verwalter, im Testmodus gar nicht.
- **Anleitung für die IT:** `RECHTE-IT.md` im Programmordner beschreibt Ordner, Gruppen und Rechte, mit fertigen Befehlen für den Server.
- **Rechteprüfung erweitert:** Die Warnzeile unter Einstellungen › Verwalter meldet auch, wenn alle Benutzer die Einstellung des Datenordners ändern dürfen.

### ⚠️ Nach dem Update
- Die Einstellung liegt in `Config\datenort.json` und gilt für alle, die diese Installation starten. `Config` braucht deshalb dieselben Rechte wie das Programm: für die Benutzer nur Lesen.

### ✅ Qualität
- 484 Tests, alle im strengen Modus wie die App.

## v0.1.13

### 🔐 Rechte
- **Einstellungen öffnen sofort:** Die Prüfung der Rechte auf die Verwalter-Liste läuft jetzt im Hintergrund. Bis das Ergebnis da ist, steht „wird geprüft …“. Auch eine langsame Freigabe hält Fenster und Programmstart nicht mehr auf.

### ✅ Qualität
- 467 Tests, alle im strengen Modus wie die App.

## v0.1.12

### 📤 Datei-Austausch
- **Beträge und Datum richtig in Listen für andere Programme:** Zahlen und Datumswerte aus Texten werden jetzt österreichisch gelesen.
  - „1,5“ bleibt 1,5 (bisher 15), „12,50“ bleibt 12,50, „1.234,56“ klappt jetzt.
  - „05.11.2026“ bleibt der 5. November (bisher 11. Mai), „13.11.2026“ geht jetzt.
  - Was weder eine Zahl noch ein Datum ist, bricht mit Zeile und Spalte ab, statt still falsch zu landen.

### 💾 Daten auf der Freigabe
- **Kurz weg heißt nicht leer:** Ist die Freigabe kurz nicht erreichbar, zeigen Module weiter den letzten Stand und melden „Ordner nicht erreichbar“, statt eine leere Liste zu zeigen.
- **Dateinamen wie „CON.txt“:** Namen, die Windows für Geräte reserviert, bekommen auch mit Endung einen Unterstrich davor.

### 🔐 Rechte
- **Rechte der Verwalter-Liste:** Unter Einstellungen › Verwalter zeigt HU-Amt eine Warnzeile, wenn alle Benutzer die Verwalter-Liste ändern oder im Ordner Data alles löschen dürfen. Details stehen im Tooltip.
  - Beim Start steht die Warnung nur im Protokoll, normale Benutzer sehen nichts.
  - Die Freigaberechte selbst kann HU-Amt nicht sehen, das steht auch im Hinweis.
  - Empfohlene Rechte für Freigabe, Data und Verwalter.txt stehen in der Anleitung unter „Rechte auf der Freigabe“.

### 📨 Outlook und Word
- **Word-Vorlagen ohne Makros:** Vorlagen müssen .docx oder .dotx sein, sonst kommt eine klare Meldung. Makros bleiben beim Füllen immer aus.
- **Kontaktsuche mit Sonderzeichen:** Suchtexte mit „[“ oder „*“ finden jetzt, was drinsteht, statt einen Fehler zu werfen.

### 👥 Kontakte
- **vCard mit Foto:** Kontakte nach einem Foto gehen beim Einlesen nicht mehr verloren.

### 📊 Excel und CSV
- **Große Excel-Dateien:** Eine einzelne Zeile weit unten im Blatt bremst das Einlesen nicht mehr aus.

### 🎨 Oberfläche
- **Kein Hängen beim Abbrechen:** Bricht HU-Amt einen Auftrag ab, der nicht antwortet, bleibt das Fenster bedienbar.
- **Zweistellige Jahre:** „1.1.35“ ist auf jedem PC der 1.1.2035, unabhängig von der Windows-Einstellung.
- **Rückfrage bei 1.500:** Module können in Betragsfeldern nachfragen, ob 1.500 (also 1500) oder 1,5 gemeint ist.
- **Standardknopf in Auswahlfenstern:** Module können festlegen, welcher Knopf auf Enter reagiert.
- **Hinweise öffnen zuverlässig:** Einige Hinweise brachen wegen eines falschen Symbols ab.

### 🧪 Testmodus
- **Geführter Test:** Die Trennlinie zwischen Liste und Schritt lässt sich verschieben, Größe und Lage bleiben erhalten, Esc schließt.
  - Sind noch Schritte unbestätigt, fragt das Fenster vor dem Schließen nach.
- **Verwalter im Testmodus:** Die Verwalter-Liste wird bei jedem Start des Testmodus neu übernommen.

### 🛟 Support
- **Netzpfade unkenntlich:** Im Support-Paket werden Server und Freigabe in Netzpfaden durch Platzhalter ersetzt.

### ✅ Qualität
- 472 Tests, alle im strengen Modus wie die App.

## v0.1.11

### 🧩 Module
- **Gemeinsame Bausteine im Hauptprogramm:** Was mehrere Module bisher jedes für sich gebaut haben, liefert jetzt HU-Amt selbst. Module werden dadurch einheitlicher und Fehler werden an einer Stelle behoben.
  - Datenablage auf der Freigabe mit Sperre, Warten und Schutz vor gleichzeitigem Speichern, Verlauf („wer hat wann was geändert“), Anhänge und gemeinsame Einstellungen.
  - Euro, Zahlen und Datum einheitlich für Österreich, unter Windows PowerShell und PowerShell 7 gleich.
  - Dunkler Kalender, Uhrzeit-Auswahl, Text- und Auswahl-Dialoge, Laden im Hintergrund.
  - PDF-Ausdrucke mit Tabellen über mehrere Seiten, Lesen von Excel- und CSV-Dateien ohne Excel, Aufruf externer Werkzeuge mit Zeitlimit.
- **Umstellung der Module folgt:** Die Module übernehmen die Bausteine mit ihren nächsten Testversionen. Bis dahin arbeiten sie unverändert weiter.

### ✅ Qualität
- 347 Tests, alle im strengen Modus wie die App.

## v0.1.10

### 📨 Outlook und Word
- **Absender für Mails aus Modulen:** Module können einstellen, von welchem Mailkonto ihre Mails kommen, auch von einem gemeinsamen Postfach („im Auftrag von“).
  - Ist nichts eingestellt und hat jemand mehrere Konten im Outlook, fragt HU-Amt einmal nach dem Konto. Auf Wunsch merkt es sich die Wahl je Benutzer.
  - Fehlt das eingestellte Konto auf einem PC, fragt HU-Amt nach, statt still ein anderes zu nehmen.
  - Die Mail wird weiterhin nur angezeigt. Im Feld „Von“ steht der Absender und lässt sich vor dem Senden ändern.
- **Outlook-Ordnerauswahl im Vordergrund:** Der Dialog „Ordner auswählen“ erscheint jetzt vor HU-Amt statt dahinter.

### 🎨 Oberfläche
- **Verschobene Bereiche bleiben erhalten:** HU-Amt merkt sich in jedem Modul-Reiter und jedem Fenster, wie weit du die Bereiche verschoben hast, und stellt sie beim nächsten Start wieder her.
  - Das gilt jetzt für alle Module, ohne dass ein Modul dafür etwas tun muss.

### 👥 Kontakte
- **Kontakte übernehmen:** Module können Kontakte aus dem eigenen Outlook oder aus vCard-Dateien (.vcf) übernehmen, z. B. für Mieter.
  - Suche nach Name, Firma, Mail oder Ort. vCard-Dateien lassen sich auch ins Fenster ziehen.
  - Übernommen wird nur, was du markierst.

### 🧪 Testmodus
- **Geführter Test öffnet zuverlässig:** In 0.1.9 brach das Öffnen des Testassistenten ab. Behoben; die Tests des Kerns laufen jetzt wie die App im strengen Modus.

### ⚠️ Nach dem Update
- Für das Senden „im Auftrag von“ eines gemeinsamen Postfachs braucht der Benutzer im Exchange das Recht dazu.

### ✅ Qualität
- 195 Tests, alle im strengen Modus wie die App.

## v0.1.9

### 🧪 Testmodus
- **Neue Module gefahrlos ausprobieren:** Verwalter starten unter Einstellungen › Verwalter › „Testmodus starten“ einen Testmodus.
  - HU-Amt startet neu und arbeitet mit einem eigenen Testbestand. Echte Daten bleiben unberührt.
  - Outlook, Word und Mails sind im Test gesperrt, die automatische Sicherung ruht.
  - Eine orange Leiste zeigt den Testmodus, „TEST“ steht im Fenstertitel.
  - „Test zurücksetzen“ löscht nur die Testdaten. „Testmodus beenden“ oder ein normaler Neustart führt zurück zu den echten Daten.
- **Geführter Test für Module:** Module können einen Testassistenten anbieten.
  - Je Schritt zeigt er, was zu tun ist und was herauskommen soll. Wer testet, bedient die echten Masken und bestätigt „Bestanden“ oder „Fehler“ mit Kommentar.
  - Was das Programm selbst prüfen kann, steht getrennt von der Bestätigung durch den Menschen.
  - Das Testprotokoll wird als CSV und HTML gespeichert, mit Version, Soll, Ist, Ergebnis und offenen Fehlern.

### 💾 Sicherung
- **Testdaten nicht in der Sicherung:** Der Testbestand wird nie mitgesichert.

### ✅ Qualität
- 168 Tests, dazu ein Starttest im Testmodus.

## v0.1.8

### 📨 Outlook und Word
- **Outlook startet zuverlässig:** Wird Outlook gerade erst gestartet oder beendet es sich noch, wartet HU-Amt bis zu 30 Sekunden und versucht es erneut, statt abzubrechen.
  - Bisher konnte dabei ein unsichtbares Outlook im Hintergrund weiterlaufen.
- **Outlook wird sauber beendet:** Hat HU-Amt Outlook selbst gestartet, wartet es nach dem Beenden kurz, bis Outlook wirklich zu ist.

### ✅ Qualität
- 147 Tests.

## v0.1.7

### 📨 Outlook und Word
- **Gemeinsame Grundlage für Module:** Module können jetzt Termine und Aufgaben in Outlook anlegen, Mails vorbereiten und Word-Vorlagen ausfüllen, ohne das selbst zu bauen.
  - Termine und Aufgaben landen im eigenen oder in einem gewählten gemeinsamen Kalender. Wird etwas geändert, ändert sich derselbe Eintrag, es entstehen keine doppelten.
  - Mails werden nur angezeigt, gesendet wird immer von Hand.
  - Word-Vorlagen bleiben unverändert. Heraus kommt ein ausgefülltes PDF oder eine Word-Datei.
  - Ohne klassisches Outlook gibt es eine Kalenderdatei (.ics) bzw. eine Mail über das Standard-Mailprogramm.
- **Kein Office-Programm bleibt offen:** Outlook wird nur beendet, wenn es vorher nicht lief und kein Fenster offen ist. Word läuft unsichtbar und wird immer geschlossen.
  - Antwortet Word zwei Minuten lang nicht, beendet HU-Amt genau diese unsichtbare Word-Instanz. Ein selbst geöffnetes Word bleibt unberührt.

### ⚠️ Nach dem Update
- Für Module, die Outlook oder Word nutzen, ist das klassische Outlook bzw. Word nötig. Das neue Outlook hat diese Schnittstelle nicht.

### ✅ Qualität
- 143 Tests.

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
