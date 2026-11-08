# HU-Amt: Ordner und Rechte auf dem Server

Anleitung für die IT. Liegt jeder Installation bei. Stand: HU-Amt 0.1.14.

HU-Amt ist ein Büroprogramm ohne Installation. Es läuft aus einem Ordner auf dem Server, alle Bediensteten starten es von dort. Das Programm selbst und seine Daten liegen in zwei getrennten Ordnern. Wer was darf, regeln nur die NTFS-Rechte. HU-Amt hat keine eigene Benutzerverwaltung.

## 1. Ordner

| Ordner (Netzwerkpfad) | Inhalt |
|---|---|
| `\\server\freigabe\HU-Amt` | Programm, wird über den Update-Knopf aktualisiert |
| `\\server\freigabe\HU-Amt-Daten` | Daten aller Module, Verwalter-Liste, persönliche Einstellungen |

Der Datenordner wird im Programm eingestellt (Einstellungen › Verwalter › „Datenordner ändern …“). Das macht ein Verwalter im Programm, die IT muss dafür nichts tun.

## 2. Gruppen im AD

| Gruppe (Vorschlag) | Mitglieder |
|---|---|
| `GG_HU-Amt-Benutzer` | alle, die HU-Amt benutzen |
| `GG_HU-Amt-Verwalter` | wer HU-Amt aktualisiert und verwaltet |

Nach dem Aufnehmen in eine Gruppe muss sich der Benutzer neu anmelden.

Die Verwalter-Gruppe regelt nur die Dateirechte. Wer im Programm Verwalter ist, steht zusätzlich in `HU-Amt-Daten\Verwalter.txt` und wird im Programm gepflegt. Beides sollte dieselben Personen enthalten.

## 3. Rechte

Liegen die Ordner unter einer Freigabe mit offenen Rechten, wird die Vererbung an beiden Ordnern **unterbrochen**, sonst erben sie diese Rechte.

| Ordner oder Datei | Benutzer | Verwalter | SYSTEM, Administratoren |
|---|---|---|---|
| `HU-Amt` | Lesen, Ausführen | Ändern | Vollzugriff |
| `HU-Amt\Data\Laufend` | Ändern | Ändern | Vollzugriff |
| `HU-Amt-Daten` | Ändern (**kein** Vollzugriff) | Ändern | Vollzugriff |
| `HU-Amt-Daten\Verwalter.txt` | nur Lesen, Vererbung unterbrochen | Ändern | Vollzugriff |

Warum:
- **Programm nur lesen:** Wer Programmdateien ändern darf, kann Code einschleusen, der bei allen anderen läuft.
- **`Data\Laufend` im Programmordner:** Dort meldet sich jedes offene HU-Amt an, damit ein Update erst startet, wenn alle geschlossen haben. Darum müssen die Benutzer dort schreiben dürfen.
- **Kein Vollzugriff auf die Daten:** Vollzugriff enthält „Unterordner und Dateien löschen“ und „Berechtigungen ändern“. Damit könnte jeder auch geschützte Dateien ersetzen.
- **`Verwalter.txt` nur lesen:** Wer diese Liste ändert, macht sich selbst zum Verwalter.
- **Einstellung des Datenordners:** Sie steht in `HU-Amt\Config\datenort.json`. Der Ordner `Config` ist durch „Programm nur lesen“ schon geschützt. Wer die Datei ändern dürfte, könnte alle Daten samt Verwalter-Liste umlenken.

Freigaberechte (SMB) der Freigabe: mindestens „Ändern“ für die Benutzer. Den Rest regelt NTFS.

HU-Amt prüft die NTFS-Rechte selbst und zeigt den Verwaltern unter Einstellungen › Verwalter eine Warnung, wenn sie zu offen sind. Freigaberechte sieht das Programm nicht.

## 4. Reihenfolge

1. Gruppen anlegen, Verwalter in beide Gruppen, alle anderen in `GG_HU-Amt-Benutzer`.
2. Die beiden Ordner anlegen (leer) und die Rechte für `HU-Amt` und `HU-Amt-Daten` setzen (Befehle unten, Schritte A und B).
3. Ein Verwalter installiert HU-Amt nach `HU-Amt` und stellt im Programm den Datenordner auf `HU-Amt-Daten` um.
4. Danach die Rechte für `HU-Amt\Data\Laufend` und `HU-Amt-Daten\Verwalter.txt` setzen (Schritte C und D). Beide gibt es erst nach Schritt 3.

## 5. Befehle

Zielmaschine: der Fileserver mit der Freigabe, PowerShell oder Eingabeaufforderung **als Administrator**. Vorher ersetzen:
- `D:\Freigaben\freigabe` durch den lokalen Pfad der Freigabe am Server,
- `DOMAENE` durch den NetBIOS-Namen der Domäne.

`*S-1-5-18` ist SYSTEM, `*S-1-5-32-544` sind die lokalen Administratoren. Über die Kennungen klappt es auf deutschem und englischem Windows gleich.

```
rem A: Programmordner
icacls "D:\Freigaben\freigabe\HU-Amt" /inheritance:r /grant:r "*S-1-5-18:(OI)(CI)F" "*S-1-5-32-544:(OI)(CI)F" "DOMAENE\GG_HU-Amt-Verwalter:(OI)(CI)M" "DOMAENE\GG_HU-Amt-Benutzer:(OI)(CI)RX"

rem B: Datenordner
icacls "D:\Freigaben\freigabe\HU-Amt-Daten" /inheritance:r /grant:r "*S-1-5-18:(OI)(CI)F" "*S-1-5-32-544:(OI)(CI)F" "DOMAENE\GG_HU-Amt-Verwalter:(OI)(CI)M" "DOMAENE\GG_HU-Amt-Benutzer:(OI)(CI)M"

rem C: Anmeldungen laufender Programme (nach der Installation)
icacls "D:\Freigaben\freigabe\HU-Amt\Data\Laufend" /grant "DOMAENE\GG_HU-Amt-Benutzer:(OI)(CI)M"

rem D: Verwalter-Liste (nach dem Umstellen des Datenordners)
icacls "D:\Freigaben\freigabe\HU-Amt-Daten\Verwalter.txt" /inheritance:r /grant:r "*S-1-5-18:F" "*S-1-5-32-544:F" "DOMAENE\GG_HU-Amt-Verwalter:M" "DOMAENE\GG_HU-Amt-Benutzer:R"
```

Prüfen: `icacls "<Pfad>"` zeigt die gesetzten Rechte. In HU-Amt zeigt Einstellungen › Verwalter danach keine Warnung mehr.

## 6. Später: Rechte je Modul

Jedes Modul hat seinen eigenen Ordner in `HU-Amt-Daten` (z. B. `Buchung`, `Kostenzuordnung`, `Reiseabrechnung`). Sollen nur bestimmte Bedienstete ein Modul sehen, bekommt dieser Ordner eine eigene Gruppe (z. B. `GG_HU-Amt-Buchung`) und die allgemeine Benutzergruppe dort keine Rechte mehr. Eine kommende Version von HU-Amt blendet Module ohne Zugriff dann aus. Dazu folgt eine eigene Anleitung, bis dahin bitte nichts an den Modulordnern ändern.

Später ist außerdem eine eigene Datenbank für HU-Amt auf einem SQL Server möglich. Datenbanken anderer Programme werden dabei nie berührt.
