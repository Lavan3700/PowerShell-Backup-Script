# PowerShell-Backup-Script

Zeile 75 & 76 Müssten noch bearbeitet werden, dort muss man noch eine E-Mail hinterlegen, welches fürs versenden vom Bestätigungsmail dient.

Das PowerShell-Skript komprimiert einen vom Benutzer angegebenen Ordner in eine ZIP-Datei. Dabei wird geprüft, ob der Ordner und der Speicherort der ZIP-Datei existieren und ob eine Datei mit demselben Namen bereits vorhanden ist. Nach der Komprimierung zählt das Skript die Anzahl der Dateien im Ordner, öffnet den Speicherort im Windows Explorer und sendet eine Bestätigungsnachricht per E-Mail an den Benutzer.
