# Korrektur
Prüfungskorrektur für Online-Prüfungen auf OLAT

In diesem Repository liegt das Tool für die Korrektur von OLAT-Prüfungen. Es ist nur ein Skript (```Korrektur.py```), welches alle Schritte der Korrektur und des Export übernimmt. Vor der Korrektur mit diesem Skript müssen die Daten von OLAT aber korrekt aufbereitet werden. Die Dokumentation ```Anleitung_Korrektur.pptx``` enthält alle Informationen und Anleitungen zum Schrittweisen Ausführen.

Zudem enthält das Repository die ausführbare Datei ```Korrektur.exe```, in welcher alle Schritte der Pipeline vereint sind. Falls Sie keine Kontrolle über Zwischenschritte brauchen, ist das die einfachste Lösung, um die OLAT-Daten zu korrigieren. Es muss aber auch da eine Spalte F geben, in welcher die Lösungen auf die einzelnen Fragen aufgeführt sind. Sonst funktioniert die automatische Korrektur nicht (siehe Anleitung).

## Beispieldaten 
Das Verzeichnis enthält zudem die Datei ```Eingabe.xls```, für welche der erste Schritt (korrekte Antworten definieren) bereits durchgeführt ist. Das Dokument enthält einen Selbsttest auf OLAT.
