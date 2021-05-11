# Korrektur
Prüfungskorrektur für Online-Prüfungen auf OLAT

In diesem Repository liegt das Tool für die Korrektur von OLAT-Prüfungen. Es ist nur ein Skript (```Korrektur.py```), welches alle Schritte der Korrektur und des Export übernimmt. Eine schrittweise Anleitung zum Ausführen des Skripts ist in der Dokumentation (```OLAT_Korrektur_Anleitung.docx```) zu finden. Das Tool nimmt als Eingabe die CSV-Datei, wie sie von OLAT exportiert wird. Öffnen Sie diese Datei zuvor nicht in Excel oder anderen Programmen, da dadurch einzelne Zellen beschädigt werden können. Laden Sie es direkt im Tool.

Zudem enthält das Repository die ausführbare Datei ```Korrektur.exe```. Falls Sie kein Python auf Ihrem Windows-Computer haben oder das Paket ```fpdf``` nicht in Python geladen haben, ist dieses Programm sinnvoller als das rohe Python-Skript.

## Update auf QTI 2.1
Im FS2021 stellt die UZH auf OLAT14 und QTI2.1 um. Dadurch verändert sich sowohl die Eingabe der Tests als auch die Auswertung der Daten etwas. Die aktuelle Version kann sowohl mit dem bisherigen Format als auch neu mit QTI2.1 Exporten arbeiten. Für QTI2.1 benötigt man entweder eine Speicherung der Prüfung als ZIP Datei und eine XLSX-Datei mit den Daten der Studierenden oder eine einzelne ZIP-Datei, die sowohl die QTI-Dateien als auch das XLSX enthält. Je nach Export kriegt man die eine oder andere Version.
Beim Import der Daten kann man auswählen, ob man die Daten aus einer alten CSV-Datei (oftmals mit Endung XLS), aus einer QTI2.1 Datei (entweder XLSX und ZIP oder nur ZIP) oder aus einer bestehenden JSON-Datei laden möchte.

## Beispieldaten 
Das Verzeichnis enthält zudem die Datei ```Testresultate_Beispiel.csv```, die genau das Format der OLAT-Testresultate (Version vor QTI 2.1) hat, aber erfundene Daten zu Demonstrationszwecken enthält. Diese Daten wurden auch in der Dokumentation verwendet. Sie können damit die einzelnen Schritte der Dokumentation nachspielen.
