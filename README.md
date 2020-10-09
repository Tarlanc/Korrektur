# Korrektur
Prüfungskorrektur für Online-Prüfungen auf OLAT

In diesem Repository liegt das Tool für die Korrektur von OLAT-Prüfungen. Es ist nur ein Skript (```Korrektur.py```), welches alle Schritte der Korrektur und des Export übernimmt. Eine schrittweise Anleitung zum Ausführen des Skripts ist in der Dokumentation (```OLAT_Korrektur_Anleitung.docx```) zu finden. Das Tool nimmt als Eingabe die CSV-Datei, wie sie von OLAT exportiert wird. Öffnen Sie diese Datei zuvor nicht in Excel oder anderen Programmen, da dadurch einzelne Zellen beschädigt werden können. Laden Sie es direkt im Tool.

Zudem enthält das Repository die ausführbare Datei ```Korrektur.exe```. Falls Sie kein Python auf Ihrem Windows-Computer haben oder das Paket ```fpdf``` nicht in Python geladen haben, ist dieses Programm sinnvoller als das rohe Python-Skript.

## Beispieldaten 
Das Verzeichnis enthält zudem die Datei ```Testresultate_Beispiel.csv```, die genau das Format der OLAT-Testresultate hat, aber erfundene Daten zu Demonstrationszwecken enthält. Diese Daten wurden auch in der Dokumentation verwendet. Sie können damit die einzelnen Schritte der Dokumentation nachspielen.
