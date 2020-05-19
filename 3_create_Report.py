
## Name of the complete result file with all manual corrections

pdf = False
import time
import os
try:
    from fpdf import FPDF
    pdf = True
except:
    pass


fname = "results_corr.json"
table = "Punktetabelle.xls"
einsicht = ".\\Einsichten\\"

os.system("mkdir "+einsicht)

def create_einsicht(res,r,folder='.\\'): ## Takes one respondent dict
    veranstaltung = "Statistik und Datenanalyse: Einführung"
    disclaimer = " ".join(["In diesem Dokument sehen Sie für alle Fragen der Prüfung",
                            "sowohl Ihre eigene Antwort als auch die Antwort, die für",
                            "diese Frage als korrekt gewertet wurde.",
                            "Darunter sind Ihre Punkte, sowie die maximale Punktzahl",
                            "für die entsprechende Aufgabe notiert.",
                            "Falls es Bemerkungen gibt, sind diese am Ende jeder Frage",
                            "notiert. Die Bemerkungen wurden zum Teil automatisch erstellt",
                            "und beziehen sich auf die Fehler, die zu Abzügen geführt haben.",
                            "\n\n",
                            "Auf jeder Seite wird genau eine Frage dargestellt. Prüfen Sie",
                            "jeweils, ob die Frage korrekt bewertet wurde und ob Sie einen",
                            "allfälligen Abzug von Punkten nachvollziehen können.",
                            "Falls Sie mit einer Entscheidung nicht einverstanden sind, steht",
                            "Ihnen das Mittel eines Wiedererwägungsgesuchs zur Verfügung. Stellen",
                            "Sie ein solches Gesuch zu Handen Ihres Dozenten und beschreiben Sie",
                            "darin möglichst genau, bei welcher Frage Sie aus welchem Grund die",
                            "Punktevergabe ungerechtfertigt finden.",
                            "Das Gesuch wird dann geprüft und Sie erhalten eine schriftliche Rückmeldung.",
                            "\n\n\n\n",
                            "Anmerkung: Dieses Dokument wurde automatisch erstellt und kann ein seltsames",
                            "Layout haben. Die Informationen sind jedoch korrekt und wurden bei",
                            "der Punktevergabe genau so verwendet."])

    p = 0
    mp = 0
    for q in res[r].keys():
        try:
            p+=res[r][q]['Points']
        except:
            p=''
        try:
            mp+=res[r][q]['Max_Points']
        except:
            mp=''

    fname = folder + 'Resultate_'+ r + '.pdf'
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", size=16)
    pdf.set_fill_color(200,200,255)
    pdf.multi_cell(0, 12, txt="Dokument für die Prüfungseinsicht\nVorlesung: "+veranstaltung+"\nTeilnehmer: "+r,
                   align="C",border=1,fill=True)

    pdf.set_font("Arial", size=14,style="B")
    pdf.y = pdf.y+20
    pdf.multi_cell(0,6, txt="Gesamtergebnis")
    if type(p)==str or type(mp)==str:
        pdf.set_font("Arial", size=12)
        pdf.y = pdf.y+5
        pdf.multi_cell(0,6, txt="Es kann kein Ergebnis ausgegeben werden. Die Korrektur für diese Prüfung ist noch nicht abgeschlossen")
    else:
        pdf.set_font("Arial", size=12,style="B")
        pdf.y = pdf.y+5
        pdf.multi_cell(0,6, txt="Sie haben "+str(p)+" von insgesamt "+str(mp)+" Punkten erreicht.")
        

    pdf.set_font("Arial", size=14,style="B")
    pdf.y = pdf.y+20
    pdf.multi_cell(0,6, txt="Erläuterung zu diesem Dokument")

    pdf.set_font("Arial", size=12)
    pdf.y = pdf.y+5
    pdf.multi_cell(0,6, txt=disclaimer)

    pdf.y = pdf.y+40
    pdf.set_font("Arial", size=10, style="I")
    pdf.multi_cell(0,6, txt="Erstellt am: "+str(time.ctime()))


    for q in sorted(list(res[r].keys())):
        pdf.add_page()
        pdf.set_font("Arial", size=16)
        pdf.set_fill_color(220,220,255)
        t = "Aufgabe: {0} ({1})\nMaximale Punktzahl: {2}".format(q,res[r][q]['Title'],res[r][q]['Max_Points'])
        pdf.multi_cell(0, 12, txt=t, align="C",border=1,fill=True)

        pdf.set_fill_color(220,250,220)

        pdf.set_font("Arial", size=14,style="B")
        pdf.y+=20
        pdf.multi_cell(0,6, txt="Korrekte Antwort:")
        pdf.y+=5
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0,6,txt=res[r][q]['Correct'].replace(" // ",'\n'),fill=True,align="C",border=1)

        pdf.set_fill_color(220,220,220)

        pdf.set_font("Arial", size=14,style="B")
        pdf.y+=20
        pdf.multi_cell(0,6, txt="Ihre Antwort:")
        pdf.y+=5
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0,6,txt=res[r][q]['Given'].replace(" // ",'\n'),fill=True,align="C",border=1)

        if len(res[r][q]['Remarks'])>1:
            pdf.set_font("Arial", size=14,style="B")
            pdf.y+=20
            pdf.multi_cell(0,6, txt="Bemerkungen:")
            pdf.y+=5
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0,6,txt=res[r][q]['Remarks'].replace(" // ",'\n'),fill=False,border=1)

        pdf.set_font("Arial", size=14,style="B")
        pdf.y+=20
        pdf.multi_cell(0,6, txt="Sie haben "+str(res[r][q]['Points'])+" von "+str(res[r][q]['Max_Points'])+" Punkten erhalten.")
        
        
        

    pdf.output(fname)


inf = open(fname,'r',encoding="utf-8",errors="ignore")
result = inf.readline()
inf.close()

result = eval(result)

respondents = sorted(list(result.keys()))
questions = sorted(list(result[respondents[0]].keys()))
#print(respondents)
#print(questions)

outf = open(table,'w')
outf.write('\t'.join(['Name','Total','Remarks']+questions)+'\n')
for r in respondents:
    line = [r,'','']
    punkte = 0
    flagged = False
    for q in questions:
        line.append(str(result[r][q]['Points']))
        try:
            punkte+=result[r][q]['Points']
        except:
            punkte = ''
        try:
            if result[r][q]['State']==1:flagged=True
        except:
            pass
        line[1]=str(punkte)
        if flagged:
            line[2]='Achtung: Mindestens eine der Fragen ist noch nicht geprüft'
    outf.write('\t'.join(line)+'\n')
outf.close()

if pdf:
    for r in respondents:
        create_einsicht(result,r,einsicht)
                    
