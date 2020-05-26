
import time
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import simpledialog
from fpdf import FPDF

class CMD: #Auxilliary function for callbacks using parameters. Syntax: CMD(function, argument1, argument2, ...)
    def __init__(s1, func, *args):
        s1.func = func
        s1.args = args
    def __call__(s1, *args):
        args = s1.args+args
        s1.func(*args)

class Anzeige(Frame):
    def __init__(self, master=None):        
        Frame.__init__(self,master)
        top=self.winfo_toplevel() #Flexible Toplevel of the window
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
        self.grid(sticky=N+S+W+E)

        self.set_window()

    def set_window(self):
        global settings

        self.m = Menubutton(self,text="Navigation",relief=RAISED)
        self.m.grid(row=0,column=0,sticky=N+W+E)
        self.m.menu = Menu(self.m,tearoff=0)
        self.m["menu"]=self.m.menu
        self.m.frage = Menu(self.m.menu,tearoff=0)
        self.m.menu.add_cascade(label="Andere Frage",menu=self.m.frage)

        #self.m.menu.add_command(label='Andere Daten Laden',command=self.read_new_data)
        self.m.menu.add_command(label='Übersicht',command=self.check_completeness)
        self.m.menu.add_command(label='Ergebnisse exportieren',command=self.export_results)


        self.cq = Label(self,text="  Aktuelle Frage: ")
        self.cq.grid(row=0,column=1,columnspan=2,sticky=W)

        b = Button(self,text="Übersicht",command=self.check_completeness)
        b.grid(row=0,column=3)

        spacer = Label(self,text='')
        spacer.grid(row=2,column=0)
        la = Label(self,text='Gegebene Antwort:')
        la.grid(row=3,column=0,columnspan=3,sticky=W)
        lb = Label(self,text='Korrekte Antwort:')
        lb.grid(row=3,column=3,sticky=W)
        lc = Label(self,text='Bemerkungen:')
        lc.grid(row=6,column=0,sticky=W)

        self.ysc1 = Scrollbar(self, orient=VERTICAL)
        self.ysc1.grid(row=4,column=2,sticky=N+S)
        self.geg = Text(self,width=60,height=10,wrap=WORD,bg="#eeeeff",
                        relief=SUNKEN,yscrollcommand=self.ysc1.set)
        self.geg.grid(row=4,column=0,columnspan=2)
        self.ysc1["command"]=self.geg.yview

        self.ysc2 = Scrollbar(self, orient=VERTICAL)
        self.ysc2.grid(row=4,column=5,sticky=N+S)
        self.req = Text(self,width=60,height=10,wrap=WORD,bg="#eeffee",
                        relief=SUNKEN,yscrollcommand=self.ysc2.set)
        self.req.grid(row=4,column=3,columnspan=2)
        self.ysc2["command"]=self.req.yview

        self.ysc3 = Scrollbar(self, orient=VERTICAL)
        self.ysc3.grid(row=7,column=2,sticky=N+S)
        self.bem = Text(self,width=60,height=10,wrap=WORD,bg="#ffffff",
                        relief=SUNKEN,yscrollcommand=self.ysc3.set)
        self.bem.grid(row=7,column=0,columnspan=2)
        self.ysc3["command"]=self.bem.yview

        self.f = Frame(self)
        self.f.grid(row=6,column=3,rowspan=2,columnspan=3)

        ld = Label(self.f,text="Punkte:")
        ld.grid(row=0,column=0,sticky=W)
        self.points=StringVar()
        self.points.set("KEINE DATEN")
        self.maxpoints=StringVar()
        self.maxpoints.set("(Max: --)")
        self.pts = Entry(self.f,textvariable=self.points)
        self.pts.grid(row=1,column=0,columnspan=2,sticky=W)
        self.mpts = Label(self.f,textvariable=self.maxpoints)
        self.mpts.grid(row=2,column=0,columnspan=2,sticky=W)

        self.insec = IntVar()
        self.insec.set(0)
        self.cb = Checkbutton(self.f,variable=self.insec)
        self.cb.grid(row=4,column=0,sticky=W)
        le = Label(self.f,text="Bitte checken!")
        le.grid(row=4,column=1,sticky=W)
        

        spacer = Label(self.f,text="     ")
        spacer.grid(row=3,column=2)
        save = Button(self.f,text="Speichern",bg="#aaffaa",command=self.store,width=20,height=3)
        save.grid(row=1,rowspan=2,column=3)

        spacer = Label(self,text="\n\n")
        spacer.grid(row=8,column=0)

        self.f2 = Frame(self)
        self.f2.grid(row=9,column=0,columnspan=10)
        p = Button(self.f2,text="<< Vorheriger <<",command=self.prev)
        n = Button(self.f2,text=">> Nächster >>",command=self.next)
        p.grid(row=0,column=0)
        n.grid(row=0,column=2)
        self.rid = StringVar()
        self.rid.set("start")
        self.cr = Label(self.f2,textvariable=self.rid,width=60)
        self.cr.grid(row=0,column=1)

        if self.read_data():
            self.refresh()

    def export_results(self):
        if messagebox.askokcancel("Wirklich exportieren?","Sollen die Daten wirklich exportiert werden? Prüfen Sie vorher, ob wirklich alle Korrekturen vorgenommen sind.\nPrüfungen mit unvollständigen Punkten werden einen fehlenden Wert bei der Gesamtpuktzahl haben."):
            write_results(self.resp)

    def read_new_data(self):
        print('hallo Welt')
        pass


    def read_data(self):
        j = 0
        try:
            inf = open(settings['J_Filename_Out'],'r',encoding='utf-8',errors='ignore')
            j = eval(inf.readline())
            inf.close()  
        except:
            try:
                inf = open(settings['J_Filename'],'r',encoding='utf-8',errors='ignore')
                j = eval(inf.readline())
                inf.close()
            except:
                try:
                    j = read_xls(settings['E_Filename'])[0]
                except:
                    messagebox.showerror("Kein Korrekturfile","Es wurde kein Korrekturfile gefunden und es gibt keine XLS-Tabelle namens '"+settings['E_Filename']+"'.\n\nBitte wählen Sie eine gültige OLAT-Ergebnistabelle MIT RESULTATEN in Spalte F aus.")
                    fname = filedialog.askopenfilename(**{'defaultextension':['.xls','.txt'],
                                                          'filetypes':[('Result File',['.xls','.txt']),('Other','.*')]})
                    try:
                        j = read_xls(fname)[0]
                    except:
                        pass ## Goddammit, at least select a valid file.
                if type(j) == dict:
                    o = open(settings['J_Filename'],'w',encoding='utf-8',errors='ignore')
                    o.write(str(j))
                    o.close()
                                        
                    
        if type(j)==dict:
            self.resp = j
            self.respondents = []
            self.questions = []
            self.question_titles = []
            for k in j.keys():
                self.respondents.append(k)
            for q in j[self.respondents[0]].keys():
                self.questions.append(q)
                self.question_titles.append(j[self.respondents[0]][q]['Title'])
                self.m.frage.add_command(label=str(q),command=CMD(self.change_q,q))
            return True
        else:
            return False
                    

    def next(self):
        global settings
        if self.store(False):
            settings['Curr_R']+=1
            if not settings['Curr_R']<len(self.respondents):
                messagebox.showerror("Geht nicht weiter","Keine Weiteren Teilnehmer")
                settings['Curr_R']=len(self.respondents)-1
            self.refresh()

    def prev(self):
        global settings
        if self.store(False):
            settings['Curr_R']-=1
            if not settings['Curr_R']>=0:
                messagebox.showerror("Geht nicht weiter","Das ist der erste Teilnehmer")
                settings['Curr_R']=0
            self.refresh()
        
    def refresh(self):
        global settings
        r = self.respondents[settings['Curr_R']]
        q = self.questions[settings['Curr_Q']]
        curr = self.resp[r][q]
        #print(curr)

        self.cq['text'] = "  Aktuelle Frage: {0} ({1})".format(q,self.question_titles[settings['Curr_Q']])

        self.rid.set("Teilnehmer: '{0}' ({1}/{2})".format(r,settings['Curr_R']+1,len(self.respondents)))

        self.points.set(curr['Points'])
        self.maxpoints.set("(Max: {0})".format(curr['Max_Points']))
        if type(curr['Points'])==str:
            self.pts["bg"]="#ffaa90"
        else:
            self.pts["bg"]="#ffffff"

        self.geg["state"]=NORMAL
        self.req["state"]=NORMAL
        self.geg.delete("1.0",END)
        self.geg.insert(END,curr['Given'].replace(' // ','\n'))
        self.req.delete("1.0",END)
        self.req.insert(END,curr['Correct'].replace(' // ','\n'))
        self.geg["state"]=DISABLED
        self.req["state"]=DISABLED

        self.bem.delete("1.0",END)
        self.bem.insert(END,curr['Remarks'].replace(' // ','\n'))

        if 'State' in curr.keys():
            if curr['State'] == 1:
                self.insec.set(1)
            else:
                self.insec.set(0)
        else:
            self.insec.set(0)


    def change_q(self,question):
        global settings
        settings['Curr_Q']=self.questions.index(question)
        self.refresh()

    def check_completeness(self):
        #print(len(self.resp))
        width = 1000
        height = 500

        nq = len(self.questions)
        nr = len(self.respondents)

        if width < nr*14:
            width = nr*14
        if height < nq*20:
            height = nq*20

        self.case = StringVar()
        self.case.set("Maus über ein Feld, um Informationen zu sehen")

        
        self.infobox = Toplevel(self)
        self.infobox.rowconfigure(1, weight=1)
        self.infobox.columnconfigure(1, weight=1)
        self.infobox.title("Übersicht über Korrekturen")
        la = Label(self.infobox,textvariable=self.case)
        la.grid(row=0,column=1)
        self.infobox.ysc = Scrollbar(self.infobox, orient=VERTICAL)
        self.infobox.ysc.grid(row=1,column=2,sticky=N+S)
        self.infobox.xsc = Scrollbar(self.infobox, orient=HORIZONTAL)
        self.infobox.xsc.grid(row=2,column=1,sticky=E+W)
        self.infobox.plot = Canvas(self.infobox,bd=0,width=width,height=height,bg="#ffffff", scrollregion=(0, 0, width, height),
                                   yscrollcommand=self.infobox.ysc.set, xscrollcommand=self.infobox.xsc.set)
        self.infobox.plot.grid(row=1,column=1,sticky=N+E+S+W)
        self.infobox.ysc["command"]=self.infobox.plot.yview
        self.infobox.xsc["command"]=self.infobox.plot.xview

        qstep = (height-50)//nq
        qt = []
        for i in range(nq):
            qt.append(i*qstep+25)

        rstep = (width-125)//nr
        rt = []
        for i in range(nr):
            rt.append(i*rstep+100)

        for qi in range(nq):
            q = qt[qi]
            self.infobox.plot.create_text(50,(q+qstep//2),text=self.questions[qi])
            for ri in range(nr):
                r = rt[ri]
                item = self.resp[self.respondents[ri]][self.questions[qi]]
                if type(item['Points'])==str:
                    col="#ff9090"
                    prob = "Noch keine Punkte"
                elif not 'State' in item.keys():
                    col="#aaffaa"
                    prob = "Noch nicht manuell kontrolliert"
                elif item['State']==-1:
                    col="#aaffaa"
                    prob = "Noch nicht manuell kontrolliert"
                elif item['State']==1:
                    col="#ffff00"
                    prob = "Muss noch geprüft werden. Es gibt ein Problem."
                else:
                    col="#00ff00"
                    prob = "OK."
                b = self.infobox.plot.create_rectangle(r,q,r+rstep-1,q+qstep-2,fill=col)
                self.infobox.plot.tag_bind(b, '<Button-1>', CMD(self.jump,ri,qi))
                self.infobox.plot.tag_bind(b, '<Enter>', CMD(self.change_lab,ri,qi,prob))

    def jump(self,resp,quest,event=""):
        #print(resp,quest)
        settings['Curr_R']=resp
        settings['Curr_Q']=quest

        self.infobox.destroy()
        self.refresh()

    def change_lab(self,resp,quest,problem,event=""):
        self.case.set("Teiln:" +str(self.respondents[resp])+' / Frage: '+str(self.questions[quest])+' ('+problem+')')
            

    def store(self,refresher=True):
        ##Entry Points
        ##Checkbox
        ##Entry Remarks

        r = self.respondents[settings['Curr_R']]
        q = self.questions[settings['Curr_Q']]
        #curr = self.resp[r][q]
        

        accept = True
        p = self.points.get()       
        cb = self.insec.get()
        rem = self.bem.get("1.0",END)[:-1].replace('\n',' // ')
        #print([p,cb,rem])

        try:
            p = float(p)
        except:
            accept=False
            p=-99
            messagebox.showinfo("Ungültige Eingabe", "Bitte bei Punkten eine Zahl eingeben")

        if accept and (p > self.resp[r][q]['Max_Points'] or p<0):
            accept=False
            messagebox.showinfo("Ungültige Eingabe", "Punkte müssen zwischen 0 und Maximum liegen.")

        if accept:
            self.resp[r][q]['Points']=p
            self.resp[r][q]['Remarks']=rem
            self.resp[r][q]['State']=cb

        outf = open(settings['J_Filename_Out'],'w',encoding='utf-8',errors='ignore')
        outf.write(str(self.resp))
        outf.close()

        if refresher:
            self.refresh()

        return accept

def baum_schreiben(tdic, einr = 0, inherit = '', spc = '    ', trunc=0, lists=False):
    if type(tdic)==dict:
        if inherit == '': inherit = '{\n'
        for k in sorted(tdic.keys()):
            if type(tdic[k]) == dict:
                inherit = inherit + einr*spc + str(k) + ': ' + baum_schreiben(tdic[k],einr+1,inherit = '{\n',trunc=trunc)
            elif type(tdic[k]) == list and lists:
                inherit = inherit + einr*spc + str(k) + ': ' + baum_schreiben(tdic[k],einr+1,inherit = '[\n',trunc=trunc)
            else:
                value = tdic[k]
                if type(value)==str:
                    value = "'"+value+"'"
                elif type(value) in [int,float]:
                    value = str(value)
                elif type(value) == list:
                    value = str(value)
                else:
                    value = str(value)
                if len(value) > trunc and trunc > 0:
                    tail = int(trunc/2)
                    value = value[:tail] + '...'+value[-tail:] + ' ('+str(len(value))+' characters)'
                    
                inherit = inherit + einr*spc + str(k) + ': '+ value + '\n'

        inherit = inherit + einr*spc + '}\n'
    elif type(tdic)==list:
        if inherit == '': inherit = '[\n'
        for e in tdic:
            if type(e) == dict:
                inherit = inherit + einr*spc + baum_schreiben(e,einr+1,inherit = spc+'{\n',trunc=trunc)
            elif type(e) == list and lists:
                inherit = inherit + einr*spc + baum_schreiben(e,einr+1,inherit = spc+'[\n',trunc=trunc)
            else:
                value = e
                if type(value)==str:
                    value = "'"+value+"'"
                elif type(value) in [int,float]:
                    value = str(value)
                elif type(value) == list:
                    value = str(value)
                else:
                    value = str(value)
                if len(value) > trunc and trunc > 0:
                    tail = int(trunc/2)
                    value = value[:tail] + '...'+value[-tail:] + ' ('+str(len(value))+' characters)'
                    
                inherit = inherit + einr*spc + value + ',\n'

        inherit = inherit + einr*spc + ']\n'
            
    return inherit


def scan_question(ls,line):
    scan = True
    outdic = {'Ans':{},'Max':0}
    while scan:
        line+=1
        l = ls[line][:-1].split('\t')
        if l[2]=='maxValue':
            try:
                outdic['Max']=int(l[3])
            except:
                pass
        elif '_' in l[2]:
            answer = l[5]
            try:
                answer = eval(answer)
            except:
                pass
            outdic['Ans'][l[2]]=answer
            
        if ls[line][0]=='Q' or ls[line][:5]=='\t\t\t\t\t':
            scan = False

    return outdic


def read_table(lines):
    d = {}
    vlist=lines[0][:-1].split('\t')
    for v in vlist:
        d[v]=[]

    for l in lines[1:]:
        values = l[:-1].split('\t')
        for vi in range(len(vlist)):
            d[vlist[vi]].append(values[vi])

    return d

def mark_question(q,a):
    #print(a)
    mark = {}
    if q['Type']=='SCQ':
        points = 0
        s1 = []
        s2 = []
        for ans in q['Ans'].keys():
            s1.append(str(q['Ans'][ans]))
            s2.append(a[ans])
            if q['Ans'][ans] in ['1',1] and a[ans] in ['1',1]:
                points = q['Max']
        mark = {'Correct':' - '.join(s1),
                'Given':' - '.join(s2),
                'Remarks':'',
                'Points':points,
                'Max_Points':q['Max'],
                'Type':q['Type'],
                'Title':q['Title']}

    elif q['Type']=='MCQ':
        points = q['Max']
        ded = float(q['Max'])/2
        s1 = []
        s2 = []
        for ans in q['Ans'].keys():
            s1.append(str(q['Ans'][ans]))
            s2.append(a[ans])
            if str(q['Ans'][ans])==str(a[ans]):
                pass
            else:
                points-=ded
            if points<0:points=0
            
        mark = {'Correct':' - '.join(s1),
                'Given':' - '.join(s2),
                'Remarks':'',
                'Points':points,
                'Max_Points':q['Max'],
                'Type':q['Type'],
                'Title':q['Title']}

    elif q['Type'] in ['KPRIM','KPR']:
        points = q['Max']
        ded = float(q['Max'])/2
        s1 = []
        s2 = []
        for ans in q['Ans'].keys():
            s1.append(str(q['Ans'][ans]))
            s2.append(a[ans])
            if str(q['Ans'][ans])==str(a[ans]):
                pass
            else:
                points-=ded
            if points<0:points=0
            
        mark = {'Correct':' '.join(s1),
                'Given':' '.join(s2),
                'Remarks':'',
                'Points':points,
                'Max_Points':q['Max'],
                'Type':q['Type'],
                'Title':q['Title']}

    elif q['Type']=='FIB':
        ac = True ## Only automatically mark if all is correct
        remarks = ''
        s1 = []
        s2 = []
        for ans in q['Ans'].keys():
            s1.append(str(q['Ans'][ans]))
            s2.append(a[ans])
            if type(q['Ans'][ans]) in (str,int,float):
                if not a[ans] == str(q['Ans'][ans]):
                    ac = False
                    remarks+=ans+': '+str(a[ans])+' nicht gleich '+ str(q['Ans'][ans])+'\n'
                
            elif type(q['Ans'][ans]) == list:
                try:
                    antwort = float(a[ans])
                    if antwort < q['Ans'][ans][0] or  antwort > q['Ans'][ans][1]:
                        ac = False
                        remarks+=ans+': '+str(a[ans])+' nicht im Intervall '+ str(q['Ans'][ans])+'\n'
                except:
                    ac=False
                    remarks+=ans+': '+str(a[ans])+' ist keine Zahl.\n'

        if ac:
            points = q['Max']
        else:
            points = ''

        mark = {'Correct':' // '.join(s1),
                'Given':' // '.join(s2),
                'Remarks':remarks.replace('\n',' // '),
                'Points':points,
                'Max_Points':q['Max'],
                'Type':q['Type'],
                'Title':q['Title']}

    else: ## for essay questions
        s1 = []
        s2 = []
        for ans in q['Ans'].keys():
            s1.append(str(q['Ans'][ans]))
            s2.append(a[ans])
                  
        mark = {'Correct':' '.join(s1),
                'Given':' '.join(s2),
                'Remarks':'',
                'Points':'',
                'Max_Points':q['Max'],
                'Type':q['Type'],
                'Title':q['Title']}           
    return mark



def read_xls(fname = "Eingabe.xls"):
    kc = 'Laufnummer'
    add_kc = ["Vorname","Nachname","Benutzername"]
    
    inf = open(fname,'r',
           encoding="utf-8",errors='ignore')
    ls = inf.readlines()
    inf.close()

    tstart = 1
    tend = 2
    while not ls[tend][:3] in ['\n','\t\t\t']:
        tend+=1

    resp = read_table(ls[tstart:tend])

    #print(list(resp.keys()))
    #print(resp[kc])

    questions = {}
    qkeys = {}
    for i in range(tend+2,len(ls)):
        l=ls[i]
        if l[0]=='Q':
            line = l.split('\t')
            q = line[0].split(':')
            if len(q)==3:
                questions[q[2]]={'Line':i,'Type':q[1]}
                ans = scan_question(ls,i)
                questions[q[2]]['Ans']=ans['Ans']
                questions[q[2]]['Max']=ans['Max']
                questions[q[2]]['Title']=line[1]
                for a in ans['Ans'].keys():
                    qkeys[a]=q[2]

    rd = {}
    for i in range(len(resp[kc])):
        r = resp[kc][i]
        add = []
        for a in add_kc:
            try:
                add.append(resp[a][i])
            except:
                pass
        if len(add)>0:
            r+=" ({0})".format(', '.join(add))
        rd[r]={}
        for q in questions.keys():
            response = {}
            for a in questions[q]['Ans'].keys():
                response[a] = resp[a][i]
            rd[r][q] = mark_question(questions[q],response)

    return [rd,questions]

def create_einsicht(res,r,folder='.\\',veranstaltung=''): ## Takes one respondent dict
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
    header = "Dokument für die Prüfungseinsicht\n"+veranstaltung+"\nTeilnehmer: "+r
    pdf.multi_cell(0, 12, txt=header,
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

    try:
        pdf.output(fname)
    except:
        er = open(fname+"_ERROR.txt","w",encoding="utf-8",errors="ignore")
        er.write("Error. Could not export case:"+r+"\nThere must be a strange special character in this respondent's answers. Please mend manually.\n")
        er.close()


def write_results(result = None, pdf=True, table="Punktetabelle.xls"):
    respondents = sorted(list(result.keys()))
    questions = sorted(list(result[respondents[0]].keys()))

    outf = open(table,'w',encoding='utf-8',errors='ignore')
    outf.write('\t'.join(['Laufnummer','Name','Total','Remarks']+questions)+'\n')
    for r in respondents:
        try:
            lf = int(r)
            n = r
        except:
            try:
                lf = int(r.split(' ')[0])
                n = ' '.join(r.split(' ')[1:]).replace('(','').replace(')','')
            except:
                lf = r
                n = r
        lf = str(lf)
        line = [lf,n,'','']
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
            line[2]=str(punkte)
            if flagged:
                line[3]='Achtung: Mindestens eine der Fragen ist noch nicht geprüft'
        outf.write('\t'.join(line)+'\n')
    outf.close()

    if pdf:
        titel = simpledialog.askstring("Ausgabetitel","Welche Überschrift sollen die Dokumente für die Prüfungseinsicht tragen?\n\nWenn Sie dieses Textfeld leer lassen, werden keine PDFs erstellt.\n\n(z.B: 'Vorlesung: Statistik und Datenanalyse')")
        if len(titel)>2:
            for r in respondents:
                create_einsicht(result,r,veranstaltung=titel)

    messagebox.showinfo("Export erfolgreich","Alle Daten wurden erfolgreich exportiert")
                    

global settings
settings = {}
settings['Curr_R']=0
settings['Curr_Q']=0
settings['Insecurities']=[]
settings['E_Filename']="Eingabe2.xls"
settings['J_Filename']="results.json"
settings['J_Filename_Out']="results_corr.json"

root = Tk()
korr = Anzeige(root)
root.mainloop()
