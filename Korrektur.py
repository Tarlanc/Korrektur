
import time
import math
import zipfile
import re
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import simpledialog
from io import BytesIO


try:
    from fpdf import FPDF
    pdf = True
except:
    pdf = False

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
        self.bind_all("<Next>",self.nextq)
        self.bind_all("<Prior>",self.prevq)

        self.set_window()

    def set_window(self):
        global settings

        self.m = Menubutton(self,text="Navigation",relief=RAISED)
        self.m.grid(row=0,column=0,sticky=N+W+E)
        self.m.menu = Menu(self.m,tearoff=0)
        self.m["menu"]=self.m.menu
        self.m.frage = Menu(self.m.menu,tearoff=0)
        self.m.laden = Menu(self.m.menu,tearoff=0)
        self.m.menu.add_cascade(label="Zu Frage springen",menu=self.m.frage)
        self.m.menu.add_cascade(label='Daten Laden',menu=self.m.laden)
        self.m.menu.add_command(label='Sicherungskopie erstellen',command=self.store_data)        
        self.m.menu.add_command(label='Übersicht',command=self.check_completeness)
        self.m.menu.add_command(label='Punkte-Übersicht',command=self.check_scores)
        self.m.menu.add_command(label='Plagiats-Detektor',command=self.check_plagiat)
        self.m.menu.add_command(label='Ergebnisse exportieren',command=self.export_results)
        self.m.menu.add_separator()
        self.m.menu.add_command(label='Daten hinzufügen',command=self.attach_new_data)
        self.m.menu.add_command(label='Lösungsschema hinzufügen',command=self.attach_new_corr)

        self.m.laden.add_command(label="CSV/XLS (QTI1.2, vor 2021)",command=CMD(self.read_new_data,"CSV"))
        self.m.laden.add_command(label="XLSX und ZIP (QTI2.1)",command=CMD(self.read_new_data,"XLSX"))
        self.m.laden.add_command(label="JSON Datei",command=CMD(self.read_new_data,"JSON"))
        self.m.laden.add_command(label="ZIP mit Resultaten (QTI2.1 Test)",command=CMD(self.read_new_data,"ZIP"))
        
        self.qbox = Frame(self)
        self.qbox.grid(row=0,column=1,columnspan=4)

        self.cq = Label(self.qbox,text="  Aktuelle Frage: ",width=80,anchor=W)
        self.cq.grid(row=0,column=1,sticky=W)
        b = Button(self.qbox,width=3,text="-",command=self.prevq)
        b.grid(row=0,column=0,sticky=E)
        b = Button(self.qbox,width=3,text="+",command=self.nextq)
        b.grid(row=0,column=2,sticky=W)

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
        self.geg = Text(self,width=60,height=14,wrap=WORD,bg="#eeeeff",
                        relief=SUNKEN,yscrollcommand=self.ysc1.set)
        self.geg.grid(row=4,column=0,columnspan=2)
        self.ysc1["command"]=self.geg.yview

        self.ysc2 = Scrollbar(self, orient=VERTICAL)
        self.ysc2.grid(row=4,column=5,sticky=N+S)
        self.req = Text(self,width=60,height=14,wrap=WORD,bg="#eeffee",
                        relief=SUNKEN,yscrollcommand=self.ysc2.set)
        self.req.grid(row=4,column=3,columnspan=2)
        self.ysc2["command"]=self.req.yview
        self.req.bind('<Button-1>',self.correctans)

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
            messagebox.showinfo("Daten geladen","Sitzung erfolgreich wiederhergestellt") ## Sinnlose Message, damit der Focus stimmt. K.A. wieso!

    def export_results(self):
        if messagebox.askokcancel("Wirklich exportieren?","Sollen die Daten wirklich exportiert werden? Prüfen Sie vorher, ob wirklich alle Korrekturen vorgenommen sind.\nPrüfungen mit unvollständigen Punkten werden einen fehlenden Wert bei der Gesamtpuktzahl haben."):
            write_results(self.data,questions=self.questions,respondents=self.respondents)


    def attach_new_data(self):
        if messagebox.askokcancel("Vorsicht!","Sie sind dabei, die bestehenden Daten durch zusätzliche Daten aus einer anderen Quelle zu ergänzen.\n\nVergewissern Sie sich, dass Sie wissen, was Sie tun und legen Sie vorher besser eine Sicherungskopie der aktuell geöffneten Daten an.\n\nDie aktuellen Daten werden überschrieben!\n\nWollen Sie fortfahren?"):
            fname = filedialog.askopenfilename(**{'defaultextension':['.json'],
                                                  'filetypes':[('JSON','.json'),('Other','.*')]})
            inf = open(fname,'r',encoding='utf-8',errors='ignore')
            j = eval(inf.readline())
            inf.close()

            add_resp = []
            add_quest = []

            for r in j['Resp'].keys():
                if not r in self.respondents:
                    add_resp.append(r)
            for q in j['Quest'].keys():
                if not q in self.questions:
                    add_quest.append(q)

            if len(add_resp)+len(add_quest)==0:
                self.combine_jsons(j['Resp'])
            else:
                messagebox.showerror("MISMATCH","Die neue Datei enthält Daten, die nicht mit der aktuell geöffneten vereibar sind.\n\n"+str(len(add_quest))+" Fragen\n"+str(len(add_resp))+" Befragte\n\nstimmen nicht überein")          

    def attach_new_corr(self):
        if messagebox.askokcancel("Vorsicht!","Sie sind dabei, das Lösungsschema dieser Korrektur durch ein anderes zu ersetzen. Damit werden auch die automatischen Korrekturen überschrieben. \n\nVergewissern Sie sich, dass Sie wissen, was Sie tun und legen Sie vorher besser eine Sicherungskopie der aktuell geöffneten Daten an.\n\nDie aktuellen Daten werden überschrieben!\n\nWollen Sie fortfahren?"):
            fname = filedialog.askopenfilename(**{'defaultextension':['.json'],
                                                  'filetypes':[('JSON','.json'),('Other','.*')]})
            inf = open(fname,'r',encoding='utf-8',errors='ignore')
            j = eval(inf.readline())
            inf.close()
            
            add_quest = []
            for q in j['Quest'].keys():
                if not q in self.questions:
                    add_quest.append(q)

            if len(add_quest)==0:
                for q in j['Quest']: ## If q is incomplete (misses questions), this won't matter. It just updates the ones that exist.
                    valid = True
                    for a in j['Quest'][q]['Ans'].keys():
                        if j['Quest'][q]['Ans'][a]['Correct']==None:valid=False ## Only take Solutions where there is a correct solution defined. Don't overwrite others.

                    print(valid, j['Quest'][q])
                    if valid:
                        self.data['Quest'][q]=j['Quest'][q]
                        self.mark_question(q)
                self.refresh()
            else:
                messagebox.showerror("MISMATCH","Die neue Datei enthält Daten, die nicht mit der aktuell geöffneten vereibar sind.\n\n"+str(len(add_quest))+" Fragen stimmen nicht überein")          


    def combine_jsons(self,r):
        results = dict(self.data['Resp'])
        counter = [0,0,0]
        conflicts = []
        overwrite = ['Points','State','Remarks']
        try:
            ## Ensure compatibility of files prior to calling this function. No Error Handling.
            for p in r.keys():
                for q in r[p].keys():
                    if not 'State' in results[p][q].keys():
                        results[p][q]['State']=-1

                    ## Overwrite if new file holds information about an unknown question/respondent
                    if not 'Points' in results[p][q].keys(): 
                        for qp in r[p][q].keys(): results[p][q][qp] = r[p][q][qp]
                        counter[0]+=1
                    ## abort if there are no new points:
                    elif type(r[p][q]['Points']) == str or r[p][q]['Points'] == None:
                        pass
                    ## Overwrite if there are no points yet
                    elif type(results[p][q]['Points'])==str and not type(r[p][q]['Points'])==str:
                        for qp in overwrite: results[p][q][qp] = r[p][q][qp]
                        counter[0]+=1
                    ## Overwrite if there are no points yet
                    elif results[p][q]['Points']==None and not r[p][q]['Points']==None:
                        for qp in overwrite: results[p][q][qp] = r[p][q][qp]
                        counter[0]+=1
                    ## Ignore new info if points match.
                    elif results[p][q]['Points']==r[p][q]['Points']:
                        if 'State' in r[p][q]:
                            if r[p][q]['State'] > results[p][q]['State']:
                                for qp in overwrite: results[p][q][qp] = r[p][q][qp]
                                counter[1]+=1
                    ## Handle conflict if there is one
                    else:
                        if 'State' in r[p][q]:  ## Only if not new score is not automated
                            counter[2]+=1
                            conflicts.append((p,q,results[p][q]['Points'],r[p][q]['Points']))
                            results[p][q]['State']=1 ## Set conflict for this question
                            results[p][q]['Remarks']+=" // conflicting remarks: "+r[p][q]['Remarks']+' // Suggested Points: '+str(r[p][q]['Points'])
                   
                        
            outt = "Successfully added "+str(counter[0])+" new marks and "
            outt += str(counter[1])+" new remarks with "+str(counter[2])
            outt += " conflicts to results."
            messagebox.showinfo("Success",outt)

            self.data['Resp']=results
                
                        
        except Exception as e:
            messagebox.showerror("FAIL","Etwas ist schief gelaufen.\nDie Daten wurden nicht verändert.\n\nFehlermeldung:\n"+str(e))
            
        

    def read_new_data(self, ltype = "JSON"):
        global settings
        if ltype == "JSON":
            attributes = {'defaultextension':['.json'],
                          'filetypes':[('JSON','.json'),
                                       ('Other','.*')]}
        elif ltype == "CSV":
            attributes = {'defaultextension':['.json','.csv','.txt'],
                          'filetypes':[('Result File',['.xls','.csv','.txt']),
                                       ('JSON','.json'),
                                       ('Other','.*')]}
        elif ltype == "XLSX":
            attributes = {'defaultextension':['.xlsx'],
                          'filetypes':[('Result File',['.xlsx']),
                                       ('Other','.*')]}
        elif ltype == "ZIP":
            attributes = {'defaultextension':['.zip'],
                          'filetypes':[('QTI2.1 results',['.zip']),
                                       ('Other','.*')]}
            
        fname = filedialog.askopenfilename(**attributes)
        
        j = None

######  use ltype to really distinguish
######  for xlsx it is a 2-step process

        if ltype == "JSON": ## Pre-Used JSON File
            try:
                inf = open(fname,'r',encoding='utf-8',errors='ignore')
                j = eval(inf.readline())
                inf.close()
            except:
                print("invalid file")
                
        elif ltype == "CSV": ## OLD OLAT
            try:
                j = read_xls(fname)
            except:
                pass

            if type(j) == dict: ## Create the temporary file.
                o = open(settings['J_Filename'],'w',encoding='utf-8',errors='ignore')
                o.write(str(j))
                o.close()

        elif ltype == "ZIP":
            j = read_zip(fname)

        elif ltype == "XLSX":
            j = read_xlsx(fname)
            

        if type(j) == dict: ## Json constructed. Create the temporary file.
            o = open(settings['J_Filename'],'w',encoding='utf-8',errors='ignore')
            o.write(str(j))
            o.close()
            
        
        if type(j)==dict:  ##from here, all approaches must be the same
            try:
                self.m.frage.delete(0,len(self.questions)-1) ## Reset the question menu
            except:
                pass ## First start of the program. No questions to remove.
            self.data = j
            self.respondents = []
            self.questions = []

            for k in j['Resp'].keys():
                self.respondents.append(k)
            for q in j['Quest'].keys():
                self.questions.append(q)
                self.m.frage.add_command(label=str(q),command=CMD(self.change_q,q))
            settings['Curr_R']=0
            settings['Curr_Q']=0

            if ltype in ["XLSX","ZIP"]:
                for q in self.questions:
                    self.mark_question(q)

            self.refresh()
            
    def read_data(self): ## Initial set-up of data
        j = 0
        fn = ''

                
        try:
            inf = open(settings['J_Filename_Out'],'r',encoding='utf-8',errors='ignore')
            j = eval(inf.readline())
            fn = settings['J_Filename_Out']
            inf.close()  
        except:
            try:
                inf = open(settings['J_Filename'],'r',encoding='utf-8',errors='ignore')
                j = eval(inf.readline())
                fn = settings['J_Filename']
                inf.close()
            except:
                pass

        if type(j)==dict:
            if messagebox.askokcancel("Temporäre Datei laden?",'Es wurde eine gültige temporäre Datei mit Namen "{0}" gefunden. Möchten Sie diese Datei laden?\nWenn Sie auf "Abbrechen" klicken, können Sie andere Dateien laden'.format(fn)):
                self.data = j
                self.respondents = []
                self.questions = []
                for k in j['Resp'].keys():
                    self.respondents.append(k)
                for q in j['Quest'].keys():
                    self.questions.append(q)
                    self.m.frage.add_command(label=str(q),command=CMD(self.change_q,q))
                if len(self.respondents)<settings['Curr_R']+1:settings['Curr_R']=0
                if len(self.questions)<settings['Curr_Q']+1:settings['Curr_Q']=0

                return True
            else:
                return False
        else:
            return False
        

       

    def next(self,event=None):
        global settings
        if self.store(False):
            settings['Curr_R']+=1
            if not settings['Curr_R']<len(self.respondents):
                messagebox.showerror("Geht nicht weiter","Keine Weiteren Teilnehmer")
                settings['Curr_R']=len(self.respondents)-1
            self.refresh()

    def prev(self,event=None):
        global settings
        if self.store(False):
            settings['Curr_R']-=1
            if not settings['Curr_R']>=0:
                messagebox.showerror("Geht nicht weiter","Das ist der erste Teilnehmer")
                settings['Curr_R']=0
            self.refresh()

    def nextq(self,event=None):
        global settings
        if self.store(False):
            settings['Curr_Q']+=1
            if not settings['Curr_Q']<len(self.questions):
                messagebox.showerror("Geht nicht weiter","Keine Weiteren Fragen")
                settings['Curr_Q']=len(self.questions)-1
            self.refresh()

    def prevq(self,event=None):
        global settings
        if self.store(False):
            settings['Curr_Q']-=1
            if not settings['Curr_Q']>=0:
                messagebox.showerror("Geht nicht weiter","Das ist der erste Frage")
                settings['Curr_Q']=0
            self.refresh()


    def correctans(self,event=None):
        global settings
        q = self.data['Quest'][self.questions[settings['Curr_Q']]]

        h = 'Question-ID: {0} ({1}/{2})'.format(self.questions[settings['Curr_Q']],settings['Curr_Q'],len(self.questions))
        h+= '\nType:  '+q['Type']
        h+= '\nTitle: '+q['Title']
        h+= '\nQuestion:\n'
        
        self.ca = Toplevel(self)
        self.ca.rowconfigure(1, weight=1)
        self.ca.columnconfigure(1, weight=1)
        self.ca.title("Lösungsschema ändern")

        self.ca.header = Text(self.ca,width=120,height=14,bg="#ffffff",wrap=WORD)
        self.ca.header.tag_config('header',font=("Arial",11,"bold"))
        self.ca.header.grid(row=0,column=1,columnspan=2,sticky=W+E)
        self.ca.header.insert(END,h)
        self.ca.header.insert(END,q['Question'],('header'))
                        

        Label(self.ca,text="Maximale Punktzahl: ").grid(row=1,column=1,sticky=E)
        self.ca.mp = Entry(self.ca,width=10)
        self.ca.mp.grid(row=1,column=2,sticky=W)
        self.ca.mp.insert(END,str(q['Max']))

        Label(self.ca,text="\nKorrekte Antwort: ").grid(row=1,column=1,sticky=W)
        self.ca.f = Frame(self.ca, borderwidth=2, width=120, relief=SUNKEN)
        self.ca.f.grid(row=2,column=1,columnspan=2,sticky=W+E)

        Label(self.ca,text="\nKorrekturhinweis: ").grid(row=3,column=1,sticky=W)
        self.ca.advice = Text(self.ca,width=120,height=3,bg="#ffffff",wrap=WORD)
        self.ca.advice.grid(row=4,column=1,columnspan=2,sticky=W+E)
        self.ca.advice.insert(END,q['Advice'])
        
        Button(self.ca, text="OK", command=self.caconf, width=20).grid(row=10,column=1,sticky=W)
        Button(self.ca, text="Abbrechen", command=self.cakill, width=20).grid(row=10,column=2,sticky=W)

        if q['Type'] in ['KPRIM','KPR']:
            Label(self.ca.f,text="+").grid(row=0,column=1)
            Label(self.ca.f,text="-").grid(row=0,column=2)
            self.answers = []
            r = 1
            for a in sorted(q['Ans'].keys()):
                Label(self.ca.f,text=linebreaks(q['Ans'][a]['Item'],90),width=80,anchor=W, justify=LEFT).grid(row=r,column=0)
                choice = IntVar()
                choice.set(0)
                if q['Ans'][a]['Correct'] == '+':
                    choice.set(1)
                Radiobutton(self.ca.f,variable=choice,value=1).grid(row=r,column=1)
                Radiobutton(self.ca.f,variable=choice,value=0).grid(row=r,column=2)
                self.answers.append(choice)
                r+=1

        elif q['Type'] == 'SCQ':
            self.answers = []
            r = 1
            items = list(sorted(q['Ans'].keys()))
            self.answers = IntVar()
            self.answers.set(0)
            for i in range(len(items)):
                Label(self.ca.f,text=linebreaks(q['Ans'][items[i]]['Item'],90),width=80,anchor=W, justify=LEFT).grid(row=r,column=0)
                if q['Ans'][items[i]]['Correct'] == '1':
                    self.answers.set(i)
                Radiobutton(self.ca.f,variable=self.answers,value=i).grid(row=r,column=1)
                r+=1

        elif q['Type'] in ['MCQ','HS']:
            self.answers = []
            r = 1
            for a in sorted(q['Ans'].keys()):
                Label(self.ca.f,text=linebreaks(q['Ans'][a]['Item'],90),width=80,anchor=W, justify=LEFT).grid(row=r,column=0)
                choice = IntVar()
                choice.set(0)
                if q['Ans'][a]['Correct'] in [1,'1']:
                    choice.set(1)
                Checkbutton(self.ca.f,variable=choice).grid(row=r,column=1)
                self.answers.append(choice)
                r+=1

        elif q['Type'] == 'FIB':
            self.ca.header['height']=20
            self.answers = []
            r = 1
            for a in sorted(q['Ans'].keys()):
                Label(self.ca.f,text=linebreaks(q['Ans'][a]['Item'],90),width=80,anchor=W, justify=LEFT).grid(row=r,column=0)
                e = Entry(self.ca.f,width=20)
                e.grid(row=r,column=1)
                if q['Ans'][a]['Correct'] == None:
                    e.insert(END,q['Ans'][a]['Item'])
                else:
                    e.insert(END,str(q['Ans'][a]['Correct']))
                self.answers.append(e)
                r+=1


        elif q['Type'] == 'ESSAY': ## It's an essay
            self.answers = Text(self.ca.f,height=10)
            self.answers.grid(row=1,column=1,sticky=W+E)
            a = list(q['Ans'].keys())[0]
            if q['Ans'][a]['Correct'] == None:
                pass
            else:
                self.answers.insert(END,q['Ans'][a]['Correct'])

        elif q['Type'] in ['MATCH','HT']:
            self.answers = []
            r = 1
            for a in sorted(q['Ans'].keys()):
                Label(self.ca.f,text=linebreaks(q['Ans'][a]['Item'],40),width=40,anchor=W, justify=LEFT).grid(row=r,column=0)
                choice = StringVar()
                choice.set(q['Ans'][a]['Correct'])
                Entry(self.ca.f,width=80,textvariable=choice).grid(row=r,column=1)
                self.answers.append(choice)
                r+=1

        else:
            print(q['Type'])
            a/0
            
                

    def caconf(self,event=None):
        global settings
        q = self.data['Quest'][self.questions[settings['Curr_Q']]]
        mp = self.ca.mp.get()
        try:
            q['Max']=float(mp)
        except:
            pass
        
        if q['Type'] in ['KPRIM','KPR']:
            k = sorted(q['Ans'].keys())
            for i in range(len(k)):
                if self.answers[i].get()==0:
                    q['Ans'][k[i]]['Correct']='-'
                else:
                    q['Ans'][k[i]]['Correct']='+'
        elif q['Type'] == 'SCQ':
            k = sorted(q['Ans'].keys())
            for i in range(len(k)):
                if self.answers.get()==i:
                    q['Ans'][k[i]]['Correct']='1'
                else:
                    q['Ans'][k[i]]['Correct']='0'

        elif q['Type'] in ['MCQ','HS']:
            k = sorted(q['Ans'].keys())
            for i in range(len(k)):
                if self.answers[i].get()==0:
                    q['Ans'][k[i]]['Correct']='0'
                else:
                    q['Ans'][k[i]]['Correct']='1'

        elif q['Type'] == 'FIB':
            k = sorted(q['Ans'].keys())
            for i in range(len(k)):
                a = self.answers[i].get()
                if a[0]=='[':
                    try:
                        a=eval(a)
                    except:
                        pass
                q['Ans'][k[i]]['Correct']=a

        elif q['Type'] == 'ESSAY':
            a = self.answers.get(1.0,END)
            k = list(q['Ans'].keys())[0]
            q['Ans'][k]['Correct']=a

        elif q['Type'] in ['MATCH','HT']:
            k = sorted(q['Ans'].keys())
            for i in range(len(k)):
                q['Ans'][k[i]]['Correct'] = self.answers[i].get()

        else:
            a = self.answers.get(1.0,END)
            k = list(q['Ans'].keys())[0]
            q['Ans'][k]['Correct']=a

        adv = self.ca.advice.get(1.0,END)
        try:
            while adv[-1]=='\n':adv=adv[:-1] ## Remove trailing line breaks
        except:
            pass ## Reached a string of len 0
        q['Advice']=adv


        self.data['Quest'][self.questions[settings['Curr_Q']]] = q

        if q['Type'] in ['KPRIM','KPR','SCQ','MCQ','FIB','MATCH','HS']:
            result = self.mark_question(self.questions[settings['Curr_Q']]) ## Auto-Correct all the answers
            messagebox.showinfo('Result','Das Korrekturschema wurde angewandt, um alle Teilnehmer automatisch zu bewerten.\n'+result)
       
        self.cakill()
        self.refresh()
        self.store() ## Whenever new solutions are entered, store and update the temporary file.

    def cakill(self,event=None):
        self.ca.destroy()

                
    def mark_question(self,q):
        ## Mark one question for all respondents.

        quest = self.data['Quest'][q]
        plist = []

        for r in self.data['Resp'].keys():
            a = self.data['Resp'][r][q]['Ans']
            points = 0
            maxpoints = quest['Max']
            rem = self.data['Resp'][r][q]['Remarks'] ## Retain previous remarks

            if quest['Type']=='SCQ':
                for ans in quest['Ans'].keys():
                    if quest['Ans'][ans]['Correct'] in [1,'1'] and a[ans] in [1,'1']: ## Agreement on 1 (on any item)
                        points = maxpoints
                        
            elif quest['Type'] in ['MCQ','HS']:
                hits = 0
                misses = 0
                pstep = maxpoints/len(a.keys())
                for ans in quest['Ans'].keys():
                    if str(quest['Ans'][ans]['Correct']) == str(a[ans]):
                        hits+=1
                    else:
                        misses+=1
                points = (hits-misses)*pstep
                if points<0:points=0

            elif quest['Type'] in ['KPRIM','KPR']:
                points = maxpoints
                ded = float(points)/2
                for ans in quest['Ans'].keys():
                    if str(quest['Ans'][ans]['Correct']) == str(a[ans]):
                        pass
                    else:
                        points-=ded
                if points<0:points=0

            elif quest['Type'] == 'FIB':
                ac = True ## Only automatically mark if all is correct
                rem = '' ## Remove previous remarks. Otherwise, the remarks are stacked over each other.

                for ans in quest['Ans'].keys():
                    if type(quest['Ans'][ans]['Correct']) in (str,int,float):
                        if not str(a[ans]) == str(quest['Ans'][ans]['Correct']):
                            ac = False
                    elif type(quest['Ans'][ans]['Correct']) == list:
                        try:
                            antwort = float(a[ans])
                            if antwort < quest['Ans'][ans]['Correct'][0] or antwort > quest['Ans'][ans]['Correct'][1]:
                                ac = False
                                rem+=ans+': '+str(a[ans])+' nicht im Intervall '+ str(quest['Ans'][ans]['Correct'])+'\n'
                        except:
                            ac = False
                            rem+=ans+': '+str(a[ans])+' ist keine Zahl.\n'
                    else:
                        ac=False
                if ac:
                    points = maxpoints
                else:
                    points = self.data['Resp'][r][q]['Points'] ## Force input by corrector

            elif quest['Type'] == 'MATCH':
                points = 0
                ded = float(maxpoints)/len(quest['Ans'].keys())
                for ans in quest['Ans'].keys():
                    if str(quest['Ans'][ans]['Correct']) == str(a[ans]):
                        points+=ded
                    else:
                        points-=ded
                if points<0:points=0
                if not points==maxpoints: points=self.data['Resp'][r][q]['Points'] ## Prevent 0 points from being stored if something is wrong.
                

            else: ## Essay question
                points = self.data['Resp'][r][q]['Points'] ## Prevent 0 points from being stored
                pass ## No automated correction for essays


            overwrite = False
            if 'State' in self.data['Resp'][r][q].keys():
                if self.data['Resp'][r][q]['State']==-1:
                    overwrite = True ## Overwrite if automated correction
            else:
                overwrite = True ## Overwrite if no manual correction was done

            if overwrite:
                self.data['Resp'][r][q]['Remarks']=rem ## Only add remarks if no remarks are present
                if type(points)==str:
                    pass
                else: 
                    self.data['Resp'][r][q]['Points']=points
                    plist.append(points)
                    

        stat = desc_stat(plist)

        if type(stat)==dict:
            return "\nErgebnis:\n{4:.1f}% der Fälle bewertet\nMittelwert: {0:.2f}\nStabw: {1:.2f}\nRange: [{2:.1f}, {3:.1f}]".format(stat['M'],stat['SD'],stat['Min'],stat['Max'],stat['Share']*100)
        else:
            return 'Keine automatisch vergebenen Punkte'              

        
    def refresh(self):
        global settings
        r = self.respondents[settings['Curr_R']]
        q = self.questions[settings['Curr_Q']]       
        curr = self.data['Resp'][r][q]
        #print(curr)
        #print(baum_schreiben(self.data))

        tnd = str(r)+': '+self.data['Resp'][r]['Info']['Nachname']+', '+self.data['Resp'][r]['Info']['Vorname']

        self.cq['text'] = "  Aktuelle Frage: {0} ({1})".format(q,self.data['Quest'][q]['Title'])
        self.rid.set("Teilnehmer: '{0}' ({1}/{2})".format(tnd,settings['Curr_R']+1,len(self.respondents)))

        if not curr['Points'] == None:
            self.points.set(str(curr['Points']))
        else:
            self.points.set('')
        
        self.maxpoints.set("(Max: {0})".format(self.data['Quest'][q]['Max']))
        if type(curr['Points'])==str:
            self.pts["bg"]="#ffaa90"
        else:
            self.pts["bg"]="#ffffff"

        self.geg["state"]=NORMAL
        self.req["state"]=NORMAL

        given,required = niceprint(curr['Ans'],self.data['Quest'][q])
        
        self.geg.delete("1.0",END)
        self.geg.insert(END,given)
        
        self.req.delete("1.0",END)
        self.req.insert(END,required)
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

        #print(self.compare("1003715389",0,2))

        self.update()


    def change_q(self,question):
        global settings
        settings['Curr_Q']=self.questions.index(question)
        self.refresh()


    def check_plagiat(self):
        matrix = {}
        for r in self.respondents:
            matrix[r] = {}
            for r2 in self.respondents:
                matrix[r][r2] = []

        nr = len(self.respondents)
                
        for q in self.data['Quest'].keys():
            #print(q,self.data['Quest'][q]["Type"])
            if self.data['Quest'][q]["Type"]=="ESSAY":
                for i in range(1,nr):
                    for k in range(i):
                        agree = self.compare(q,i,k)
                        matrix[self.respondents[i]][self.respondents[k]].append(agree)
                        matrix[self.respondents[k]][self.respondents[i]].append(agree)

        for r in self.respondents:
            for r2 in self.respondents:
                if len(matrix[r][r2])>0:
                    matrix[r][r2] = sum(matrix[r][r2])/len(matrix[r][r2])
                else:
                    matrix[r][r2] = 0

        #print(matrix)


        width = 500
        height = 500

        if width < nr*20:
            width = nr*20
        if height < nr*20:
            height = nr*20

        self.case = StringVar()
        self.case.set("Maus über ein Feld, um Informationen zu sehen")

        self.infobox = Toplevel(self)
        self.infobox.rowconfigure(1, weight=1)
        self.infobox.columnconfigure(1, weight=1)
        self.infobox.title("Übereinstimmung in offenen Antworten")
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

        rowstep = (height-50)//nr
        rowt = []
        for i in range(nr):
            rowt.append(i*rowstep+25)

        colstep = (width-125)//nr
        colt = []
        for i in range(nr):
            colt.append(i*colstep+100)

        for i in range(nr):
            ypos = rowt[i]
            self.infobox.plot.create_text(50,(ypos+rowstep//2),text=self.respondents[i])
            for k in range(nr):
                xpos = colt[k]
                agree = matrix[self.respondents[i]][self.respondents[k]]
                #print(agree)
                
                col = heat_color(agree)
                #print(col)
                
                b = self.infobox.plot.create_rectangle(xpos,ypos,
                                                       xpos+colstep-1,
                                                       ypos+rowstep-2,
                                                       fill=col)
                #self.infobox.plot.tag_bind(b, '<Button-1>', CMD(self.jump,ri,qi))
                self.infobox.plot.tag_bind(b, '<Enter>', CMD(self.change_lab_plag,i,k,agree))
                        

    def compare(self,q,i,k, window=7):
        a1 = '\n'.join(list(self.data["Resp"][self.respondents[i]][q]["Ans"].values())).lower()
        a2 = '\n'.join(list(self.data["Resp"][self.respondents[k]][q]["Ans"].values())).lower()

        #print(a1,a2)

        d = {}
        hit = 0
        for i in range(len(a1)-window):
            d[a1[i:i+window]]=0
        for i in range(len(a2)-window):
            try:
                d[a2[i:i+window]]+=1
                hit+=1
                #print(a2[i:i+window])
            except:
                pass

        if min(len(a1),len(a2)) >0:
            agreement = hit / min(len(a1),len(a2))
        else:
            agreement = 0
        return agreement
            

    def check_completeness(self):
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

        xpadding = 175
        rstep = (width-xpadding-25)//nr
        rt = []
        for i in range(nr):
            rt.append(i*rstep+xpadding)

        for qi in range(nq):
            q = qt[qi]
            self.infobox.plot.create_text(7,(q+qstep//2),text=self.questions[qi],anchor=W)
            for ri in range(nr):
                r = rt[ri]
                item = self.data['Resp'][self.respondents[ri]][self.questions[qi]]
                if type(item['Points']) == str or item['Points']==None:
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


    def check_scores(self):
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
                item = self.data['Resp'][self.respondents[ri]][self.questions[qi]]
                #print(item)
                try:
                    pscore = float(item['Points'])/self.data['Quest'][self.questions[qi]]['Max']
                except:
                    pscore = ''

                if type(pscore)==float:
                    col = heat_color(pscore,'bw')
                    t = str(pscore) #str(item['Points'])
                
                    b = self.infobox.plot.create_rectangle(r,q,r+rstep-1,q+qstep-2,fill=col)
                    self.infobox.plot.tag_bind(b, '<Button-1>', CMD(self.jump,ri,qi))
                    self.infobox.plot.tag_bind(b, '<Enter>', CMD(self.change_lab,ri,qi,t))
                    

    def jump(self,resp,quest,event=""):
        #print(resp,quest)
        settings['Curr_R']=resp
        settings['Curr_Q']=quest

        self.infobox.destroy()
        self.refresh()

    def change_lab(self,resp,quest,problem,event=""):
        tn = self.respondents[resp]
        info = self.data['Resp'][tn]['Info']
        tnd = "{0}: {1}, {2}".format(tn,info['Nachname'],info['Vorname'])
        self.case.set("Teiln:" +tnd+' / Frage: '+str(self.questions[quest])+' ('+problem+')')


    def change_lab_plag(self,r1,r2,agreement,event=""):
        t1 = self.respondents[r1]
        t2 = self.respondents[r2]
        info1 = self.data['Resp'][t1]['Info']
        info2 = self.data['Resp'][t2]['Info']

        if t1==t2:
            tnd = "<{0}: {1}, {2}>".format(t1,info1['Nachname'],info1['Vorname'])
        else:
            tnd = "<{0}: {1}, {2}> agrees with <{3}: {4}, {5}> by {6:.2f}%".format(t1,info1['Nachname'],info1['Vorname'],
                                                                              t2,info2['Nachname'],info2['Vorname'],
                                                                              agreement*100)
        self.case.set(tnd)


    def store_data(self,fname=""):
        if fname == '':
            fname = filedialog.asksaveasfilename(**{'defaultextension':['.json'],
                                                  'filetypes':[('JSON','.json'),('Other','.*')]})

        try:
            outf = open(fname,'w',encoding='utf-8',errors='ignore')
            outf.write(str(self.data))
            outf.close()
            messagebox.showinfo("Success","Datei erfolgreich erstellt")
        except:
            messagebox.showerror("FAIL","Datei konnte nicht erstellt werden")
            
            

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

        if accept and (p > self.data['Quest'][q]['Max'] or p<0):
            accept=False
            messagebox.showinfo("Ungültige Eingabe", "Punkte müssen zwischen 0 und Maximum liegen.")

        elif accept:
            self.data['Resp'][r][q]['Points']=p
            self.data['Resp'][r][q]['Remarks']=rem
            self.data['Resp'][r][q]['State']=cb

        outf = open(settings['J_Filename_Out'],'w',encoding='utf-8',errors='ignore')
        outf.write(str(self.data))
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
    outdic = {'Ans':{},'Max':0,'Question':''}
    while scan:
        line+=1
        l = ls[line][:-1].split('\t')
        if len(l)<4:
            scan = False
        else: 
            if l[2]=='maxValue':
                try:
                    outdic['Max']=float(l[3])
                except:
                    pass
            elif l[2]=='':
                q = l[4]
                for tag in ['<p>','</p>','<em>','</em>','"']: q = q.replace(tag,'')
                q = q.replace("<br />",'\n')
                outdic['Question'] = q
            elif '_' in l[2]:
                item = l[4] ## Here is the item
                try:
                    answer = eval(l[5]) ## Fetch column F if there is a solution in it.
                except:
                    answer = None ## Answer is missing from the excel sheet. Native CSV document.
                    pass
                outdic['Ans'][l[2]] = {'Item':item,'Correct':answer}
                
            if ls[line][0]=='Q' or ls[line][:5]=='\t\t\t\t\t': ## End of this question
                ## This is just a security measure if someone opened and saved it in Excel.
                scan = False

    #print(baum_schreiben(outdic))

    return outdic


def read_table(lines):
    d = {}
    vlist=lines[0][:-1].split('\t')
    for v in vlist:
        d[v]=[]

    for l in lines[1:]:
        values = l[:-1].split('\t')
        ## Strip Quotation marks from values and put them in the dict
        for vi in range(len(vlist)):
            v = values[vi]
            if len(v)>1:
                if v[0] == '"': v = v[1:]
                if v[-1] == '"': v = v[:-1]
            ## Probably replace some shit if something occurs. (like tabs, bullet points...)
            d[vlist[vi]].append(v)
    return d

def heat_color(sval,mode='red'): ##Return a color value from a float in range [0;1]
    ##Modes: bw: Black and white (return a grayscale) 0=black
    ##       ibw: Grayscales with 0=white as highest value
    ##       red: Red-Green continuum
    
    ccode = "#ff0000"
    if mode == 'bw':
        v = int(sval * 255)
        if v > 255:
            v = 255
        if v < 0:
            v = 0

        hn = hex(v)[2:]
        if len(hn)==1:
            hn = '0'+hn
        ccode = '#' + hn + hn + hn
    elif mode == 'ibw':
        v = int(sval * 255)
        if v > 255:
            v = 255
        if v < 0:
            v = 0

        v = 255-v
        hn = hex(v)[2:]
        if len(hn)==1:
            hn = '0'+hn
        ccode = '#' + hn + hn + hn
    elif mode == 'red':
        gval = math.sin(sval*3.14159)/2+0.5
        bval = math.sin(sval*4+.9)/2+0.5
        rval = math.sin(sval*3-1)/2+0.5
        rv = int(rval * 255)
        if rv > 255:
            rv = 255
        if rv < 0:
            rv = 0
        gv = int(gval * 255)
        if gv > 255:
            gv = 255
        if gv < 0:
            gv = 0
        bv = int(bval * 200)
        if bv > 255:
            bv = 255
        if bv < 0:
            bv = 0


        rn = hex(rv)[2:]
        if len(rn)==1:
            rn = '0'+rn
        gn = hex(gv)[2:]
        if len(gn)==1:
            gn = '0'+gn
        bn = hex(bv)[2:]
        if len(bn)==1:
            bn = '0'+bn
        ccode = '#' + rn + gn + bn
        #ccode = '#' + bn + bn + bn 

    else:
        ccode = '#000000'
    return ccode


def desc_stat(l):
    numl = []
    for e in l:
        if type(e) in [int,float]: numl.append(float(e))
    if len(numl)>0:
        mean = sum(numl)/len(numl)
        sd = 0
        for e in numl:
            sd+=(e-mean)**2
        sd = (sd/len(numl))**.5
        
        return {'M':mean,'SD':sd,'Min':min(numl),'Max':max(numl),'Share':len(numl)/len(l)}
    else:
        return None

def grab(s,start,end):
    s1 = s.split(start)[1:]
    #print(type(s1),len(s1))
    s2 = [substr.split(end)[0] for substr in s1]
    if len(s2)==1: s2 = s2[0]
    #print(s2)
    return s2

def striptags(s):
    for n in ['<p>','<div>','</div>','<em>','</em>','\n']: s = s.replace(n,"")
    for n in ['</p>','<br/>']: s = s.replace(n,"\n")
    while '  ' in s: s=s.replace('  ',' ')
    if s[0]==' ':s = s[1:]
    if s[-1]==' ':s=s[:-1]
    if s[-1]=='\n':s=s[:-1]
    return s

def gather(zfile, questions):
    questiondic = {}
    for i in range(len(questions)):
        qf = zfile.open(questions[i]['XML'])
        lin = qf.readlines()
        qf.close()
        content = "\n".join([l.decode("utf-8") for l in lin])
##        print('\n\n--------------------------------\n',questions[i]['XML'],'\n')
##        print('Question Number: '+str(i+1))
        xname = questions[i]['XML']
        if '/' in xname:
            xname = xname[xname.rfind("/")+1:]

        ### Single Choice

        if xname[:2].lower() == 'sc':
            qtype = "SC"
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qend = qcont.find("<choiceInter")
            qcont, qitems = qcont[:qend], qcont[qend:]
            qcont = striptags(qcont)
            qresp = grab(content,"<correctResponse>","</correctResponse")
            qresp = grab(qresp,'<value>','</value>')

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass
           
            items = grab(qitems,'<simpleChoice','</simpleChoice>')
            rorder = []
            rdic = {}
            for qi in items:
                rid = grab(qi,'identifier="','"')
                rc = grab(qi,'<p>','</p>')
                if type(rc)==list:rc='\n'.join(rc)
                rdic[rid]=rc
                rorder.append(rid)

            
            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'SCQ'}
            for ai in range(len(rorder)):
                alab = f"{i+1}_C{ai}"
                c = 0
                if rorder[ai] == qresp: c=1
                questiondic[qkey]['Ans'][alab]={'Correct':c,
                                                'Item':rdic[rorder[ai]]}

        ### Multiple Choice

        elif xname[:2].lower() == 'mc':
            qtype = "MC"
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qend = qcont.find("<choiceInter")
            qcont, qitems = qcont[:qend], qcont[qend:]
            qcont = striptags(qcont)
            qresp = grab(content,"<correctResponse>","</correctResponse")
            qresp = grab(qresp,'<value>','</value>')
            if type(qresp)==str:qresp=[qresp]

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass

            items = grab(qitems,'<simpleChoice','</simpleChoice>')
            rorder = []
            rdic = {}
            for qi in items:
                rid = grab(qi,'identifier="','"')
                rc = grab(qi,'<p>','</p>')
                if type(rc)==list:rc='\n'.join(rc)
                rdic[rid]=rc
                rorder.append(rid)

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'MCQ'}
            for ai in range(len(rorder)):
                alab = f"{i+1}_C{ai}"
                c = 0
                if rorder[ai] in qresp: c=1
                questiondic[qkey]['Ans'][alab]={'Correct':c,
                                                'Item':rdic[rorder[ai]]}

        ### HOTSPOT

        elif xname[:7].lower() == 'hotspot':
            qtype = "HS"
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qend = qcont.find("<hotspotInter")
            qcont, qitems = qcont[:qend], qcont[qend:]
            qcont = striptags(qcont)
            qresp = grab(content,"<correctResponse>","</correctResponse")
            qresp = grab(qresp,'<value>','</value>')
            if type(qresp)==str:qresp=[qresp]

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass

            items = grab(qitems,'<hotspotChoice','/>')
            rorder = []
            rdic = {}
            for qi in items:
                rid = grab(qi,'identifier="','"')
                rc = grab(qi,'coords="','"')
                rt = grab(qi,'shape="','"')
                rc = f"{rt} ({rc})"
                if type(rc)==list:rc='\n'.join(rc)
                rdic[rid]=rc
                rorder.append(rid)

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'HS'}
            for ai in range(len(rorder)):
                alab = f"{i+1}_C{ai}"
                c = 0
                if rorder[ai] in qresp: c=1
                questiondic[qkey]['Ans'][alab]={'Correct':c,
                                                'Item':rdic[rorder[ai]]}


        ### KPRIM

        elif xname[:3].lower() == 'kpr':
            qtype = "KPRIM"
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qend = qcont.find("<matchInter")
            qcont, qitems = qcont[:qend], qcont[qend:]
            qcont = striptags(qcont)
            qresp = grab(content,"<correctResponse>","</correctResponse")
            qresp = grab(qresp,'<value>','</value>')
            qresp = dict([tuple(r.split(' ')) for r in qresp])

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass


            items = grab(qitems,'<simpleAssociableChoice','</simpleAssociableChoice>')
            rorder = []
            rdic = {}
            for qi in items[:4]:
                rid = grab(qi,'identifier="','"')
                rc = grab(qi,'<p>','</p>')
                if type(rc)==list:rc='\n'.join(rc)
                rdic[rid]=rc
                rorder.append(rid)

            #print(rdic,rorder)

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'KPRIM'}
            for ai in range(len(rorder)):
                alab = f"{i+1}_KP{ai+1}"
                c = '-'
                if qresp[rorder[ai]] =='correct': c = '+'
                questiondic[qkey]['Ans'][alab]={'Correct':c,
                                                'Item':rdic[rorder[ai]]}
            

        ### Lückentext

        elif xname[:3].lower() == 'fib':
            qtype = 'FIB'
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qresp = grab(content,"<responseDeclaration","</responseDeclaration")
            #print('\n'.join(qresp))

            blanks = re.findall('<textEntry.*?/>',qcont)
            rorder = []

            for blank in blanks:
                qcont = qcont.replace(blank,"[...]")
                rorder.append(grab(blank,'Identifier="','"'))
            qcont = striptags(qcont)

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass
                    
            rdic = {}
            for r in qresp:
                rid = grab(r,'identifier="','"')
                if len(rid)==0:rid='<>'
                if type(rid)==list:rid=rid[0]
                rc = grab(r,'<value>','</value>')
                rdic[rid] = rc

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'FIB'}
            for ai in range(len(rorder)):
                alab = f"{i+1}_FIB{ai+1}"
                try:
                    questiondic[qkey]['Ans'][alab]={'Correct':rdic[rorder[ai]],
                                                    'Item':''}
                except:
                    questiondic[qkey]['Ans'][alab]={'Correct':'',
                                                    'Item':''}

        ### HOTTEXT


        elif xname[:7].lower() == 'hottext':
            qtype = 'HT'
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qresp = grab(content,"<responseDeclaration","</responseDeclaration")
            qresp = grab(qresp,"<value>","</value>")
            if type(qresp)==str:qresp=[qresp]
            qcont = grab(qcont,'<p>','</p>')
            if type(qcont)==list:
                qcont='\n'.join(qcont)
            #print('\n'.join(qresp))

            blanks = re.findall('<hottext .*?/hottext>',qcont)
            rorder = []

            #print("Blanks",blanks)

            for blank in blanks:
                qcont = qcont.replace(blank," [...] ")
                rorder.append(grab(blank,'identifier="','"'))
            qcont = striptags(qcont)

##            print(rorder)
##            print(qcont)

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass

            #print(mscore)
                    
            rdic = {}
            for r in blanks:
                #print(r)
                rid = grab(r,'identifier="','"')
                rc = grab(r,'>','<')[0]
                rdic[rid] = rc
                #print(rid,rc)

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'HT'}
            #print(qresp)
            for ai in range(len(rorder)):
                alab = f"{i+1}_HT{ai+1}"
                c = ''
                if rorder[ai] in qresp: c = rdic[rorder[ai]]
                questiondic[qkey]['Ans'][alab]={'Correct':c,
                                                'Item':rdic[rorder[ai]]}



        #### Some kind of matching (Matrix, Drag&Drop)

        elif xname[:5].lower() == 'match':
            qtype = 'MATCH'
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qend = qcont.find("<matchInter")
            qcont, qitems = qcont[:qend], qcont[qend:]
            qcont = striptags(qcont)
            qresp = grab(content,"<correctResponse>","</correctResponse")
            qresp = grab(qresp,'<value>','</value>')
            qitems = grab(qitems,"<simpleMatchSet","</simpleMatchSet")
            if type(qresp)==str:qresp=[qresp]
##            print(qresp)
##            print(qitems)
##            print(len(qitems))

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass
            #print(mscore)

            items1 = grab(qitems[0],'<simpleAssociableChoice','</simpleAssociableChoice>')
            items2 = grab(qitems[1],'<simpleAssociableChoice','</simpleAssociableChoice>')
##            print('1: ',items1)
##            print('2: ',items2)
            rorder = []
            rdic = {}
            for qi in items2:
                rid = grab(qi,'identifier="','"')
                rc = grab(qi,'<p>','</p>')
                #print(rid,rc)
                if type(rc)==list:rc='\n'.join(rc)
                rdic[rid]=rc
                rorder.append(rid)

            qcont += '\n\nOptions:\n'
            for qi in items1:
                rid = grab(qi,'identifier="','"')
                rc = grab(qi,'<p>','</p>')
                qcont+=f'\n - {rid}: {rc}'

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'MATCH'}
            for ai in range(len(rorder)):
                alab = f"{i+1}_K{ai+1}"
                c = ''
                sol =[]
                for qr in qresp:
                    if rorder[ai] in qr:
                        sol.append(qr.split(' ')[0])
                    c = ', '.join(sol)
                questiondic[qkey]['Ans'][alab]={'Correct':c,
                                                'Item':rdic[rorder[ai]]}


        ### ESSAY


        elif xname[:5].lower() == 'essay':
            qtype = 'ESSAY'
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qcont = grab(qcont,'<p>','</p>')
            if type(qcont)==list:qcont = '\n'.join(qcont)
            qcont = striptags(qcont)
            #print(qcont)

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass
            #print(mscore)

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{f"{i+1}_U1":{'Correct':'Please click to enter a solution',
                                                     'Item':'Text'}},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'ESSAY'}

        ### DRAWING

        elif xname[:5].lower() == 'drawi':
            qtype = 'DRAW'
            qtitle = grab(content,' title="','"')
            if type(qtitle) == list: qtitle = qtitle[0]
            qcont = grab(content,"<itemBody>","</itemBody>")
            qcont = grab(qcont,'<p>','</p>')
            if type(qcont)==list:qcont = '\n'.join(qcont)
            qcont = striptags(qcont)
            #print(qcont)

            scores = grab(content,'<outcomeDeclaration','</outcomeDeclaration>')
            mscore = ''
            for s in scores:
                if 'MAXSCORE' in s and mscore=='':
                    try:
                        mscore = float(grab(s,'<value>','</value>'))
                    except:
                        pass
            #print(mscore)

            qkey = f"Question #{i+1} ({qtype}): {qtitle}"

            questiondic[qkey] = {'Advice':'',
                                 'Ans':{f"{i+1}_U1":{'Correct':'Can not handle graphics. Refer to external material',
                                                     'Item':'Text'}},
                                 'Line':questions[i]['XML'],
                                 'Max':mscore,
                                 'Question':qcont,
                                 'Title':qtitle,
                                 'Type':'DRAW'}

        else:

            print("ERROR: INVALID: ",xname)

    #print('\n'.join(list(questiondic.keys())))

    return questiondic
        
def parsetable(xml,refs):
    rows = grab(xml,'<row','</row>')
    #print(len(rows))
    ## Create an empty dataset
    cols = []
    for first in ['','A','B','C','D','E','F','G','H','I','J','K','L','M']:
        for last in [chr(n) for n in range(65,91)]:
            cols.append(first+last)
    data = {}
    for c in cols:
        data[c] = (len(rows)+1)*['']

    limit = 50
    i = 0

    bcols = dict(list(zip(cols,[False]*len(cols))))
    
    for r in rows:
        cells = grab(r,'<c','</c>')
        for c in cells:
            name = grab(c, 'r="','"')
            col = ''
            row = ''
            for ch in name:
                if ch.isalpha():
                    col+=ch
                else:
                    row+=ch
            sval = False
            try:
                bcols[col]=True
            except:
                print("Column not found: ",col)
            if 't="s"' in c: sval = True ## Contains string value
            val = grab(c,'<v>','</v>')
            try:
                val = int(val)
                if sval:
                    val = refs[val]
            except:
                try:
                    val = float(val)
                except:
                    pass
                
            try:
                data[col][int(row)] = val
            except:
                print('ERROR grabbing cell in parsetable',name,val,col,row)

    cn = len(cols)-1
    beef = False
    while not cn == 0 and not beef:
        if bcols[cols[cn]] == False:
            del data[cols[cn]]
            cn = cn-1
        else:
            beef = True
    return data


def qti_reader(fname):
    z = zipfile.ZipFile(fname,mode="r")
    files = z.namelist()

    testfile = ''
    for f in files:
        if f[:4] == "test": testfile=f

    tf = z.open(testfile)
    lin = tf.readlines()
    tf.close()

    content = "\n".join([l.decode("utf-8") for l in lin])
    q1 = content.split("<assessmentItemRef")[1:]
    q2 = [q.split("/>")[0] for q in q1]

    questions = []
    for i in range(len(q2)):
        ident = q2[i].split('identifier="')[1].split('"')[0]
        fname = q2[i].split('href="')[1].split('"')[0]
        questions.append({"Idx":i+1, "Ident":ident, "XML":fname})

    outdic = gather(z,questions)
    return outdic    

def xparse(fname):
    z = zipfile.ZipFile(fname,mode="r")
    files = z.namelist()
    #print(files)

    sref = 'xl/sharedStrings.xml'
    sheet = 'xl/worksheets/sheet1.xml'

    sr = z.open(sref)
    lin = sr.readlines()
    sr.close()
    content = "\n".join([l.decode("utf-8") for l in lin])
    refs = grab(content,'<si>','</si>')
    refs = [r.replace('<t>','').replace('</t>','') for r in refs]
    
    sr = z.open(sheet)
    lin = sr.readlines()
    sr.close()
    content = "\n".join([l.decode("utf-8") for l in lin])

    data = parsetable(content,refs)

    outdata = {}
    num = 1
    for v in data.keys():
        if data[v][2] == '':
            vname = f"Var_{num}"
            num+=1
        else:
            vname = data[v][2]
        outdata[vname] = data[v][3:]

    #print(outdata)

    return outdata
      
def create_gd(resp,questions, keys = ['Laufnummer', "Vorname","Nachname","Benutzername","Matrikelnummer"]):
    gd = {'Resp':{},'Quest':questions}
    
    for i in range(len(resp[keys[0]])): ## Use the sequence number
        r = resp[keys[0]][i] ## Identifier of this person
        add = []
        for a in keys[1:]:
            try:
                add.append(str(resp[a][i]))
            except:
                pass ## If there is no name or matrikelnummer
        gd['Resp'][r]={}
        
        if len(add)>0:
            gd['Resp'][r]['Info'] = {'ID':" ({0})".format(', '.join(add))}
        else:
            gd['Resp'][r]['Info'] = {'ID':" Anonymous"}

        for k in keys:
            try:
                gd['Resp'][r]['Info'][k] = resp[k][i]
            except:
                gd['Resp'][r]['Info'][k] = "-"
        
        for q in questions.keys():
            response = {}
            for a in questions[q]['Ans'].keys():
                response[a] = str(resp[a][i])
                #print(response[a])
                if '\\n' in response[a]:
                    response[a] = response[a].replace('\\n','\n') ## Rectify newlines
                    response[a] = response[a].replace('\\t','\t') ## Rectify Tabstopps
                    response[a] = response[a].replace('\\r','')   ## Remove carriage returns
            gd['Resp'][r][q] = {'Ans':response,'Points':None,'Remarks':''}
            
    return gd

def read_zip(fname):
    ## Reading a QTI2.1 result ZIP
    z = zipfile.ZipFile(fname,mode="r")
    files = z.namelist()

    testfile = ''
    xfile = ''
    for f in files:
        fin = re.findall("Resultate.*test.*test",f)
        if len(fin)>0:
            testfile = f
        if f[-5:]=='.xlsx':
            xfile=f

    tf = z.open(testfile)
    lin = tf.readlines()
    tf.close()
    path = testfile[:testfile.rfind('/')+1]

    content = "\n".join([l.decode("utf-8") for l in lin])
    q1 = content.split("<assessmentItemRef")[1:]
    q2 = [q.split("/>")[0] for q in q1]

    questions = []
    for i in range(len(q2)):
        ident = q2[i].split('identifier="')[1].split('"')[0]
        fname = q2[i].split('href="')[1].split('"')[0]
        questions.append({"Idx":i+1, "Ident":ident, "XML":path+fname})

    q = gather(z,questions)

    #print(baum_schreiben(q))

    ## Now load the XLSX with all answers:
    
    xcontents = BytesIO(z.read(xfile)) 
    r, k = read_xlsx(xcontents,False)
    #print(baum_schreiben(r))
    #print(k)
    fulldata = create_gd(r,q,k)
    
    return fulldata
    

def read_xlsx(fname="Eingabe.xlsx",addzip=True):
    keys = ['Laufnummer', "Vorname","Nachname","Benutzername","Matrikelnummer"]
    eng_keys = ['Sequence number','First name', 'Last name', 'User name','Matriculation nr.']

    resp = xparse(fname)
    #print(resp.keys())
    
    for i in range(len(keys)): ## Fetch English column names
        if not keys[i] in resp.keys() and eng_keys[i] in resp.keys():
            resp[keys[i]] = resp[eng_keys[i]] ## Translate all to German

    ## Replace x and '' by 1 and 0

    for v in resp.keys():
        if '_C' in v:
            nc = []
            for i in range(len(resp[v])):
                if resp[v][i]=='x':
                    nc.append(1)
                elif resp[v][i]=='':
                    nc.append(0)
                else:
                    print(f"The column {v} is not a dummy column. It contains other values than 'x' and ''")
            if len(nc)==len(resp[v]):resp[v] = list(nc)

    msg = f"Daten von {len(resp[v])} Befragten erfolgreich geladen."

    if addzip:
        msg+= "\n\nIn der Datei waren keine Informationen zum Inhalt der Fragen vorhanden.\n"
        msg+= "Für die Weiterarbeit muss die Information über die Inhalte der Prüfung geladen werden.\n"
        msg+= "\nDie Prüfungsinformationen sind in einer ZIP-Datei, die als qti-Export gekennzeichnet ist.\n"
        msg+= "Bitte laden Sie jetzt diese Datei, die zu dieser Prüfung gehört."
        
        messagebox.showinfo("Zweiter Schritt",msg)

                
        attributes = {'defaultextension':['.zip'],
                      'filetypes':[('QTI2.1 dump',['.zip']),
                                   ('Other','.*')]}      
        fname = filedialog.askopenfilename(**attributes)

        questions = qti_reader(fname)

        fulldata = create_gd(resp,questions,keys)
            
        return fulldata
    else:
        return resp, keys
    


def read_xls(fname = "Eingabe.xls"):
    keys = ['Laufnummer', "Vorname","Nachname","Benutzername","Matrikelnummer"]
    eng_keys = ['Sequence number','First name', 'Last name', 'User name','Matriculation nr.']

    inf = open(fname,'r',
           encoding="utf-8",errors='ignore')
    ls = inf.readlines()
    inf.close()

    tstart = 1
    tend = 2
    while not ls[tend][:3] in ['\n','\t\t\t']:
        tend+=1 ## Find end of the table of results (and beginning of questions)

    resp = read_table(ls[tstart:tend])

    for i in range(len(keys)): ## Fetch English column names
        if not keys[i] in resp.keys() and eng_keys[i] in resp.keys():
            resp[keys[i]] = resp[eng_keys[i]] ## Translate all to German
    
    #print(list(resp.keys()))

    questions = {}  ## Dictionary of all questions and solutions
    qkeys = {} ## Hash table to get the question from the answers identifier.
    for i in range(tend+2,len(ls)):
        l=ls[i]
        if l[0]=='Q':
            line = l.split('\t')
            q = line[0].split(':') ## First cell
            if len(q)==3:
                questions[q[2]]={'Line':i,'Type':q[1]}
                ans = scan_question(ls,i) ## Fetch remainder of the question
                questions[q[2]]['Ans']=ans['Ans']
                questions[q[2]]['Max']=ans['Max']
                questions[q[2]]['Question']=ans['Question']
                t = line[1]
                if t[0]== '"':t = t[1:]
                if t[-1]== '\n': t = t[:-1]
                if t[-1]== '"': t = t[:-1]
                questions[q[2]]['Title']=t ## Title without Quotes
                questions[q[2]]['Advice']=''
                
                for a in ans['Ans'].keys():
                    qkeys[a]=q[2] ## record keys in hash table

    fulldata = create_gd(resp,questions,keys)
    return fulldata


def niceprint(a,q):
    ## Add comments on correction
    ## Add a cleaning device to clean up special characters above 255 (Just scan and remove the suckers)
    valid = True
    anslist = []
    reqlist = []
    for ans in sorted(a.keys()):
        anslist.append(str(a[ans]))
        if not q['Ans'][ans]['Correct'] == None:
            reqlist.append(str(q['Ans'][ans]['Correct']))

    if q['Type'] in ['SCQ','MCQ','KPRIM','KPR']:
        answer = '   '.join(anslist)
        required = '   '.join(reqlist)
    elif q['Type'] in ['FIB']:
        answer = '\n'.join(anslist)
        required = '\n'.join(reqlist)
    else: ## Essay
        answer = '\n'.join(anslist)
        required = '\n'.join(reqlist)

    if 'Advice' in q.keys(): ## Korrekturhinweis
        if len(q['Advice'])>1:
            required+='\n\nHinweis zur Korrektur:\n----------------------\n'+str(q['Advice'])

    invalid = []
    for c in answer+required:
        if ord(c)>255:
            invalid.append(c)
    if len(invalid)>0:
        for c in invalid:
            answer = answer.replace(c,'<{0}>'.format(ord(c)))
            required = required.replace(c,'<{0}>'.format(ord(c)))

    if len(reqlist)==len(anslist):
        return answer,required
    else:
        return answer,"No valid answer.\n\nClick here to add one"

def linebreaks(line, maxlen=100):
    tokens = line.split(" ")
    lines = []
    curline = []
    curlen = 0
    for t in tokens:
        if curlen + len(t) > maxlen:
            lines.append(curline)
            curline = [t]
            curlen = len(t) + 1
        else:
            curline.append(t)
            curlen = curlen + len(t) + 1
    lines.append(curline)

    for i in range(len(lines)):
        lines[i] = " ".join(lines[i])

    line = "\n".join(lines)
    
    return line

def create_einsicht(res,r,folder='.\\',veranstaltung='',questions=[]): ## Takes one respondent dict
    disclaimer = " ".join(["In diesem Dokument sehen Sie für alle Fragen der Prüfung",
                            "sowohl Ihre eigene Antwort als auch die Antwort, die für",
                            "diese Frage als korrekt gewertet wurde.",
                            "Darunter sind Ihre Punkte, sowie die maximale Punktzahl",
                            "für die entsprechende Aufgabe notiert.\n",
                            "Falls es Bemerkungen gibt, sind diese am Ende jeder Frage",
                            "notiert. Die Bemerkungen wurden zum Teil automatisch erstellt",
                            "und beziehen sich auf die Fehler, die zu Abzügen geführt haben.",
                            "\n\n",
                            "Auf jeder Seite wird genau eine Frage dargestellt. Prüfen Sie",
                            "jeweils, ob die Frage korrekt bewertet wurde und ob Sie einen",
                            "allfälligen Abzug von Punkten nachvollziehen können.\n\n",
                            "Pro Fehler bei einer Multiple Choice / KPRIM  Frage wurden 50% der",
                            "Maximalpunktzahl abgezogen. Single Choice können nur das maximum",
                            "oder 0 Punkte geben. Es gibt bei keiner Aufgabe weniger als 0 Punkte."
                            "Bei Essay und offenen Fragen sind die Abzüge jeweils in den Kommentaren"
                            "erläutert.\n"
                            "Falls Sie mit einer Entscheidung nicht einverstanden sind, steht",
                            "Ihnen das Mittel eines Wiedererwägungsgesuchs zur Verfügung. Stellen",
                            "Sie ein solches Gesuch zu Handen Ihres Dozenten und beschreiben Sie",
                            "darin möglichst genau, bei welcher Frage Sie aus welchem Grund die",
                            "Punktevergabe ungerechtfertigt finden.",
                            "Das Gesuch wird dann geprüft und Sie erhalten eine schriftliche Rückmeldung.",
                            "\n\n\n\n",
                            "Anmerkung: Dieses Dokument wurde automatisch erstellt und kann ein seltsames",
                            "Layout haben. Die Informationen (Punkte / Bemerkungen) wurden bei",
                            "der Notenvergabe genau so verwendet."])

    p = 0
    mp = 0
    for q in res['Resp'][r].keys():
        if not q=='Info': ## This would be the personal information of the respondent
            try:
                p+=res['Resp'][r][q]['Points']
            except:
                p=''
            try:
                mp+=res['Quest'][q]['Max']
            except:
                mp=''

    rinf = res['Resp'][r]['Info']
    rname = "{0}, {1} ({2})".format(rinf['Nachname'],rinf['Vorname'],rinf['Matrikelnummer'])

    fname = folder + 'Resultate_'+ r + '.pdf'
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", size=16)
    pdf.set_fill_color(200,200,255)
    header = "Dokument für die Prüfungseinsicht\n"+veranstaltung+"\nTeilnehmer*in: "+rname
    pdf.multi_cell(0, 12, txt=header,
                   align="C",border=1,fill=True)

    pdf.set_font("Arial", size=14,style="B")
    pdf.y = pdf.y+15
    pdf.multi_cell(0,6, txt="Gesamtergebnis")
    if type(p)==str or type(mp)==str:
        pdf.set_font("Arial", size=12)
        pdf.y = pdf.y+5
        pdf.multi_cell(0,6, txt="Es kann kein Ergebnis ausgegeben werden. Die Korrektur für diese Prüfung ist noch nicht abgeschlossen")
    else:
        pdf.set_font("Arial", size=12,style="B")
        pdf.y = pdf.y+5
        pdf.multi_cell(0,6, txt="Sie haben {0:.1f} von insgesamt {1:.1f} Punkten erreicht.".format(p,mp))
        

    pdf.set_font("Arial", size=14,style="B")
    pdf.y = pdf.y+15
    pdf.multi_cell(0,6, txt="Erläuterung zu diesem Dokument")

    pdf.set_font("Arial", size=12)
    pdf.y = pdf.y+5
    pdf.multi_cell(0,6, txt=disclaimer)

    pdf.y = pdf.y+20
    pdf.set_font("Arial", size=10, style="I")
    pdf.multi_cell(0,6, txt="Erstellt am: "+str(time.ctime()))

    qnr = 0

    if questions == []:questions = sorted(res['Quest'].keys())


    for q in questions:
        qnr+=1
        given,required = niceprint(res['Resp'][r][q]['Ans'],res['Quest'][q])
        pdf.add_page()
        pdf.set_font("Arial", size=16)
        pdf.set_fill_color(220,220,255)
        t = str(qnr)+". Aufgabe: {0} ({1})\nMaximale Punktzahl: {2} / Typ: {3}".format(q,res['Quest'][q]['Title'],
                                                                                       res['Quest'][q]['Max'],
                                                                                       res['Quest'][q]['Type'])
        pdf.multi_cell(0, 12, txt=t, align="C",border=1,fill=True)

        pdf.set_fill_color(220,250,220)

        pdf.set_font("Arial", size=14,style="B")
        pdf.y+=20
        pdf.multi_cell(0,6, txt="Korrekte Antwort:")
        pdf.y+=5
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0,6,txt=required,fill=True,align="C",border=1)

        pdf.set_fill_color(220,220,220)

        pdf.set_font("Arial", size=14,style="B")
        pdf.y+=20
        pdf.multi_cell(0,6, txt="Ihre Antwort:")
        pdf.y+=5
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0,6,txt=given,fill=True,align="C",border=1)

        if len(res['Resp'][r][q]['Remarks'])>1:
            pdf.set_font("Arial", size=14,style="B")
            pdf.y+=20
            pdf.multi_cell(0,6, txt="Bemerkungen:")
            pdf.y+=5
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0,6,txt=res['Resp'][r][q]['Remarks'].replace(" // ",'\n'),fill=False,border=1)

        pdf.set_font("Arial", size=14,style="B")
        pdf.y+=20
        pdf.multi_cell(0,6, txt="Sie haben {0:.1f} von {1:.1f} Punkten erhalten.".format(res['Resp'][r][q]['Points'],res['Quest'][q]['Max']))

    inv = open(folder+'_inventory.txt','a',encoding="utf-8",errors="ignore")
    rnr = r.split(' ')[0]
    
    try:
        pdf.output(fname)
        inv.write(r+'\t'+rname+'\t'+fname+'\n')
        inv.close()
    except Exception as f:
        er = open(fname+"_ERROR.txt","w",encoding="utf-8",errors="ignore")
        er.write("Error. Could not export case:"+r+"\nThere must be a strange special character in this respondent's answers. Please mend manually.\n")
        er.close()
        inv.write(r+'\t'+rname+'\t'+str(f)+'\n')
        inv.close()


def write_results(result = None, pdf=True, table="Punktetabelle.xls",questions=[],respondents=[]):
    if respondents == []: respondents = sorted(list(result['Resp'].keys()))
    if questions == []: questions = sorted(list(result['Quest'].keys()))

    outf = open(table,'w',encoding='utf-8',errors='ignore')
    outf.write('\t'.join(['Laufnummer','Name','Vorname','Benutzername','Matrikelnummer','Total','Remarks']+questions)+'\n')
    for r in respondents:
        info = result['Resp'][r]['Info']
        line = [str(info['Laufnummer']),
                info['Nachname'],
                info['Vorname'],
                info['Benutzername'],
                info['Matrikelnummer'],'','']
        
        punkte = 0
        flagged = False
        for q in questions:
            if not result['Resp'][r][q]['Points'] == None:
                line.append(str(result['Resp'][r][q]['Points']))
            else:
                #print(result['Resp'][r][q])
                line.append('')
            try:
                punkte+=result['Resp'][r][q]['Points']
            except:
                punkte = ''
            try:
                if result['Resp'][r][q]['State']==1:flagged=True
            except:
                pass
            line[5]=str(punkte)
            if flagged:
                line[6]='Achtung: Mindestens eine der Fragen ist noch nicht geprüft. '
        if punkte == '':
            line[6]+='Noch nicht alle Punkte vergeben'
        #print(line)
        outf.write('\t'.join(line)+'\n')
    outf.close()

    if pdf:
        pdf = settings['PDF']

    if pdf:
        try:
            inv = open('_inventory.txt','w',encoding="utf-8",errors="ignore")
            inv.write('Laufnr.\tResp\tFile\n')
            inv.close()
            titel = simpledialog.askstring("Ausgabetitel","Welche Überschrift sollen die Dokumente für die Prüfungseinsicht tragen?\n\nWenn Sie dieses Textfeld leer lassen, werden keine PDFs erstellt.\n\n(z.B: 'Vorlesung: Statistik und Datenanalyse')")
            if len(titel)>2:
                for r in respondents:
                    create_einsicht(result,r,veranstaltung=titel,questions=questions)
        except:
            print("could not produce PDFs.")

    messagebox.showinfo("Export erfolgreich","Alle Daten wurden erfolgreich exportiert")
                    

global settings
settings = {}
settings['Curr_R']=0
settings['Curr_Q']=0
settings['Insecurities']=[]
settings['E_Filename']="Eingabe2.xls"
settings['J_Filename']="Korrektur.json" ## This filename could be set to some other name to store the raw input after reading an xls/csv-file.
settings['J_Filename_Out']="Korrektur.json"
settings['PDF']=pdf


root = Tk()
korr = Anzeige(root)
root.mainloop()
