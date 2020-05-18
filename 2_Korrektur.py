
import time
from tkinter import *
from tkinter import messagebox

class CMD: #Auxilliary function for callbacks using parameters. Syntax: CMD(function, argument1, argument2, ...)
    def __init__(s1, func, *args):
        s1.func = func
        s1.args = args
    def __call__(s1, *args):
        args = s1.args+args
        s1.func(*args)

class Anzeige(Frame):
    def __init__(self, master=None, daten = {'r':{'q':{}}}):        
        Frame.__init__(self,master)
        top=self.winfo_toplevel() #Flexible Toplevel of the window
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
        self.grid(sticky=N+S+W+E)

        self.resp = daten

        self.respondents = []
        self.questions = []
        for k in daten.keys():
            self.respondents.append(k)
        for q in daten[self.respondents[0]].keys():
            self.questions.append(q)

        self.set_window()

    def set_window(self):
        global settings

        self.m = Menubutton(self,text="Andere Frage",relief=RAISED)
        self.m.grid(row=0,column=0,sticky=N+W+E)
        self.m.menu = Menu(self.m,tearoff=0)
        self.m["menu"]=self.m.menu
        for q in self.questions:
            self.m.menu.add_command(label=str(q),command=CMD(self.change_q,q))
   
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

        self.geg = Text(self,width=60,height=10,wrap=WORD,bg="#eeeeff",relief=SUNKEN)
        self.geg.grid(row=4,column=0,columnspan=3)

        self.req = Text(self,width=60,height=10,wrap=WORD,bg="#eeffee",relief=SUNKEN)
        self.req.grid(row=4,column=3,columnspan=3)

        self.bem = Text(self,width=60,height=10,wrap=WORD,bg="#ffffff",relief=SUNKEN)
        self.bem.grid(row=7,column=0,columnspan=3)

        self.f = Frame(self)
        self.f.grid(row=6,column=3,rowspan=2,columnspan=3)

        ld = Label(self.f,text="Punkte:")
        ld.grid(row=0,column=0,sticky=W)
        self.points=StringVar()
        self.points.set("hallo")
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

        self.refresh()


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

        self.cq['text'] = "  Aktuelle Frage: "+str(q)

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
        self.geg.insert(END,curr['Given'])
        self.req.delete("1.0",END)
        self.req.insert(END,curr['Correct'])
        self.geg["state"]=DISABLED
        self.req["state"]=DISABLED

        self.bem.delete("1.0",END)
        self.bem.insert(END,curr['Remarks'])

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
        rem = self.bem.get("1.0",END)[:-1]
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

        outf = open(settings['Filename_Out'],'w')
        outf.write(str(self.resp))
        outf.close()

        if refresher:
            self.refresh()

        return accept
        

global settings
settings = {}
settings['Curr_R']=0
settings['Curr_Q']=0
settings['Insecurities']=[]
settings['Filename']="results.json"
settings['Filename_Out']="results_corr.json"

## If the outfile already exists (because it was closed), open the outfile, not the original.
try:
    inf = open(settings['Filename_Out'],'r',encoding='utf-8',errors='ignore')
except:
    inf = open(settings['Filename'],'r',encoding='utf-8',errors='ignore')
json = inf.readline()
inf.close()


j = eval(json)
if type(j)==dict:
    print("Daten von "+str(len(j.keys()))+" Studierenden erfolgreich geladen.")
    #print(j[list(j.keys())[0]])
    root = Tk()
    korr = Anzeige(root,j)
    root.mainloop()
