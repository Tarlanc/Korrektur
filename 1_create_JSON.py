
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
                'Type':q['Type']}

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
                'Type':q['Type']}

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
                'Type':q['Type']}

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
                'Type':q['Type']}

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
                'Type':q['Type']}           
    return mark


################## Anpassen für unterschiedliche Tests

kc = 'Laufnummer' ## Key column. Might be revised for tests.
fname = "Eingabe.xls"
outfname = "results.json"

##################


inf = open(fname,'r',
           encoding="utf-8",errors='ignore')
ls = inf.readlines()
inf.close()

#print(ls[205:209])
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
            for a in ans['Ans'].keys():
                qkeys[a]=q[2]
            

print('\nStruktur der Fragen:\n------------\n\n'+baum_schreiben(questions))
#print(baum_schreiben(qkeys))

## Now we know all the questions and the correct answers
## And a key dictionary to find the correct question

## Now for the respondent dictionary

rd = {}
for i in range(len(resp[kc])):
    r = resp[kc][i]
    rd[r]={}
    for q in questions.keys():
        response = {}
        for a in questions[q]['Ans'].keys():
            response[a] = resp[a][i]
        rd[r][q] = mark_question(questions[q],response)

#print(list(rd.keys()))

o = open(outfname,'w')
o.write(str(rd))
o.close()

input("Die automatischen Korrekturen sind erledigt und liegen in "+outfname+".\nEingabe drücken, um dieses Fenster zu schliessen...")

            
        



