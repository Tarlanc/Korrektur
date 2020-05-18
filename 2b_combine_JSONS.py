

## The first filename will be considered the master.
## If some other file disagrees with the master,
## an insecurity will be marked.
## Likewise, the second file is more master than the third, and so on.

fnames = ["results_1.json",
          "results_2.json",
          "results_3.json",
          "results_4.json",
          "results_5.json"]

outf = 'results_corr.json'
conf = 'conflicts.txt'

conflicts = []
results = {}
for f in fnames:
    inf = open(f,'r',encoding='utf-8',errors='ignore')
    l = inf.readline()
    inf.close()
    counter = [0,0,0] ## New scores, replacements, conflicts
    try:
        r = eval(l)
    except:
        r = {}
        print("WARNING: File '"+f+"' is not a valid file")

    for p in r.keys():
        try:
            a = results[p]
        except:
            results[p]={}
        for q in r[p].keys():
            try:
                a = results[p][q]
            except:
                results[p][q] = {}

            if not 'State' in results[p][q].keys():
                results[p][q]['State']=-1

            ## Overwrite if new file holds information about an unknown question/respondent
            if not 'Points' in results[p][q].keys(): 
                for qp in r[p][q].keys(): results[p][q][qp] = r[p][q][qp]
                counter[0]+=1
            ## abort if there are no new points:
            elif type(r[p][q]['Points'])==str:
                pass          
            ## Overwrite if there are no points yet
            elif type(results[p][q]['Points'])==str:
                for qp in r[p][q].keys(): results[p][q][qp] = r[p][q][qp]
                counter[0]+=1
            ## Ignore new info if points match.
            elif results[p][q]['Points']==r[p][q]['Points']:
                if 'State' in r[p][q]:
                    if r[p][q]['State'] > results[p][q]['State']:
                        for qp in r[p][q].keys(): results[p][q][qp] = r[p][q][qp]
                        counter[1]+=1
            ## Handle conflict if there is one
            else:
                counter[2]+=1
                conflicts.append((p,q,f,results[p][q]['Points'],r[p][q]['Points']))
                results[p][q]['State']=1 ## Set conflict for this question
                results[p][q]['Remarks']+=" conflicting remarks: "+r[p][q]['Remarks']
    print("Successfully added "+str(counter[0])+" new marks and "+str(counter[1])+" new remarks with "+str(counter[2])+" conflicts to results from file "+str(f))


print("\nFinished with "+str(len(conflicts))+" conflicts")
outf = open(outf,'w')
outf.write(str(results))
outf.close()

outf = open(conf,'w')
for c in conflicts:
    outf.write(str(c)+'\n')
outf.close()
