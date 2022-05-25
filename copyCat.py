from lib import *
import docx

def getText(filename):
    alltext=[]
    data=""
    doc=docx.Document(filename)
    for docpar in doc.paragraphs:
        alltext.append(docpar.text)
    data=''.join(alltext)
    return(data)
def plgcheck(f1,f2):
    srcstring=getText(f1)
    ansstring=getText(f2)
    sentsrc=srcstring.split('.')
    sentans=ansstring.split('.')
    matchlist=[]
    per=0
    for s in sentsrc:
        for a in sentans:
            if(s==a):
                matchlist.append(s)
                per=per+1
    per=(per/len(sentans))*100
    return(int(per),matchlist)