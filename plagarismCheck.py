from lib import *
import requests

def checkpg(q,url):  
    copied=[]
    request_result=requests.get( url )
    soup = bs4.BeautifulSoup(request_result.text, "html.parser")
    links=soup.select('a')
    hrefs=[]
    for l in links:
      x=l.get('href')
      hrefs.append(x)
    for l in hrefs[1:]:
      request_result=requests.get(str(l))
      soup = bs4.BeautifulSoup(request_result.text, "html.parser")
      srcstring=soup.get_text()
      if(q in srcstring):
        copied.append(q)
        break;
      else:
        pass
    return(copied)