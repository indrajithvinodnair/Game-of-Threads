from lib import *
from plagarismCheck import *
app = Flask(__name__)
import docx
from requirements import *
from copyCat import *


#app.wsgi_app = StreamConsumingMiddleware(app.wsgi_app)

app.config["UPLOAD_FOLDER"] = "static" #folder to upload

@app.route("/",methods=['GET', 'POST'])
def index():
    #return render_template("welcome.html")
    return render_template("index.html")

@app.route("/login.html",methods=['GET', 'POST'])
def login():
    return render_template("login.html")

@app.route("/dashboard.html",methods=['GET', 'POST'])
def dashboard():
    return render_template("dashboard.html")

@app.route("/viewExcel.html",methods=['GET', 'POST'])
def viewExcel():
    return render_template("viewExcel.html")

@app.route("/qbresult.html",methods=['GET', 'POST'])
def qbresult():
    return render_template("qbresult.html")


_FILE = ""
@app.route('/answerAnalyzer.html',methods=['GET', 'POST']) # redirecting url : answerAnalyzer.html
def upload_excel():
    global _FILE
    if request.method == "POST":
        upload_answerkey = request.files['answer-key']
        upload_answer_std = request.files['answer-std']
        if upload_answerkey.filename != '' and upload_answer_std.filename != '':
            filepath = os.path.join(app.config["UPLOAD_FOLDER"] ,upload_answerkey.filename)
            filepath2 = os.path.join(app.config["UPLOAD_FOLDER"] ,upload_answer_std.filename)
            upload_answerkey.save(filepath)
            upload_answer_std.save(filepath2)
            _FILE = filepath2
            A_obj = an.KeyWord(filepath2,filepath)
            data = pd.read_excel(filepath2)
            return render_template("viewExcel.html", data=data.to_html(index=False))
    return render_template("uploadExcel.html")

@app.route('/viewExcel')
def download_excel_file():
    p = _FILE
    return send_file(p,as_attachment=True)

_FILE2 = ""
@app.route('/questionGenerator.html',methods=['GET', 'POST']) # redirecting url : answerAnalyzer.html
def upload_qb():
    qb.deleteStaticFiles()
    if request.method == 'POST':
        upload_questionbank = request.files['qb_file']
        if upload_questionbank.filename != '':
            filepath = os.path.join(app.config["UPLOAD_FOLDER"],upload_questionbank.filename)
            upload_questionbank.save(filepath)
            fpath = filepath.split("'\'")
            qb.acceptPath(fpath[0])
            return render_template("qbresult.html")
    return render_template("qbgen.html")



@app.route('/qbresult')
def download_qb_file():
    p = "E:\Ravana\workstation\general\Coronis\GameOFThreads\static\demo.docx"
    return send_file(p,as_attachment=True)

        
@app.route('/plagarismCheck.html',methods=['GET', 'POST'])
def plagarismCheck():
    return render_template("plagarismCheck.html")

@app.route('/copyCatChecker.html',methods=['GET', 'POST'])
def copyCatChecker():
    return render_template("copyCatChecker.html")

@app.route('/plagarism',methods=['GET', 'POST'])

def checkplgweb():
    if(request.method=='POST'):
        srcurl=request.form['url1']
        ansfile=request.files['src1']
        alltext=[]
        data=""
        doc=docx.Document(ansfile)
        for docpar in doc.paragraphs:
            alltext.append(docpar.text)
        data=''.join(alltext)
        copied=checkpg(data,srcurl)
        return(render_template('plagarismResult.html',c=copied))

@app.route('/chkplgsrc',methods=['GET','POST'])
def updateres():
        if(request.method=='POST'):
            srcfile=request.files['src']
            ansfile=request.files['ans']
            per,matchlist=plgcheck(srcfile,ansfile)
            return(render_template('copycatResult.html',per=per,m=matchlist))
    
if __name__ == '__main__':
    app.run(debug=True)