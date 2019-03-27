import os
from flask import Flask, render_template, request, redirect, url_for, Response
from werkzeug.utils import secure_filename
from helper import checking
import time


UPLOAD_FOLDER = os.path.abspath(os.path.dirname(__file__)) + '\\upload_folder'
ALLOWED_EXTENSIONS = set(['xls', 'xlsx', 'xlsm'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/', methods=["GET", "POST"])
def index():
    return render_template('home.html')


@app.route('/upload', methods=["POST"])
def upload():
    if 'file' not in request.files:
        return "No File Uploaded1"

    file = request.files['file']

    if file.filename == '':
        return "No File Uploaded2"

    startdate = request.form['StartDate']

    if startdate =='':
        return "No Date selected."

    if file and allowed_file(file.filename) and startdate:
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        data = checking.Checking(os.path.join(app.config['UPLOAD_FOLDER'] + '\\' + filename), startdate)
        return render_template('home.html', data=data)
    return "Test"


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/download_excel', methods=["POST"])
def download_excel():
    timestr = time.strftime("%Y%m%d-%H%M%S")
    filename1 = 'result %s.xlsx' % (timestr)
    excel = checking.download_excel()

    return Response(excel, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-disposition":"attachment; filename =" + filename1})

@app.errorhandler(404)
def error_404(error):
    return render_template('404.html'),404






#if __name__ == ("__main__"):
#    app.run(debug = True)
