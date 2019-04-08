import os
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from helper import data_process
#test github

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

    if startdate == '':
        return "No Date selected."

    if file and allowed_file(file.filename) and startdate:
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], 'upload.xls'))
        data = data_process.process(os.path.join(app.config['UPLOAD_FOLDER'] + '\\' + 'upload.xls'), startdate)
        return render_template('home.html', data=data)
    return "Test"


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/download_excel', methods=["POST"])
def download_excel():
    excel = data_process.download_excelfile()
    return send_file('./download_folder/result.xlsx', as_attachment=True)

@app.errorhandler(404)
def error_404(error):
    return render_template('404.html'),404
