import os
import subprocess

from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

BASE_DIR = os.getcwd()

app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = 'uploads'

def doc2pdf_linux(doc):
    """
    convert a doc/docx document to pdf format (linux only, requires libreoffice)
    :param doc: path to document
    """
    cmd = 'libreoffice --convert-to pdf'.split() + [doc]
    cmd += ['--outdir', 'output']
    print(cmd)
    p = subprocess.Popen(cmd, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
    p.wait(timeout=100)
    stdout, stderr = p.communicate()
    print(stderr)
    if stderr:
        raise subprocess.SubprocessError(stderr)


@app.route('/', methods=['GET'])
def home_post():
    os.system('rm -r uploads')
    try:
        os.mkdir('uploads')
        os.mkdir('output')
    except Exception as e:
        print(e)
    return render_template('index.html')


@app.route('/', methods=['POST'])
def home():
    for file in request.files.getlist('files'):
        filename = file.filename
        file.save(
            os.path.join(
                app.config['UPLOAD_FOLDER'],
                secure_filename(filename)
            )
        )
    os.chdir('uploads')
    for file in os.listdir():
        if not file.endswith('.docx'):
            continue
        doc2pdf_linux(file)
    os.system('zip output.zip output/*')
    os.chdir(BASE_DIR)
    return render_template('download.html')


@app.route('/download', methods=['GET'])
def download():
    return send_from_directory(app.config["UPLOAD_FOLDER"], 'output.zip')
