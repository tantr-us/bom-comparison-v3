from openpyxl import load_workbook
from flask import Flask, render_template, url_for, redirect, request
from werkzeug.utils import secure_filename
import os, secrets


app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)

# secret key
app.config['SECRET_KEY'] = secrets.token_hex(32)
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')


@app.route('/test', methods=['GET', 'POST'])
def test():
    if request.method == 'GET':
        return render_template('read_template.html')



def get_column_index(template_path):
    pass



if __name__ == '__main__':
    app.run(debug=True)
