import os
import json
from openpyxl import Workbook, load_workbook
import datetime
from dataclasses import dataclass
from flask import Flask, request, render_template, url_for, redirect, jsonify
from classes import Product

app = Flask(__name__)
ALLOWED_EXTENSIONS = set(['xlsx'])
REMOVE_TABS = ['Cover Sheet', 'Index', 'General Info', 'DataSet', 'Changes']
categories = []
products = []
workbook = Workbook()


def get_categories():
    categories = workbook.sheetnames
    return categories


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part.')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            workbook = load_workbook(filename=file.filename, read_only=True)
            sheet = workbook.active
            categories = workbook.sheetnames
            print(categories)
            return render_template('fileform.html', categories = categories)
            return redirect('/')
        else:
            flash('Allowed file type is xlsx')
            return redirect(request.url)
    return render_template('upload.html')


@app.route("/main")
def mainPage():
    return render_template('index.html')

@app.route("/uploadfile", methods=['GET', 'POST'])
def uploadfile():
    if request.method == 'POST':
        filename = request.args.get('file')
        category = request.args.get('category')
        if category:
            workbook = load_workbook(filename, read_only=True)
            sheet = workbook.get_sheet_by_name('category')
            print(category)
        if filename:
            workbook = load_workbook(filename=request.args.get('file'), read_only=True)
            products = workbook.sheetnames
            for sheet in products:
                if sheet in REMOVE_TABS:
                    pass
                else:
                    categories.append(sheet)
            return render_template('index.html', categories = categories)
    return render_template('index.html')
#    '''
#    <!doctype html>
#    <title>Upload an excel file</title>
#    <h1>Excel file upload (csv, tsv, csvz, tsvz only)</h1>
#    <form action="" method=post enctype=multipart/form-data>
#    <p><input type=file name=file><input type=submit value='Load Data'>
#    </form>
#    '''

@app.route("/handleFileUpload", methods=['POST'])
def handleFileUpload():
    if 'input_file' in request.files:
        input_file = request.files['input_file']
        if input_file.filename != '':
            input_file.save(os.path.join('/tmp', input_file.filename))
        workbook = load_workbook(filename=input_file.filename, read_only=True)
        sheet = workbook.active
        categories = workbook.sheetnames
        print(categories)
        return render_template('fileform.html', categories)
    return redirect(url_for('fileFrontPage'))


if __name__ == '__main__':
    app.run()
