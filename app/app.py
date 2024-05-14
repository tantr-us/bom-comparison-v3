from datetime import datetime
from sqlite3 import Error

from flask import Flask, render_template, flash, request, redirect, send_file, url_for, jsonify
from werkzeug.utils import secure_filename

from core import validate_bc_template, run_bc
# from core_raw import validate_bc_template_raw, run_bc_raw
from modules.errors import MissingRequiredWorksheetError, RefDesError
from modules.utils import get_header_index_from_xl

import os
import secrets
import sqlite3
import traceback

app = Flask(__name__)

# Prevent jsonify from sorting the header indies
app.config['JSON_SORT_KEYS'] = False # for flask version < 2.3
app.json.sort_keys = False

# secret key
app.config['SECRET_KEY'] = secrets.token_hex(32)


BASE_DIR = os.path.dirname(__file__)


# Config upload folder
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
REPORT_DIR = os.path.join(BASE_DIR, 'bc_reports')


# BOM DIFF TEMPLATE
# BOM_DIFF_TEMPLATE = '# BOMDIFF_TEMPLATE (BLANK).xlsx'
BOM_DIFF_TEMPLATE = '# BOMDIFF_TEMPLATE (BLANK).xlsm'


# DB
DATA_DB = 'data.sqlite3'
MAPPING_DB = 'mapping.sqlite3'

### METHOD 1 - Using Excel template with VBA to generate comparable BOM ###
@app.route('/', methods=['GET', 'POST'])
def bom_diff():
    mapping_list = None
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        
        if file.filename == '':
            flash('No file selected.')
            return redirect(request.url)
        elif file and file.filename.lower().endswith('.xlsm'):
            filename = secure_filename(file.filename)
            bc_template = os.path.join(app.config['UPLOAD_FOLDER'], f'{datetime.now().strftime("%Y-%m-%d_%H%M%S")}_{filename}')
            file.save(bc_template)
            
            # Validate bc template
            try:
                # get mapping from database
                mapping_id = request.form['customer-mapping']
                mapping_id = int(mapping_id)
                mapping_setting = get_mapping_setting(mapping_id)
            
                # get compare fields:
                compare_description = request.form['checkbox-description']
                compare_uom = request.form['checkbox-uom']
                compare_qty = request.form['checkbox-quantity']
                compare_rev = request.form['checkbox-revision']
                compare_ref_des = request.form['checkbox-refdes']
                compare_mfr = request.form['checkbox-mfr']
                
                compare_fields_setting = {
                    'desc': int(compare_description),
                    'uom': int(compare_uom),
                    'qty': int(compare_qty),
                    'rev': int(compare_rev),
                    'refdes': int(compare_ref_des),
                    'mfr': int(compare_mfr),
                }
            
                validate_bc_template(bc_template)
                bc_report = run_bc(bc_template, mapping_setting, compare_fields_setting)
                flash('Done! The report is ready for download.')

                customer_name = request.form['customer-name']
                if customer_name == '':
                    customer_name = 'no name'

                memo = request.form['memo']
                if memo == '':
                    memo = 'no memo'
                current_time = datetime.now()
                save_to_db(customer_name, memo, bc_report, current_time)
                return redirect(url_for('bom_diff', compare_status='success'))
            except MissingRequiredWorksheetError as err:
                os.remove(bc_template)
                flash(str(err))
                return redirect(url_for('bom_diff', compare_status='error'))
            except TypeError as err:
                traceback.print_exc()
                flash(str(err))
                return redirect(url_for('bom_diff', compare_status='error'))
            except ValueError as err:
                traceback.print_exc()
                flash(str(err))
                return redirect(url_for('bom_diff', compare_status='error'))
            except RefDesError as err:
                traceback.print_exc()
                flash(str(err))
                return redirect(url_for('bom_diff', compare_status='error'))
        else:
            flash('Only accept .xlsm file.')
            return redirect(request.url)
    elif request.method == 'GET':
        mapping_list = get_mapping_list()

    return render_template('compare.html', page='compare', mapping_list=mapping_list)
## End METHOD 1 ###


@app.route('/view_downloads')
def view_downloads():
    download_data_list = None
    try:
        connection = sqlite3.connect(DATA_DB)
        cursor = connection.cursor()
        query = "SELECT * FROM downloads ORDER BY reported_date DESC"
        cursor.execute(query)
        download_data_list = cursor.fetchall()
        cursor.close()
        connection.close()
    except Error as e:
        print(e)
    return render_template('download.html', page='download', download_data_list=download_data_list)


@app.route('/download_report/<filename>')
def download_report(filename):
    file_path = os.path.join(REPORT_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash('Report not found.')
        return redirect(url_for('view_downloads'))


@app.route('/delete_report/<filename>')
def delete_report(filename):
    file_path = os.path.join(REPORT_DIR, filename)
    # print(f"Report: {file_path}")
    if os.path.exists(file_path):
        os.remove(file_path)
    # Delete record from database
    try:
        connection = sqlite3.connect(DATA_DB)
        cursor = connection.cursor()
        query = "DELETE FROM downloads WHERE filename=?"
        cursor.execute(query, (filename,))
        connection.commit()
        cursor.close()
        connection.close()
    except Error as e:
        print(e)
    return redirect(url_for('view_downloads'))


@app.route('/get_template')
def get_template():
    template_filename = BOM_DIFF_TEMPLATE
    template_path = os.path.join(os.path.join(os.getcwd(), 'bc_templates'), template_filename)
    return send_file(template_path, as_attachment=True)


@app.route('/customer_mapping')
def customer_mapping():
    mapping_list = None
    try:
        connection = sqlite3.connect(MAPPING_DB)
        cursor = connection.cursor()
        query = "SELECT * FROM mapping ORDER BY mapping_name"
        cursor.execute(query)
        mapping_list = cursor.fetchall()
        cursor.close()
        connection.close()
    except Error as e:
        print(e)
    return render_template('mapping.html', page='mapping', mapping_list=mapping_list)


@app.route('/add_mapping', methods=['GET', 'POST'])
def add_mapping():
    if request.method == 'POST':
        mapping_name = request.form['mapping-name'].upper()
        make_prefix = request.form['make-prefix'].upper()
        buy_prefix = request.form['buy-prefix'].upper()
        consigned_suffix = request.form['consigned-suffix'].upper()
        cus_docs_prefix = request.form['cus-doc-prefix'].upper()
        rev_delimiter = request.form['rev-delimiter']
        special_char_delimiter = None if request.form['spec-char-delimiter'] == '' else request.form['spec-char-delimiter']
        sample_customer_number = None if request.form['sample-customer-number'] == '' else request.form['sample-customer-number']
        save_mapping_to_db(mapping_name, make_prefix, buy_prefix, consigned_suffix, cus_docs_prefix, rev_delimiter, special_char_delimiter, sample_customer_number)
        return redirect(url_for('customer_mapping', page='mapping'))
    return render_template('add_mapping.html', page='mapping')


@app.route('/modify_mapping/<id>', methods=['GET', 'POST'])
def modify_mapping(id):
    mapping = None
    if request.method == 'GET':
        mapping = list(get_mapping_from_db(id))
        if mapping[7] is None:
            mapping[7] = ''
        if mapping[8] is None:
            mapping[8] = ''
    elif request.method == 'POST':
        mapping_name = request.form['mapping-name'].upper()
        make_prefix = request.form['make-prefix'].upper()
        buy_prefix = request.form['buy-prefix'].upper()
        consigned_suffix = request.form['consigned-suffix'].upper()
        cus_docs_prefix = request.form['cus-doc-prefix'].upper()
        rev_delimiter = request.form['rev-delimiter']
        special_char_delimiter = None if request.form['spec-char-delimiter'] == '' else request.form['spec-char-delimiter']
        sample_customer_number = None if request.form['sample-customer-number'] == '' else request.form['sample-customer-number']
        update_mapping(id, mapping_name, make_prefix, buy_prefix, consigned_suffix, cus_docs_prefix, rev_delimiter, special_char_delimiter, sample_customer_number)
        return redirect(url_for('customer_mapping', page='mapping'))
    
    return render_template('modify_mapping.html', page='mapping', mapping=mapping)
    

@app.route('/delete_mapping/<id>')
def delete_mapping(id):
    try:
        connection = sqlite3.connect('mapping.sqlite3')
        cur = connection.cursor()
        sql = "delete from mapping where id=?"
        cur.execute(sql, (id,))
        connection.commit()
        cur.close()
        connection.close()
        return redirect(url_for('customer_mapping'))
    except Error as e:
        print(e)


def save_to_db(customer_name, memo, report_name, current_time):
    try:
        connection = sqlite3.connect(DATA_DB)
        cursor = connection.cursor()
        query = "INSERT INTO downloads(customer, name, filename, reported_date) VALUES (?,?,?,?);"
        cursor.execute(query, (customer_name, memo, report_name, current_time))
        connection.commit()
        cursor.close()
        connection.close()
    except Error as e:
        print(e)


def save_mapping_to_db(mapping_name, make_prefix, buy_prefix, consigned_suffix, cus_docs_prefix, rev_delimiter, special_char_delimiter, sample_customer_number):
    try:
        connection = sqlite3.connect(MAPPING_DB)
        cursor = connection.cursor()
        query = "INSERT INTO mapping(mapping_name, make_prefix, buy_prefix, consigned_suffix, cus_document, rev_delimiter, special_delimiter, sample_number) VALUES (?,?,?,?,?,?,?,?);"
        cursor.execute(query, (mapping_name, make_prefix, buy_prefix, consigned_suffix, cus_docs_prefix, rev_delimiter, special_char_delimiter, sample_customer_number))
        connection.commit()
        cursor.close()
        connection.close()
    except Error as e:
        print(e)


def update_mapping(id, mapping_name, make_prefix, buy_prefix, consigned_suffix, cus_docs_prefix, rev_delimiter, special_char_delimiter, sample_customer_number):
    try:
        connection = sqlite3.connect(MAPPING_DB)
        cursor = connection.cursor()
        query = "UPDATE mapping SET mapping_name=?, make_prefix=?, buy_prefix=?, consigned_suffix=?, cus_document=?, rev_delimiter=?, special_delimiter=?, sample_number=? WHERE id=?"
        cursor.execute(query, (mapping_name, make_prefix, buy_prefix, consigned_suffix, cus_docs_prefix, rev_delimiter, special_char_delimiter, sample_customer_number, id))
        connection.commit()
        cursor.close()
        connection.close()
    except Error as e:
        print(e)


def get_mapping_from_db(id): 
    try:
        connection = sqlite3.connect('mapping.sqlite3')
        cur = connection.cursor()
        sql = "select * from mapping where id=?"
        cur.execute(sql, (id,))
        mapping = cur.fetchall()[0]
        cur.close()
        connection.close()
        return mapping
    except Error as err:
        print(err)


def get_mapping_setting(id):
    # get mapping from database   
    try:
        connection = sqlite3.connect('mapping.sqlite3')
        cur = connection.cursor()
        sql = "select * from mapping where id=?"
        cur.execute(sql, (id,))
        mapping = cur.fetchall()[0]
        cur.close()
        connection.close()
        return {
            'make': mapping[2],
            'buy': mapping[3],
            'consigned suffix': mapping[4],
            'customer docs': mapping[5],
            'rev delimiter': mapping[6],
            'special delimiter': mapping[7],
            'sample customer number': mapping[8],
        }
    except Error as err:
        print(err)


def get_mapping_list():
    try:
        connection = sqlite3.connect(MAPPING_DB)
        cursor = connection.cursor()
        query = "SELECT * FROM mapping ORDER BY mapping_name"
        cursor.execute(query)
        mapping_list = cursor.fetchall()
        cursor.close()
        connection.close()
        return mapping_list
    except Error as e:
        print(e)

if __name__ == '__main__':
    app.run(debug=True)
