# =========================================================================
# == FINAL UPDATED PYTHON FILE (Sends Filename to API - Error 4150 Fix) ==
# =========================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import pdfplumber
from io import BytesIO
import subprocess

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)


# --- 1. PDF to Word Tool (FIXED ERROR 4150) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    API_SECRET = os.getenv('API_SECRET')
    if not API_SECRET:
        return jsonify({"error": "API Secret Key server par set nahi ki gayi hai."}), 500
    
    API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={API_SECRET}'
    
    if 'file' not in request.files:
        return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    try:
        # --- FIX: File ka naam, data aur type, teeno cheezein saath bhejein ---
        # Yeh line humne update ki hai
        files_to_send = {'file': (file.filename, file.read(), 'application/pdf')}
        
        response = requests.post(API_URL, files=files_to_send)
        response.raise_for_status()
        
        converted_file_data = BytesIO(response.content)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except requests.exceptions.HTTPError as e:
        return jsonify({"error": f"API se error aaya: {e.response.text}"}), 500
    except Exception as e:
        return jsonify({"error": f"Koi masla ho gaya: {str(e)}"}), 500


# --- 2. PDF to Excel Tool ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']

    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        all_dataframes = []
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table:
                        df = pd.DataFrame(table)
                        all_dataframes.append(df)
        if not all_dataframes:
            os.remove(input_path)
            return jsonify({"error": "Is PDF mein koi table nahi mila."}), 400
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        output_buffer = BytesIO()
        combined_df.to_excel(output_buffer, sheet_name='Extracted_Data', index=False, header=False)
        output_buffer.seek(0)
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        os.remove(input_path)
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, 
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        if os.path.exists(input_path):
            os.remove(input_path)
        return jsonify({"error": f"Excel banate waqt masla hua: {str(e)}"}), 500


# --- 3. Word to PDF Tool ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']

    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Invalid file type, please upload a .doc or .docx file."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path):
            os.remove(input_path)
            return jsonify({"error": "Conversion failed, output file not found."}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path)
        os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path):
            os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

# --- 4. Excel to PDF Tool ---
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']

    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Invalid file type, please upload an .xls or .xlsx file."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path):
            os.remove(input_path)
            return jsonify({"error": "Conversion failed, output file not found."}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path)
        os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path):
            os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

# --- Server Chalane ke liye Code ---
if __name__ == '__main__':
    app.run(debug=True, port=5000)