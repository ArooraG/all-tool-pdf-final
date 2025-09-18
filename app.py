# =================================================================
# == YEH AAPKI MUKAMMAL AUR UPDATED PYTHON BACKEND FILE HAI ==
# =================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import pdfplumber
from io import BytesIO
import subprocess

app = Flask(__name__)
CORS(app)  # Isse aapka frontend backend se connect ho payega

# File upload karne ke liye ek folder banayega
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Yeh ek helper function hai file ka path banane ke liye
def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)


# --- 1. PDF to Word Tool (FIXED ERROR 401) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # --- NAYA AUR BEHTAR CODE ---
    # Yeh API Secret Key ab Render ke "Environment" se khud-ba-khud uthayega
    # Aapko yeh key Render ke dashboard mein save karni hogi
    API_SECRET = os.getenv('API_SECRET')
    if not API_SECRET:
        # Agar key set nahi hogi to yeh error aayega
        return jsonify({"error": "API Secret Key server par set nahi ki gayi hai."}), 500

    API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={API_SECRET}'

    if 'file' not in request.files:
        return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    try:
        # File ko seedha ConvertAPI ko bhej rahe hain
        response = requests.post(API_URL, files={'file': file.read()})
        response.raise_for_status()  # Agar API se error aaye to yahin ruk jayega
        
        converted_file_data = BytesIO(response.content)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except requests.exceptions.HTTPError as e:
        # Agar API key ghalat ho ya koi aur masla ho to API wala error dikhayega
        return jsonify({"error": f"API se error aaya: {e.response.text}"}), 500
    except Exception as e:
        return jsonify({"error": f"Koi masla ho gaya: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (UPDATED: SAARI TABLES EK SHEET MEIN) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    # --- NAYA AUR BEHTAR LOGIC ---
    if 'file' not in request.files:
        return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        
        all_dataframes = []
        with pdfplumber.open(input_path) as pdf:
            # Har page se tables nikaal kar ek list mein daal rahe hain
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table: # Agar table khali nahi hai
                        df = pd.DataFrame(table)
                        all_dataframes.append(df)

        if not all_dataframes:
            os.remove(input_path)
            return jsonify({"error": "Is PDF mein koi table nahi mila."}), 400

        # Sab tables ko jod kar ek bada table bana rahe hain
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        
        output_buffer = BytesIO()
        # Is bade table ko ek he sheet par save kar rahe hain
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
# Note: Iske liye Render par LibreOffice ka hona zaroori hai (render.yaml file ke zariye)
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.endswith(('.doc', '.docx')):
        return jsonify({"error": "Invalid file type, please upload a .doc or .docx file."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        
        if not os.path.exists(pdf_path):
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
# Note: Iske liye bhi Render par LibreOffice ka hona zaroori hai
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Invalid file type, please upload an .xls or .xlsx file."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        
        if not os.path.exists(pdf_path):
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
    # Yeh hissa sirf local testing ke liye hai. Render Gunicorn istemal karega.
    app.run(debug=True, port=5000)