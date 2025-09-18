# =========================================================================================
# == FINAL PROFESSIONAL VERSION V11.0 - Stable and Reliable Logic for All Tools ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (Using ConvertAPI with SECRET) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    CONVERTAPI_SECRET = os.getenv('CONVERTAPI_SECRET')
    if not CONVERTAPI_SECRET:
        return jsonify({"error": "ConvertAPI Secret is not set on the server. Please check the environment variables."}), 500

    API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={CONVERTAPI_SECRET}'
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        files_to_send = {'file': (file.filename, file.read(), 'application/pdf')}
        response = requests.post(API_URL, files=files_to_send)
        response.raise_for_status()
        converted_file_data = BytesIO(response.content)
        if converted_file_data.getbuffer().nbytes < 100:
             raise Exception("Converted file is empty. Check your API plan or file.")
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except requests.exceptions.HTTPError as e:
        return jsonify({"error": f"API Error: Please check your API Secret. Details: {e.response.text}"}), 500
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (SAFE & RELIABLE LOGIC) ---
# Yeh har line ko ek cell mein rakhega
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_text = ""
        for page in doc:
            all_text += page.get_text()

        doc.close()

        if not all_text.strip():
            return jsonify({"error": "No text could be extracted from this PDF."}), 400
        
        # Har line ko ek list item banayein
        lines = all_text.strip().split('\n')
        df = pd.DataFrame(lines, columns=["Data"])

        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Extracted_Data', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Excel file: {str(e)}"}), 500


# --- Other Tools (No changes needed) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main(): return internal_word_to_pdf()

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main(): return internal_excel_to_pdf()

# ... in dono functions ko yahan paste karna hai (unmein koi change nahi) ...
def internal_word_to_pdf():
    # Pehle wala code
    pass
def internal_excel_to_pdf():
    # Pehle wala code
    pass
    
if __name__ == '__main__':
    app.run(debug=True, port=5000)