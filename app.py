# =====================================================================================
# == FINAL PROFESSIONAL VERSION V8.0 - Word on ConvertAPI TOKEN | Pro Excel Layout ==
# =====================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess
import statistics

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (FINAL VERSION - Using Production Token) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # Render ke Environment se ab hum 'Token' uthayenge
    CONVERTAPI_TOKEN = os.getenv('CONVERTAPI_TOKEN')
    if not CONVERTAPI_TOKEN:
        return jsonify({"error": "ConvertAPI TOKEN is not set on the server. Please check the environment variables."}), 500

    # API URL se Secret hata diya gaya hai
    API_URL = 'https://v2.convertapi.com/convert/pdf/to/docx'

    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400

    try:
        # NAYI CHEEZ: Ab hum 'Token' ko 'Header' mein bhejenge
        headers = {'Authorization': f'Bearer {CONVERTAPI_TOKEN}'}
        
        files_to_send = {'file': (file.filename, file.read(), 'application/pdf')}
        
        # Request ke saath 'headers' bhi bhej rahe hain
        response = requests.post(API_URL, files=files_to_send, headers=headers)
        response.raise_for_status()
        
        converted_file_data = BytesIO(response.content)
        if converted_file_data.getbuffer().nbytes < 100:
             raise Exception("The converted file is empty, which may indicate an API issue.")

        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except requests.exceptions.HTTPError as e:
        return jsonify({"error": f"API Error. Check your API Token and plan limits. Details: {e.response.text}"}), 500
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (FINAL PROFESSIONAL LAYOUT LOGIC) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    #... (Yeh code pehle jaisa hi hai, ismein koi badlaav nahi) ...
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_pages_data = []
        for page in doc:
            words = page.get_text("words")
            if not words: continue
            lines = {}
            for w in words:
                y0 = round(w[1])
                line_key = min(lines.keys(), key=lambda y: abs(y-y0), default=None)
                if line_key is not None and abs(line_key - y0) < 5: lines[line_key].append(w)
                else: lines[y0] = [w]
            
            x_coords = sorted(list(set([round(w[0]) for w in words])))
            if not x_coords: continue
            clusters = [[x_coords[0]]]
            for x in x_coords[1:]:
                if x - clusters[-1][-1] < 10: clusters[-1].append(x)
                else: clusters.append([x])
            column_positions = sorted([statistics.mean(c) for c in clusters])
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0])
                row = [""] * len(column_positions)
                for word in line_words:
                    x0 = round(word[0])
                    col_index = min(range(len(column_positions)), key=lambda i: abs(column_positions[i] - x0))
                    if row[col_index] == "": row[col_index] = word[4]
                    else: row[col_index] += " " + word[4]
                all_pages_data.append(row)

        doc.close()
        if not all_pages_data: return jsonify({"error": "No text data could be extracted from this PDF."}), 400
        
        df = pd.DataFrame(all_pages_data)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Data_In_One_Sheet', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Excel file: {str(e)}"}), 500

# --- Baaki ke tools (No changes) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main(): return internal_word_to_pdf()
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main(): return internal_excel_to_pdf()
def internal_word_to_pdf():
    # ... code
    pass
def internal_excel_to_pdf():
    # ... code
    pass
if __name__ == '__main__':
    app.run(debug=True, port=5000)