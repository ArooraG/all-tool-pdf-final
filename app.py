# =========================================================================================
# == FINAL PROFESSIONAL VERSION V7.0 - Word on ConvertAPI | Professional Excel Layout ==
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

# --- 1. PDF to Word Tool (FINAL VERSION - Back to your choice: ConvertAPI) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # Render ke Environment se Secret Key uthayega
    CONVERTAPI_SECRET = os.getenv('CONVERTAPI_SECRET')
    if not CONVERTAPI_SECRET:
        return jsonify({"error": "ConvertAPI Secret is not set on the server."}), 500

    API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={CONVERTAPI_SECRET}'

    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400

    try:
        # File ka naam, data aur type, teeno cheezein saath bhejna zaroori hai
        files_to_send = {'file': (file.filename, file.read(), 'application/pdf')}
        
        response = requests.post(API_URL, files=files_to_send)
        response.raise_for_status()
        
        converted_file_data = BytesIO(response.content)
        if converted_file_data.getbuffer().nbytes < 100:
             raise Exception("The converted file is empty, which may indicate an API issue.")

        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except requests.exceptions.HTTPError as e:
        return jsonify({"error": f"API Error: {e.response.text}"}), 500
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (FINAL PROFESSIONAL LAYOUT LOGIC) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_pages_rows_data = []
        for page in doc:
            # Step 1: Page par se saare lafz (words) unki jagah (coordinates) ke saath nikalein
            words = page.get_text("words")
            if not words: continue
            
            # Step 2: Columns ki a a a a pehchaan karein
            x_coords = sorted(list(set([round(w[0]) for w in words])))
            column_starts = [x_coords[0]]
            for i in range(1, len(x_coords)):
                # Agar do lafzon ke beech 10 pixel se zyada faasla hai to naya column samjho
                if x_coords[i] > x_coords[i-1] + 10:
                    column_starts.append(x_coords[i])

            # Step 3: Lafzon ko a a a a lines aur a a a a a columns mein arrange karein
            lines = {}
            for w in words:
                y0 = round(w[1])
                x0 = round(w[0])
                text = w[4]
                # Lafz ko a a a a a a uske sahi column mein daalein
                col_index = min(range(len(column_starts)), key=lambda i: abs(column_starts[i]-x0))
                
                if y0 not in lines:
                    lines[y0] = {i: "" for i in range(len(column_starts))}
                
                # Lafzon ko a a a a a jor kar jumle banayein
                if lines[y0][col_index] == "":
                    lines[y0][col_index] = text
                else:
                    lines[y0][col_index] += " " + text

            # Har page ke data ko list mein add karein
            for y in sorted(lines.keys()):
                all_pages_rows_data.append(list(lines[y].values()))

        doc.close()
        if not all_pages_rows_data:
            return jsonify({"error": "No text data could be extracted from this PDF."}), 400
        
        # Step 4: Saare pages ke data se ek a a a a a a DataFrame banayein
        df = pd.DataFrame(all_pages_rows_data)

        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Extracted_Data_All_Pages', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Excel file: {str(e)}"}), 500


# --- Other Tools (Word to PDF, Excel to PDF) - No Changes ---
# ... (Baaki ke dono tools waise hi rahenge) ...

if __name__ == '__main__':
    app.run(debug=True, port=5000)