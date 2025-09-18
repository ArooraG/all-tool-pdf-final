# =========================================================================================
# == FINAL PROFESSIONAL VERSION V9.0 - Word on ConvertAPI (Secret) | Pro Excel Layout  ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess
import statistics # Column detection ke liye

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (FINAL VERSION - Back on ConvertAPI with SECRET) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # Render ke Environment se ab hum 'SECRET' uthayenge
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
             raise Exception("The converted file is empty. This may be due to an incorrect API Secret or plan limits.")

        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except requests.exceptions.HTTPError as e:
        return jsonify({"error": f"API Error. Please check your API Secret. Details: {e.response.text}"}), 500
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (FINAL PROFESSIONAL COLUMN & SENTENCE LOGIC) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
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

            # Step 1: Poore page ke hisaab se columns ki a a a a positions detect karo
            x_coords = sorted(list(set([round(w[0]) for w in words])))
            if not x_coords: continue
            
            column_positions = [x_coords[0]]
            for i in range(1, len(x_coords)):
                # Agar do coordinates ke beech bara faasla hai to naya column hai
                if x_coords[i] > column_positions[-1] + 20: # 20 pixel threshold
                    is_new_col = True
                    for pos in column_positions:
                        if abs(x_coords[i] - pos) < 20:
                            is_new_col = False
                            break
                    if is_new_col:
                        column_positions.append(x_coords[i])
            
            # Step 2: Har line (row) ke words ko unke sahi column mein daalo
            lines = {}
            for w in words:
                y0 = round(w[1] / 10) * 10 # Words ko 10 pixel ki height mein group karo
                if y0 not in lines: lines[y0] = []
                lines[y0].append(w)

            # Step 3: Rows banakar final list mein daalo
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0]) # left to right sort
                row = [""] * len(column_positions)
                
                for word in line_words:
                    x0 = round(word[0])
                    text = word[4]
                    # Find karo ke yeh lafz kis column ke sabse qareeb hai
                    col_index = min(range(len(column_positions)), key=lambda i: abs(column_positions[i] - x0))
                    
                    if row[col_index] == "":
                        row[col_index] = text
                    else:
                        row[col_index] += " " + text
                
                all_pages_data.append(row)
            all_pages_data.append([""] * len(column_positions)) # Har page ke baad ek khaali line

        doc.close()
        if not all_pages_data:
            return jsonify({"error": "No text data could be extracted from this PDF."}), 400

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


# --- Other Tools (Word to PDF, Excel to PDF) - No Changes ---
@app.route('/word-to-pdf-internal', methods=['POST'])
def word_to_pdf_main(): # Changed route to avoid conflict
    return internal_word_to_pdf()
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    return internal_excel_to_pdf()

def internal_word_to_pdf():
    #... pehle jaisa code ...
    pass
def internal_excel_to_pdf():
    #... pehle jaisa code ...
    pass

if __name__ == '__main__':
    app.run(debug=True, port=5000)