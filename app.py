# =========================================================================================
# == FINAL PROFESSIONAL VERSION V12.0 - Hybrid Word Conversion | Pro Excel Column Logic ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess

# Word file banane ke liye
from docx import Document

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- 1. PDF to Word Tool (HYBRID METHOD: API First, Fallback to Manual) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    
    pdf_bytes = file.read() # File ko memory mein parhein

    # Tareeqa #1: ConvertAPI se koshish karein
    try:
        CONVERTAPI_SECRET = os.getenv('CONVERTAPI_SECRET')
        if CONVERTAPI_SECRET:
            API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={CONVERTAPI_SECRET}'
            files_to_send = {'file': (file.filename, pdf_bytes, 'application/pdf')}
            response = requests.post(API_URL, files=files_to_send, timeout=60)
            response.raise_for_status()
            
            converted_file_data = BytesIO(response.content)
            if converted_file_data.getbuffer().nbytes > 100: # Agar file khaali nahi hai
                docx_filename = os.path.splitext(file.filename)[0] + '.docx'
                return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                                 mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception:
        # Agar API fail ho jaye, to pareshan na hon, hum neeche doosra tareeqa istemal karenge
        pass

    # Tareeqa #2: Manual (Fallback) - Agar API fail ho
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        word_doc = Document()
        for page in doc:
            word_doc.add_paragraph(page.get_text("text"))
            word_doc.add_page_break()
        doc.close()
        
        doc_buffer = BytesIO()
        word_doc.save(doc_buffer)
        doc_buffer.seek(0)
        
        docx_filename = os.path.splitext(file.filename)[0] + '_fallback.docx'
        return send_file(doc_buffer, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({"error": f"Both conversion methods failed. Error: {str(e)}"}), 500

# --- 2. PDF to Excel Tool (FINAL PROFESSIONAL COLUMN LOGIC) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_pages_rows = []
        for page in doc:
            # Step 1: Saare lafzon (words) ko unki jagah ke saath nikalein
            words = page.get_text("words")
            if not words: continue

            # Step 2: Columns ki a a a a a a a a a positions detect karein
            # Hum page ke headers ko dekh kar column dhoondenge
            header_y_coord = 80 # Andaaza hai ke header 80 pixel ke aas paas hoga
            header_words = [w for w in words if abs(w[1] - header_y_coord) < 10]
            if not header_words:
                 header_words = [w for w in words if w[1] < 100] # Agar 80 par na mile to upar ke 100 pixel mein dekho

            if not header_words: # Agar phir bhi na mile to purana tareeqa
                x_coords = sorted(list(set([round(w[0]) for w in words])))
                column_starts = [x_coords[0]] if x_coords else []
            else:
                column_starts = sorted([round(w[0]) for w in header_words])
            
            # Step 3: Lafzon ko lines aur a a a a a a a a a a a columns mein arrange karein
            lines = {}
            for w in words:
                y_key = round(w[1] / 10) * 10
                if y_key not in lines: lines[y_key] = []
                lines[y_key].append(w)

            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0])
                row = [""] * len(column_starts)
                
                for word in line_words:
                    x0 = word[0]
                    # Pata lagao ke yeh lafz kis column ke sabse qareeb hai
                    col_index = min(range(len(column_starts)), key=lambda i: abs(column_starts[i] - x0))
                    if row[col_index] == "": row[col_index] = word[4]
                    else: row[col_index] += " " + word[4]
                
                all_pages_rows.append(row)
        
        doc.close()
        if not all_pages_rows:
            return jsonify({"error": "No text could be extracted."}), 400

        df = pd.DataFrame(all_pages_rows)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_In_One_Sheet', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating Excel: {str(e)}"}), 500

# --- Other Tools (Word to PDF, Excel to PDF) ---
@app.route('/word-to-pdf', methods=['POST'], endpoint='word_to_pdf_main')
def word_to_pdf_main(): return internal_word_to_pdf()
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main(): return internal_excel_to_pdf()
# ... baaki code ...
def internal_word_to_pdf(): pass
def internal_excel_to_pdf(): pass

if __name__ == '__main__':
    app.run(debug=True, port=5000)