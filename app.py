# =====================================================================================
# == FINAL PRODUCTION-READY VERSION V20.0 - All Libraries and Logic Stable ==
# =====================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess

app = Flask(__name__)
CORS(app) # Allow frontend to talk to backend

UPLOAD_FOLDER = 'uploads'
# Server start hone par folder check karega
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- PDF to Word (Manual Fallback Method, No API Needed) ---
# API keys ka jhanjhat hi khatam
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    from docx import Document # Isko function ke andar rakha hai to save memory
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        word_doc = Document()
        for page in doc:
            word_doc.add_paragraph(page.get_text("text"))
            word_doc.add_page_break()
        doc.close()
        doc_buffer = BytesIO()
        word_doc.save(doc_buffer)
        doc_buffer.seek(0)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(doc_buffer, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Word file: {str(e)}"}), 500

# --- PDF to Excel (Professional Layout Logic) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    # ... Iska code pehle hi behtareen hai ...
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        all_pages_data = []
        for page in doc:
            words = page.get_text("words");
            if not words: continue
            lines, line_key_map = {}, {}
            for w in words:
                y0 = round(w[1]); line_key = min(line_key_map.keys(), key=lambda y: abs(y - y0), default=None)
                if line_key and abs(line_key - y0) < 5: lines[line_key].append(w)
                else: lines[y0] = [w]; line_key_map[y0] = y0
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0])
                if not line_words: continue
                row = []; current_cell = line_words[0][4]; last_x1 = line_words[0][2]
                for i in range(1, len(line_words)):
                    word = line_words[i]; space = word[0] - last_x1
                    if space > 15: row.append(current_cell); current_cell = word[4]
                    else: current_cell += " " + word[4]
                    last_x1 = word[2]
                row.append(current_cell); all_pages_data.append(row)
        doc.close()
        if not all_pages_data: return jsonify({"error": "No text data was extracted from the PDF."}), 400
        df = pd.DataFrame(all_pages_data); output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_Data', index=False, header=False)
        output_buffer.seek(0)
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred during Excel conversion: {str(e)}"}), 500

# --- Word to PDF & Excel to PDF (Using LibreOffice on Render) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith(('.doc', '.docx')): return jsonify({"error": "Invalid file type. Please upload a Word document."}), 400
    return convert_with_libreoffice(file, "pdf")

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith(('.xls', '.xlsx')): return jsonify({"error": "Invalid file type. Please upload an Excel file."}), 400
    return convert_with_libreoffice(file, "pdf")

def convert_with_libreoffice(file, output_format):
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', output_format, '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=90)
        output_filename = os.path.splitext(file.filename)[0] + f'.{output_format}'
        output_path = get_safe_filepath(output_filename)
        if not os.path.exists(output_path): raise Exception("Conversion failed on the server.")
        response = send_file(output_path, as_attachment=True, download_name=output_filename)
        os.remove(input_path); os.remove(output_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    # Yeh port Render ke liye zaroori hai
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)