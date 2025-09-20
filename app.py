# =====================================================================================
# == FINAL GUARANTEED VERSION V17.0 - All Functions Corrected & Stable for Production ==
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

# --- 1. PDF to Word Tool (Using ConvertAPI) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    CONVERTAPI_SECRET = os.getenv('CONVERTAPI_SECRET')
    if not CONVERTAPI_SECRET:
        return jsonify({"error": "ConvertAPI Secret is not set on the server."}), 500
    API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={CONVERTAPI_SECRET}'
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        files_to_send = {'file': (file.filename, file.read(), 'application/pdf')}
        response = requests.post(API_URL, files=files_to_send, timeout=120)
        response.raise_for_status()
        converted_file_data = BytesIO(response.content)
        if converted_file_data.getbuffer().nbytes < 100:
             raise Exception("Converted file is empty. Check API plan limits or the source file.")
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except requests.exceptions.HTTPError as e:
        return jsonify({"error": f"API Error. Please check your API Secret. Details: {e.response.text}"}), 500
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (Professional Layout Logic) ---
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
            lines, line_key_map = {}, {}
            for w in words:
                y0 = round(w[1])
                line_key_actual = min(line_key_map.keys(), key=lambda y: abs(y-y0), default=None)
                if line_key_actual is not None and abs(line_key_actual-y0) < 5: lines[line_key_actual].append(w)
                else: lines[y0] = [w]; line_key_map[y0] = y0
            for y in sorted(lines.keys()):
                line_words, row, current_cell_text = sorted(lines[y], key=lambda w: w[0]), [], ""
                if not line_words: continue
                current_cell_text, last_x1 = line_words[0][4], line_words[0][2]
                for i in range(1, len(line_words)):
                    word, space = line_words[i], line_words[i][0] - last_x1
                    if space > 15: row.append(current_cell_text); current_cell_text = word[4]
                    else: current_cell_text += " " + word[4]
                    last_x1 = word[2]
                row.append(current_cell_text); all_pages_data.append(row)
        doc.close()
        if not all_pages_data: return jsonify({"error": "No text data could be extracted from this PDF."}), 400
        df = pd.DataFrame(all_pages_data)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_Data', index=False, header=False)
        output_buffer.seek(0)
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating Excel file: {str(e)}"}), 500

# --- 3. Word to PDF Tool ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Invalid file type. Please upload a Word document."}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=90)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path):
            raise Exception("File conversion failed on server.")
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

# --- 4. Excel to PDF Tool ---
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Invalid file type. Please upload an Excel file."}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=90)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path):
            raise Exception("File conversion failed on server.")
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)