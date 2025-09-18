# =========================================================================
# == FINAL VERSION 3.0 - Word conversion changed to LibreOffice, Excel logic improved ==
# =========================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import fitz  # PyMuPDF library
from io import BytesIO
import subprocess

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (CHANGED: AB API ISTEMAL NAHI HOGA) ---
# Yeh ab server par maujood LibreOffice se convert karega
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files:
        return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        
        # LibreOffice command to convert PDF to DOCX
        # Note: 'docx' format works best
        subprocess.run(
            ['soffice', '--headless', '--infilter="writer_pdf_import"', '--convert-to', 'docx', '--outdir', UPLOAD_FOLDER, input_path], 
            check=True, 
            timeout=60 # 60 seconds ka timeout, taake server hang na ho
        )
        
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        docx_path = get_safe_filepath(docx_filename)
        
        if not os.path.exists(docx_path):
            os.remove(input_path)
            return jsonify({"error": "Conversion failed. File bohot mushkil ho sakti hai."}), 500

        response = send_file(docx_path, as_attachment=True, download_name=docx_filename)
        
        # Files ko delete kar dein
        os.remove(input_path)
        os.remove(docx_path)
        
        return response
    except subprocess.TimeoutExpired:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": "File convert hone mein bohot time lag raha hai. Shayad file bohot badi ya complex hai."}), 500
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"Koi masla ho gaya: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (IMPROVED LOGIC - EXTRACTS ALL TEXT) ---
# Ab yeh table ke alawa saara text bhi uthayega
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']

    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_text_data = []
        for page in doc:
            # Table ke bajaye, hum page par likha saara text nikalenge,
            # uske coordinates (jagah) ke saath.
            blocks = page.get_text("blocks")
            for b in blocks:
                # b[0], b[1] -> x0, y0 (start coordinates)
                # b[2], b[3] -> x1, y1 (end coordinates)
                # b[4] -> text
                # Hum har text block ko ek line ke taur par save kar rahe hain
                all_text_data.append(b[4].replace('\n', ' ').strip())

        doc.close()

        if not all_text_data:
            return jsonify({"error": "Is PDF se koi text nahi nikal saka."}), 400
        
        # Ab in lines ko DataFrame bana dein
        df = pd.DataFrame(all_text_data, columns=["Extracted Text"])

        output_buffer = BytesIO()
        df.to_excel(output_buffer, sheet_name='Extracted_Data', index=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, 
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return jsonify({"error": f"Excel banate waqt masla hua: {str(e)}"}), 500


# --- 3. Word to PDF & 4. Excel to PDF Tools (No Changes) ---
# ... (in dono functions mein koi badlaav nahi hai, yeh waise hi rahenge) ...

@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    return word_to_pdf() # Avoid duplicate endpoint names

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    return excel_to_pdf()

# ... Original functions with dummy renaming for internal use ...

@app.route('/internal-word-to-pdf', methods=['POST'])
def internal_word_to_pdf():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Invalid file type"}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename, pdf_path = os.path.splitext(file.filename)[0] + '.pdf', get_safe_filepath(os.path.splitext(file.filename)[0] + '.pdf')
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion failed"}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"{e}"}), 500

@app.route('/internal-excel-to-pdf', methods=['POST'])
def internal_excel_to_pdf():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Invalid file type"}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename, pdf_path = os.path.splitext(file.filename)[0] + '.pdf', get_safe_filepath(os.path.splitext(file.filename)[0] + '.pdf')
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion failed"}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"{e}"}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)