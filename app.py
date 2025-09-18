# ======================================================================================
# == FINAL PROFESSIONAL VERSION - Using Cloudmersive for Word & Advanced Logic for Excel ==
# ======================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess
import time

# Cloudmersive Library Imports (PDF to Word ke liye)
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (USING POWERFUL 'CLOUDMERSIVE' API) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # Iske liye Cloudmersive ki API Key zaroori hai
    CLOUDMERSIVE_API_KEY = os.getenv('CLOUDMERSIVE_API_KEY')
    if not CLOUDMERSIVE_API_KEY:
        return jsonify({"error": "Cloudmersive API Key server par set nahi hai."}), 500

    if 'file' not in request.files: return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400
    
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        
        # Cloudmersive API ko configure karein
        configuration = cloudmersive_convert_api_client.Configuration()
        configuration.api_key['Apikey'] = CLOUDMERSIVE_API_KEY
        api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))
        
        # Conversion ka process
        api_response = api_instance.convert_document_pdf_to_docx(input_path)
        
        # Response ko file ki tarah bhejein
        converted_file_data = BytesIO(api_response)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        os.remove(input_path)
        
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except ApiException as e:
        os.remove(input_path)
        return jsonify({"error": f"API se error aaya: {e.body}"}), 500
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"Koi masla ho gaya: {str(e)}"}), 500

# --- 2. PDF to Excel Tool (ADVANCED LOGIC - Preserves Layout) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        writer = pd.ExcelWriter(BytesIO(), engine='xlsxwriter')

        for page_num, page in enumerate(doc):
            # Page par se saare lafz unki jagah (coordinates) ke saath nikalein
            words = page.get_text("words")
            if not words: continue
                
            # Har lafz ko uski line (y-coordinate) ke hisaab se group karein
            lines = {}
            for w in words:
                y0 = int(w[1])
                if y0 not in lines: lines[y0] = []
                lines[y0].append(w[4]) # sirf text save karein

            # Data ko a columns mein daal kar ek DataFrame banayein
            df = pd.DataFrame(list(lines.values()))
            df.to_excel(writer, sheet_name=f'Page_{page_num + 1}', index=False, header=False)

        doc.close()
        
        # Excel file ko a save karke memory mein rakhein
        writer.close()
        excel_data = writer.book.filename
        excel_data.seek(0)

        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(excel_data, as_attachment=True, download_name=excel_filename, 
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return jsonify({"error": f"Excel banate waqt masla hua: {str(e)}"}), 500

# Baaki ke dono tools waise hi rahenge, woh theek kaam kar rahe hain
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Sahi Word file upload karein."}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion failed."}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"Koi masla ho gaya: {str(e)}"}), 500

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Sahi Excel file upload karein."}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion failed."}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"Koi masla ho gaya: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
