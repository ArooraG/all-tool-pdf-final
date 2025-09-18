# =========================================================================================
# == FINAL PROFESSIONAL VERSION V4.0 - All issues fixed and English Errors ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import pandas as pd
import fitz
from io import BytesIO
import subprocess
import time
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (FINAL HYBRID APPROACH: API first, then Fallback) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file part received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Invalid file. Please upload a valid PDF."}), 400

    input_path = get_safe_filepath(file.filename)
    file.save(input_path)

    # --- Step 1: Try with Cloudmersive API First ---
    try:
        CLOUDMERSIVE_API_KEY = os.getenv('CLOUDMERSIVE_API_KEY')
        if CLOUDMERSIVE_API_KEY:
            configuration = cloudmersive_convert_api_client.Configuration()
            configuration.api_key['Apikey'] = CLOUDMERSIVE_API_KEY
            api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))
            api_response = api_instance.convert_document_pdf_to_docx(input_path)
            
            converted_file_data = BytesIO(api_response)
            if converted_file_data.getbuffer().nbytes > 100: # Check if file is not empty
                docx_filename = os.path.splitext(file.filename)[0] + '.docx'
                os.remove(input_path)
                return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                                 mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception:
        # If API fails, we will proceed to the fallback method below.
        pass

    # --- Step 2: Fallback to LibreOffice if API fails ---
    try:
        subprocess.run(['soffice', '--headless', '--infilter="writer_pdf_import"', '--convert-to', 'docx', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=60)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        docx_path = get_safe_filepath(docx_filename)
        if not os.path.exists(docx_path): raise Exception("Conversion with LibreOffice failed.")
        response = send_file(docx_path, as_attachment=True, download_name=docx_filename)
        os.remove(input_path); os.remove(docx_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"Both API and server conversion failed. The file may be too complex. Details: {str(e)}"}), 500

# --- 2. PDF to Excel Tool (FIXED 'xlsxwriter' error and improved layout) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        output_buffer = BytesIO()
        writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')

        for page_num, page in enumerate(doc):
            words = page.get_text("words")
            if not words: continue
            
            df = pd.DataFrame(words, columns=['x0', 'y0', 'x1', 'y1', 'text', 'block_no', 'line_no', 'word_no'])
            df['y0'] = df['y0'].round()
            
            lines = df.groupby('y0')['text'].apply(lambda x: ' '.join(x)).reset_index()
            final_df = lines['text'].str.split(r'\s{2,}', expand=True) # Split columns by 2 or more spaces

            if not final_df.empty:
                final_df.to_excel(writer, sheet_name=f'Page_{page_num + 1}', index=False, header=False)
        doc.close()

        if len(writer.sheets) == 0:
            return jsonify({"error": "No text could be extracted from this PDF."}), 400
        
        writer.close()
        output_buffer.seek(0)
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Excel file: {str(e)}"}), 500


# --- Other Tools (No changes needed) ---

@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    return internal_word_to_pdf()
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    return internal_excel_to_pdf()

def internal_word_to_pdf():
    # ... code is same ...
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')): return jsonify({"error": "Invalid file type"}), 400
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
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

def internal_excel_to_pdf():
    # ... code is same ...
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')): return jsonify({"error": "Invalid file type"}), 400
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
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)