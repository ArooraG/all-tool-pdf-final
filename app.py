# =========================================================================================
# == FINAL PROFESSIONAL VERSION V6.0 - Word on Cloudmersive API | Pro Excel Layout Logic ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess

# Cloudmersive Library (PDF to Word ke liye zaroori)
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (FINAL VERSION - Back to Professional Cloudmersive API) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # Render ke Environment se API key uthayega
    CLOUDMERSIVE_API_KEY = os.getenv('CLOUDMERSIVE_API_KEY')
    if not CLOUDMERSIVE_API_KEY:
        return jsonify({"error": "Cloudmersive API Key is not set on the server."}), 500

    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400
    
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        
        configuration = cloudmersive_convert_api_client.Configuration()
        configuration.api_key['Apikey'] = CLOUDMERSIVE_API_KEY
        api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))
        
        api_response = api_instance.convert_document_pdf_to_docx(input_path)
        
        converted_file_data = BytesIO(api_response)
        if converted_file_data.getbuffer().nbytes < 100:
             raise Exception("The converted file is empty. The API might have failed.")
             
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        os.remove(input_path)
        
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except ApiException as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"API Error: {e.body}"}), 500
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (FINAL VERSION - Advanced Column Logic) ---
# Yeh ab spaces ko dekh kar columns banayega
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
            words = page.get_text("words")
            if not words: continue
            
            # Lines ko group karein
            lines = {}
            for w in words:
                y0 = round(w[1])
                if y0 not in lines: lines[y0] = []
                lines[y0].append(w)
            
            # Har line ke andar, spaces ke hisaab se columns banayein
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0]) # left to right sort
                
                if not line_words: continue

                current_row = []
                current_cell_text = line_words[0][4]
                last_x1 = line_words[0][2]
                
                for i in range(1, len(line_words)):
                    word = line_words[i]
                    x0 = word[0]
                    text = word[4]
                    
                    # Agar space 10 pixels se zyada hai to naya cell
                    space_threshold = 10 
                    if x0 > last_x1 + space_threshold:
                        current_row.append(current_cell_text)
                        current_cell_text = text
                    else: # Warna usi cell mein jor do
                        current_cell_text += " " + text
                    
                    last_x1 = word[2]

                current_row.append(current_cell_text)
                all_pages_rows.append(current_row)

        doc.close()
        if not all_pages_rows:
            return jsonify({"error": "No text could be extracted from this PDF."}), 400

        df = pd.DataFrame(all_pages_rows)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Extracted_Data', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Excel file: {str(e)}"}), 500

# --- Baaki ke tools (No changes) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main(): return internal_word_to_pdf()
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main(): return internal_excel_to_pdf()
def internal_word_to_pdf():
    #... pehle jaisa code ...
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')): return jsonify({"error": "Invalid file type"}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion failed"}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500
def internal_excel_to_pdf():
    #... pehle jaisa code ...
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')): return jsonify({"error": "Invalid file type"}), 400
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=30)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion failed"}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)