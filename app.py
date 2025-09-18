# =========================================================================================
# == FINAL PROFESSIONAL VERSION - Cloudmersive for WORD | Advanced Layout Logic for EXCEL ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import pandas as pd
import fitz  # PyMuPDF library
from io import BytesIO
import subprocess
import time

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

# --- 1. PDF to Word Tool (FINAL VERSION - USING CLOUDMERSIVE API) ---
# Yeh professional API hai aur har tarah ki PDF ko handle karti hai
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # Render ke Environment se API key uthayega
    CLOUDMERSIVE_API_KEY = os.getenv('CLOUDMERSIVE_API_KEY')
    if not CLOUDMERSIVE_API_KEY:
        return jsonify({"error": "Cloudmersive API Key server par set nahi hai."}), 500

    if 'file' not in request.files: return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400
    
    # File ko a server par save karein API ko a bhejne ke liye a
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        
        # Cloudmersive API ko a set karein
        configuration = cloudmersive_convert_api_client.Configuration()
        configuration.api_key['Apikey'] = CLOUDMERSIVE_API_KEY
        api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))
        
        # File convert karein
        api_response = api_instance.convert_document_pdf_to_docx(input_path)
        
        # Response ko a file ki tarah download karwayein
        converted_file_data = BytesIO(api_response)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        
        # Temp file delete karein
        os.remove(input_path)
        
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except ApiException as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"API se error aaya: {e.body}"}), 500
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"Koi masla ho gaya: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (FINAL VERSION - ADVANCED LAYOUT LOGIC) ---
# Yeh ab layout/columns ko a sahi se banayega
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "File nahi mili."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Sahi PDF file upload karein."}), 400

    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Excel file banane ke liye a setup
        output_buffer = BytesIO()
        writer = pd.ExcelWriter(output_buffer, engine='xlsxwriter')

        for page_num, page in enumerate(doc):
            # Page par se saare lafz unki jagah (coordinates) ke saath nikalein
            words = page.get_text("words")
            if not words: continue
                
            # Har lafz (word) ko a (y-coordinate, x-coordinate, text) ki shakal mein save karein
            word_tuples = [(int(w[1]), int(w[0]), w[4]) for w in words]
            # Pehle y (line) ke hisaab se sort karein, phir x (left to right) ke hisaab se
            word_tuples.sort()

            # Ab in lafzon ko a lines mein jorein
            lines = {}
            for y, x, text in word_tuples:
                # Agar y-coordinate (line number) thora bohot (5 pixels tak) aage peeche hai to usko a ek hi line samjhein
                found_line = False
                for line_y in range(y - 5, y + 6):
                    if line_y in lines:
                        lines[line_y].append(text)
                        found_line = True
                        break
                if not found_line:
                    lines[y] = [text]
            
            # Lines ko a Excel ke rows mein convert karein
            all_rows = []
            for y in sorted(lines.keys()):
                all_rows.append(lines[y])

            if all_rows:
                df = pd.DataFrame(all_rows)
                # Har page ka data alag sheet mein daalein
                df.to_excel(writer, sheet_name=f'Page_{page_num + 1}', index=False, header=False)

        doc.close()
        
        if len(writer.sheets) == 0:
            return jsonify({"error": "Is PDF se koi text nahi nikal saka."}), 400
        
        writer.close()
        output_buffer.seek(0)

        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, 
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return jsonify({"error": f"Excel banate waqt masla hua: {str(e)}"}), 500


# --- 3. Word to PDF & 4. Excel to PDF Tools (Yeh pehle se theek hain) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    return internal_word_to_pdf()

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    return internal_excel_to_pdf()

def internal_word_to_pdf():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Invalid file type"}), 400
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
        return jsonify({"error": f"{e}"}), 500

def internal_excel_to_pdf():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Invalid file type"}), 400
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
        return jsonify({"error": f"{e}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)