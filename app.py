# =========================================================================================
# == FINAL PROFESSIONAL VERSION V15.0 - Corrected Routes and LibreOffice Commands ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
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

# --- 1. PDF to Word Tool (FINAL VERSION - Corrected LibreOffice Command) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400

    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        
        # Sahi command: PDF ko "import" filter ke saath DOCX mein convert karein
        subprocess.run(
            ['soffice', '--headless', '--infilter="writer_pdf_import"', '--convert-to', 'docx', '--outdir', UPLOAD_FOLDER, input_path], 
            check=True, timeout=90
        )
        
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        docx_path = get_safe_filepath(docx_filename)
        
        if not os.path.exists(docx_path) or os.path.getsize(docx_path) < 100:
            os.remove(input_path)
            return jsonify({"error": "Conversion failed. The file might be too complex."}), 500

        response = send_file(docx_path, as_attachment=True, download_name=docx_filename)
        os.remove(input_path); os.remove(docx_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

# --- 2. PDF to Excel Tool (FINAL LOGIC - No Changes) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    # ... Iska code pehle hi theek tha, usko nahi chherenge ...
    # (Full code neeche for reference)
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        all_pages_data = []
        for page in doc:
            words = page.get_text("words")
            if not words: continue
            lines = {}
            for w in words:
                y0 = round(w[1])
                line_key = min(lines.keys(), key=lambda y: abs(y-y0), default=None)
                if line_key is not None and abs(line_key - y0) < 5: lines[line_key].append(w)
                else: lines[y0] = [w]
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0])
                row = []
                current_cell = ""
                if not line_words: continue
                current_cell = line_words[0][4]
                last_x1 = line_words[0][2]
                for i in range(1, len(line_words)):
                    word = line_words[i]
                    space = word[0] - last_x1
                    if space > 15: row.append(current_cell); current_cell = word[4]
                    else: current_cell += " " + word[4]
                    last_x1 = word[2]
                row.append(current_cell)
                all_pages_data.append(row)
        doc.close()
        if not all_pages_data: return jsonify({"error": "No text data extracted."}), 400
        df = pd.DataFrame(all_pages_data)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Data_Sheet', index=False, header=False)
        output_buffer.seek(0)
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error creating Excel file: {str(e)}"}), 500

# --- 3. Word to PDF & 4. Excel to PDF Tools (Corrected Route Names) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.doc', '.docx')):
        return jsonify({"error": "Invalid file type. Please upload a Word file."}), 400
    
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=90)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion to PDF failed."}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Invalid file type. Please upload an Excel file."}), 400
    
    input_path = get_safe_filepath(file.filename)
    try:
        file.save(input_path)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path], check=True, timeout=90)
        pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
        pdf_path = get_safe_filepath(pdf_filename)
        if not os.path.exists(pdf_path): return jsonify({"error": "Conversion to PDF failed."}), 500
        response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
        os.remove(input_path); os.remove(pdf_path)
        return response
    except Exception as e:
        if os.path.exists(input_path): os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)