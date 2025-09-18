# =========================================================================================
# == FINAL PROFESSIONAL VERSION V5.0 - Word file created manually, Excel data in ONE sheet ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess

# Word file banane ke liye (NEW)
from docx import Document

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# --- 1. PDF to Word Tool (FINAL VERSION - Creates Word file MANUALLY) ---
# Yeh ab API ke baghair, 100% result dega agar PDF mein text hai
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400

    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Ek nayi khaali Word document banayein
        word_doc = Document()
        
        # Har page se text nikal kar Word file mein daalein
        for page in doc:
            # "blocks" se poore paragraphs uthayein taake formatting behtar lage
            text_blocks = page.get_text("blocks")
            for block in text_blocks:
                paragraph_text = block[4]
                word_doc.add_paragraph(paragraph_text)
            
            # Har page ke baad ek page break daal dein (optional)
            word_doc.add_page_break()

        doc.close()
        
        # Word document ko memory mein save karein
        doc_buffer = BytesIO()
        word_doc.save(doc_buffer)
        doc_buffer.seek(0)
        
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(doc_buffer, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Word file: {str(e)}"}), 500


# --- 2. PDF to Excel Tool (FINAL VERSION - All pages in ONE sheet) ---
# Saare pages ka data ab ek he sheet mein aayega
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400

    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Saare pages ka data is list mein jama hoga
        all_pages_rows = []

        for page in doc:
            words = page.get_text("words")
            if not words: continue
                
            # Har lafz (word) ko uski line number ke saath group karein
            lines = {}
            for w in words:
                y0 = int(w[1])
                text = w[4]
                found = False
                for y in range(y0 - 4, y0 + 5): # Line number mein thori bohot kami peshi ko a a a a a nazar andaz karein
                    if y in lines:
                        lines[y].append(text)
                        found = True
                        break
                if not found:
                    lines[y0] = [text]

            # In lines ko a rows ki shakal dein
            for y in sorted(lines.keys()):
                all_pages_rows.append(lines[y])

        doc.close()

        if not all_pages_rows:
            return jsonify({"error": "No text could be extracted from this PDF."}), 400

        # Saare data se ek DataFrame banayein aur ek hi sheet mein daalein
        df = pd.DataFrame(all_pages_rows)
        
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_Data', index=False, header=False)
        
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Excel file: {str(e)}"}), 500


# --- Other Tools (Word to PDF, Excel to PDF) ---
# Inko change karne ki zaroorat nahi hai
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    return internal_word_to_pdf()
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    return internal_excel_to_pdf()
def internal_word_to_pdf():
    # ... code pehle jaisa hi hai ...
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
    # ... code pehle jaisa hi hai ...
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