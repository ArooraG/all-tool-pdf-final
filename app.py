# =========================================================================================
# == FINAL PROFESSIONAL VERSION V14.0 - Advanced Table Detection Logic ==
# =========================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO

# Word file banane ke liye
from docx import Document

app = Flask(__name__)
CORS(app)

# --- PDF to Word (Hybrid Method - Professional API Key Usage) ---
# Is code ko change karne ki zaroorat nahi. Yeh professional tareeqay se aapka API key istemal kar raha hai.
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF file."}), 400
    pdf_bytes = file.read()
    try:
        # Yeh environment variable se aapka secret key uthata hai, jo sab se mehfooz tareeka hai.
        CONVERTAPI_SECRET = os.getenv('CONVERTAPI_SECRET')
        if CONVERTAPI_SECRET:
            API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={CONVERTAPI_SECRET}'
            files_to_send = {'file': (file.filename, pdf_bytes, 'application/pdf')}
            response = requests.post(API_URL, files=files_to_send, timeout=60)
            response.raise_for_status()
            converted_file_data = BytesIO(response.content)
            if converted_file_data.getbuffer().nbytes > 100:
                docx_filename = os.path.splitext(file.filename)[0] + '.docx'
                return send_file(converted_file_data, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception: pass # Agar API fail ho, to neeche wala method istemal hoga
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        word_doc = Document()
        for page in doc: word_doc.add_paragraph(page.get_text("text")); word_doc.add_page_break()
        doc.close(); doc_buffer = BytesIO(); word_doc.save(doc_buffer); doc_buffer.seek(0)
        docx_filename = os.path.splitext(file.filename)[0] + '_fallback.docx'
        return send_file(doc_buffer, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e: return jsonify({"error": f"Dono tareeqay fail ho gaye. Ghalati: {str(e)}"}), 500


# --- PDF to Excel Tool (NEW ADVANCED LOGIC) ---
# Yeh naya code tables ko aachi tarah samajh kar data nikalega.
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_dfs = [] # Sab pages ke data ke liye
        for page_num, page in enumerate(doc):
            # Step 1: PDF se aasal tables dhoondhein
            tabs = page.find_tables()
            if tabs.tables: # Agar page par aasal tables milein
                for tab in tabs:
                    df = tab.to_pandas()
                    all_dfs.append(df)
            else:
                # Step 2: Agar table na milay, to text ko behtar tareeqay se nikalien
                blocks = page.get_text("dict")["blocks"]
                page_data = []
                for b in blocks:
                    if b['type'] == 0: # 0 matlab text block
                        for l in b["lines"]:
                            line_text = " ".join([s["text"] for s in l["spans"]])
                            # Har line ko alag cell mein daalein taake data saaf rahe
                            page_data.append([line_text])
                if page_data:
                    df = pd.DataFrame(page_data)
                    all_dfs.append(df)
        
        doc.close()
        
        if not all_dfs:
            return jsonify({"error": "PDF se koi data nahi nikala ja saka."}), 400

        # Sab data ko ek Excel file mein alag-alag sheets par ya ek he sheet par daal dein
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            # Sab data ko ek he sheet mein daal dein
            final_df = pd.concat(all_dfs, ignore_index=True)
            final_df.to_excel(writer, sheet_name='Extracted_Data', index=False, header=False)

        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return jsonify({"error": f"Excel banatay waqt ghalati hui: {str(e)}"}), 500


# --- Other Tools (In par koi tabdeeli nahi ki gayi) ---
@app.route('/word-to-pdf', methods=['POST'], endpoint='word_to_pdf_main')
def word_to_pdf_main():
     # Aap yahan ConvertAPI ya doosra logic daal sakte hain
    return jsonify({"error": "Yeh feature abhi nahi banaya gaya."}), 501

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    # Aap yahan ConvertAPI ya doosra logic daal sakte hain
    return jsonify({"error": "Yeh feature abhi nahi banaya gaya."}), 501


if __name__ == '__main__':
    app.run(debug=True, port=5000)