# =========================================================================================
# == FINAL PROFESSIONAL VERSION V16.0 - Robustness & Scanned PDF Detection ==
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

# --- PDF to Word (FIXED: Ab Corrupt file nahi banegi) ---
# NOTE: Server par CONVERTAPI_SECRET environment variable set karna LAZMI hai.
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF file."}), 400
    pdf_bytes = file.read()
    
    # PEHLA TAREEQA (ASAL): ConvertAPI (Images aur Formatting ke saath)
    try:
        CONVERTAPI_SECRET = os.getenv('CONVERTAPI_SECRET')
        if not CONVERTAPI_SECRET:
            raise Exception("ConvertAPI secret key server par set nahi hai.")
            
        API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={CONVERTAPI_SECRET}'
        files_to_send = {'file': (file.filename, pdf_bytes, 'application/pdf')}
        response = requests.post(API_URL, files=files_to_send, timeout=120)
        response.raise_for_status() 
        
        converted_file_data = BytesIO(response.content)
        if converted_file_data.getbuffer().nbytes > 100:
            docx_filename = os.path.splitext(file.filename)[0] + '.docx'
            return send_file(converted_file_data, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        else:
             raise Exception("ConvertAPI ne khaali file bheji.")
             
    except Exception as api_error:
        # DOOSRA TAREEQA (BACKUP): Sirf Text Nikalna (Ab pehle se zyada behtar)
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            full_text = ""
            for page in doc:
                full_text += page.get_text("text") + "\n"
            
            # *** YEH HAI ASAL WORD FILE FIX ***
            # Agar text khaali nahi hai, tabhi file banayein
            if full_text.strip():
                 word_doc = Document()
                 word_doc.add_paragraph(full_text)
                 doc_buffer = BytesIO()
                 word_doc.save(doc_buffer)
                 doc_buffer.seek(0)
                 docx_filename = os.path.splitext(file.filename)[0] + '_text_only.docx'
                 return send_file(doc_buffer, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            else:
                # Agar text nahi mila to corrupt file banane ke bajaye error dein
                return jsonify({"error": f"API conversion fail hui. Backup तरीक़ے se is PDF mein koi text nahi mila, shayad yeh ek scanned file hai."}), 500
        except Exception as fallback_error:
            return jsonify({"error": f"Dono tareeqay fail ho gaye. API Error: {str(api_error)}, Fallback Error: {str(fallback_error)}"}), 500

# --- PDF to Excel Tool (FIXED: Scanned PDF ka pata lagayega) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # *** YEH HAI SCANNED PDF KA PATA LAGANE WALA HISSA ***
        total_text_length = 0
        has_images = False
        for page in doc:
            total_text_length += len(page.get_text("text"))
            if page.get_images(full=True):
                has_images = True
        
        # Agar text bohot kam hai aur images hain, to yeh scanned PDF hai
        if has_images and total_text_length < (100 * doc.page_count): # Har page par 100 characters se kam
            return jsonify({"error": "Yeh file aek scanned image jaisi lagti hai. Hamara tool abhi tasveeron se text nahi nikal sakta."}), 400

        all_dfs = [] 
        for page in doc:
            tabs = page.find_tables()
            if tabs.tables: 
                for tab in tabs:
                    df = tab.to_pandas()
                    if not df.empty:
                        all_dfs.append(df)
            else:
                words = page.get_text("words")
                if not words: continue
                df = pd.DataFrame(page.get_text("blocks"))
                if not df.empty:
                    all_dfs.append(df)

        doc.close()
        
        if not all_dfs:
            return jsonify({"error": "Is PDF se koi table ya text nahi nikala ja saka."}), 400

        final_df = pd.concat(all_dfs, ignore_index=True)
        
        if final_df.empty:
            return jsonify({"error": "PDF se data to mila lekin Excel file khaali ban rahi hai. Shayad file format theek nahi."}), 400
            
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Extracted_Data', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"Excel banatay waqt ek ghalati hui: {str(e)}"}), 500


# --- Baaqi tools (In par koi kaam nahi kiya gaya) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    return jsonify({"error": "Yeh feature abhi nahi banaya gaya."}), 501
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    return jsonify({"error": "Yeh feature abhi nahi banaya gaya."}), 501

if __name__ == '__main__':
    app.run(debug=True, port=5000)