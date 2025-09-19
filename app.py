# =========================================================================================
# == FINAL PROFESSIONAL VERSION V15.0 - Robust Word & Excel Conversion ==
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

# --- PDF to Word (Hybrid Method - With clear Primary/Fallback logic) ---
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
            # Agar API key server par set nahi hai, to aage na barhein
            raise Exception("ConvertAPI secret key is not configured on the server.")
            
        API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={CONVERTAPI_SECRET}'
        files_to_send = {'file': (file.filename, pdf_bytes, 'application/pdf')}
        
        # Timeout barha diya gaya hai taake bari files bhi convert ho sakein
        response = requests.post(API_URL, files=files_to_send, timeout=120)
        response.raise_for_status() # Agar koi error ho (jaise 4xx, 5xx), to yeh ruk jayega
        
        converted_file_data = BytesIO(response.content)
        # Check karein ke file aasal mein aayi hai ya nahi
        if converted_file_data.getbuffer().nbytes > 100:
            docx_filename = os.path.splitext(file.filename)[0] + '.docx'
            return send_file(converted_file_data, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        else:
             raise Exception("ConvertAPI returned an empty file.")
             
    except Exception as api_error:
        # DOOSRA TAREEQA (BACKUP): Sirf Text Nikalna
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            word_doc = Document()
            full_text = ""
            for page in doc:
                full_text += page.get_text("text") + "\n"
            
            # Agar PDF mein text hai to file banayein
            if full_text.strip():
                 word_doc.add_paragraph(full_text)
                 doc_buffer = BytesIO()
                 word_doc.save(doc_buffer)
                 doc_buffer.seek(0)
                 docx_filename = os.path.splitext(file.filename)[0] + '_text_only.docx'
                 return send_file(doc_buffer, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            else:
                return jsonify({"error": f"API conversion failed: {str(api_error)}. The fallback method also could not find any text in the PDF."}), 500
        except Exception as fallback_error:
            # Agar dono tareeqay fail ho jayein
            return jsonify({"error": f"Both conversion methods failed. API Error: {str(api_error)}, Fallback Error: {str(fallback_error)}"}), 500

# --- PDF to Excel Tool (FIXED: Ab khaali file nahi aayegi) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_dfs = [] # Sab pages se dataframes (tables) jamaa karne ke liye
        for page in doc:
            tabs = page.find_tables()
            if tabs.tables: # Agar aasal table milay to...
                for tab in tabs:
                    df = tab.to_pandas()
                    if not df.empty: # Khaali table ko ignore karein
                        all_dfs.append(df)
            else: # Agar koi aasal table na milay, to text ko columns mein daalein
                words = page.get_text("words")
                if not words: continue
                
                lines = {}
                for w in words:
                    y0 = round(w[1])
                    line_key = min(lines.keys(), key=lambda y: abs(y - y0), default=None)
                    if line_key is not None and abs(line_key - y0) < 5: lines[line_key].append(w)
                    else: lines[y0] = [w]
                
                page_rows = []
                for y in sorted(lines.keys()):
                    line_words = sorted(lines[y], key=lambda w: w[0])
                    if not line_words: continue
                    
                    row = []
                    current_cell = line_words[0][4]
                    last_x1 = line_words[0][2]
                    for i in range(1, len(line_words)):
                        word = line_words[i]
                        space = word[0] - last_x1
                        if space > 10: # Space ke hisaab se naya column
                            row.append(current_cell)
                            current_cell = word[4]
                        else:
                            current_cell += " " + word[4]
                        last_x1 = word[2]
                    row.append(current_cell)
                    page_rows.append(row)
                
                if page_rows:
                    df = pd.DataFrame(page_rows)
                    all_dfs.append(df)
        doc.close()
        
        if not all_dfs:
            return jsonify({"error": "Is PDF se koi table ya text nahi nikala ja saka."}), 400

        # Sab dataframes ko ek DataFrame mein jamaa karein
        final_df = pd.concat(all_dfs, ignore_index=True)
        
        # **** YEH HAI ASAL FIX: Check karein ke final data khaali to nahi ****
        if final_df.empty:
            return jsonify({"error": "Data to mila lekin anjaam khaali hai. Shayad file mein ajeeb formatting hai."}), 400
            
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Extracted_Data', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        # Ghalati ki aasan wazaahat
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