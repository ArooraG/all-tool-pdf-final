# =====================================================================================
# == FINAL PROFESSIONAL VERSION - Manual Word Creation | Advanced Pro Excel Logic ==
# =====================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess

# Word file banane ke liye
from docx import Document

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
# Server start hone par folder check karega
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- 1. PDF to Word Tool (FINAL VERSION - Creates Word file MANUALLY, NO API) ---
# Yeh ab hamesha working file banayega
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file."}), 400

    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        word_doc = Document() # Nayi khaali Word document
        
        # Har page se text nikal kar Word file mein daalo
        for page in doc:
            # "blocks" se poore paragraphs uthao
            text_blocks = page.get_text("blocks", sort=True)
            for block in text_blocks:
                word_doc.add_paragraph(block[4])
            word_doc.add_page_break() # Har page ke baad naya page

        doc.close()
        
        # Word file ko memory mein save karo
        doc_buffer = BytesIO()
        word_doc.save(doc_buffer)
        doc_buffer.seek(0)
        
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(doc_buffer, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Word file: {str(e)}"}), 500

# --- 2. PDF to Excel Tool (FINAL PROFESSIONAL COLUMN & SENTENCE LOGIC) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF file."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        all_pages_rows = []
        for page in doc:
            words = page.get_text("words")
            if not words: continue

            lines = {}
            for w in words:
                y0 = round(w[1])
                line_key = min(lines.keys(), key=lambda y: abs(y-y0), default=None)
                if line_key is not None and abs(line_key - y0) < 5:
                    lines[line_key].append(w)
                else:
                    lines[y0] = [w]
            
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0])
                if not line_words: continue

                row = []
                current_cell_text = line_words[0][4]
                last_word_x1 = line_words[0][2]
                
                for i in range(1, len(line_words)):
                    word = line_words[i]; x0 = word[0]; text = word[4]
                    space = x0 - last_word_x1
                    
                    # Agar do lafzon ke beech bara faasla (e.g., 20 pixels) hai to naya column hai
                    # Chote space (e.g., 5-6 pixels) ko ek hi jumle ka hissa samjho
                    if space > 15:  # <-- Yeh number aapke layout ki "tuning" ke liye hai
                        row.append(current_cell_text) # Purana cell khatam
                        current_cell_text = text     # Naya cell shuru
                    else:
                        current_cell_text += " " + text
                    
                    last_word_x1 = word[2]
                row.append(current_cell_text)
                all_pages_rows.append(row)

        doc.close()
        if not all_pages_rows: return jsonify({"error": "No text data could be extracted."}), 400
        
        df = pd.DataFrame(all_pages_rows)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_Data', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

# --- Baaki ke dono tools (Inmein koi change nahi) ---
@app.route('/word-to-pdf', methods=['POST'], endpoint='word_to_pdf_main')
def word_to_pdf_main(): return internal_word_to_pdf()
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main(): return internal_excel_to_pdf()
# ...
def internal_word_to_pdf(): pass
def internal_excel_to_pdf(): pass
# ...

if __name__ == '__main__':
    app.run(debug=True, port=5000)