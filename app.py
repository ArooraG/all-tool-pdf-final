# =====================================================================================
# == FINAL PRODUCTION-READY VERSION V25.0 - Explicit Command & Detailed Logging ==
# =====================================================================================

from flask import Flask, request, send_file, jsonify, after_this_request
from flask_cors import CORS
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess
from werkzeug.utils import secure_filename
import mimetypes

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    safe_filename = secure_filename(filename)
    return os.path.join(UPLOAD_FOLDER, safe_filename)

# --- PDF to Word (Using a more robust LibreOffice command) ---
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    if 'file' not in request.files: return jsonify({"error": "No file part."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Invalid file type. Please upload a PDF."}), 400
    # We will specify the docx filter to be more explicit
    return convert_with_libreoffice(file, "docx", "writer_pdf_import")

# --- PDF to Excel (Remains the same, it's efficient) ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    # ... (code remains the same)
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        all_pages_data = []
        for page in doc:
            words = page.get_text("words");
            if not words: continue
            lines, line_key_map = {}, {}
            for w in words:
                y0 = round(w[1]); line_key = min(line_key_map.keys(), key=lambda y: abs(y - y0), default=None)
                if line_key and abs(line_key - y0) < 5: lines[line_key].append(w)
                else: lines[y0] = [w]; line_key_map[y0] = y0
            for y in sorted(lines.keys()):
                line_words = sorted(lines[y], key=lambda w: w[0])
                if not line_words: continue
                row = []; current_cell = line_words[0][4]; last_x1 = line_words[0][2]
                for i in range(1, len(line_words)):
                    word = line_words[i]; space = word[0] - last_x1
                    if space > 15: row.append(current_cell); current_cell = word[4]
                    else: current_cell += " " + word[4]
                    last_x1 = word[2]
                row.append(current_cell); all_pages_data.append(row)
        doc.close()
        if not all_pages_data: return jsonify({"error": "No text data was extracted from the PDF."}), 400
        df = pd.DataFrame(all_pages_data); output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_Data', index=False, header=False)
        output_buffer.seek(0)
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({"error": f"An error occurred during Excel conversion: {str(e)}"}), 500

# --- Universal Converter Function (Final, Robust Version) ---
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith(('.doc', '.docx')): return jsonify({"error": "Invalid file type. Please upload a Word document."}), 400
    return convert_with_libreoffice(file, "pdf")

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_main():
    if 'file' not in request.files: return jsonify({"error": "No file part."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith(('.xls', '.xlsx')): return jsonify({"error": "Invalid file type. Please upload an Excel file."}), 400
    return convert_with_libreoffice(file, "pdf")

def convert_with_libreoffice(file, output_format, in_filter=None):
    input_path = get_safe_filepath(file.filename)
    output_path = None
    output_dir = os.path.abspath(UPLOAD_FOLDER) # Use absolute path for safety

    command = ['soffice', '--headless']
    # Agar in_filter dia gaya hai (jaise PDF to Word ke liye), to use shamil karein
    if in_filter:
        command.extend(['--infilter', in_filter])
    
    command.extend(['--convert-to', output_format, '--outdir', output_dir, input_path])
    
    try:
        file.save(input_path)
        print(f"Executing command: {' '.join(command)}") # Log the exact command

        result = subprocess.run(
            command,
            check=True, timeout=120, capture_output=True, text=True
        )

        print("LibreOffice stdout:", result.stdout) # Log success output
        print("LibreOffice stderr:", result.stderr) # Log any warnings

        output_filename = os.path.splitext(os.path.basename(input_path))[0] + f'.{output_format}'
        output_path = get_safe_filepath(output_filename)
        
        if not os.path.exists(output_path):
            raise Exception("Conversion failed: Output file was not created by LibreOffice.")

        mimetype = mimetypes.guess_type(output_path)[0] or 'application/octet-stream'
        return send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)

    except subprocess.CalledProcessError as e:
        # Is se humein LibreOffice ka اصل error pata chalega
        error_details = e.stderr or e.stdout or "No detailed error from converter."
        print(f"LibreOffice failed with error: {error_details}") # Log the detailed error
        return jsonify({"error": f"Server-side conversion failed. Details: {error_details}"}), 500
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}") # Log unexpected errors
        return jsonify({"error": f"An unexpected server error occurred: {str(e)}"}), 500
    finally:
        # Yeh file cleanup ka sab se behtar tareeqa hai
        try:
            if input_path and os.path.exists(input_path):
                os.remove(input_path)
            if output_path and os.path.exists(output_path):
                os.remove(output_path)
        except OSError as e:
            print(f"Error during file cleanup: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)