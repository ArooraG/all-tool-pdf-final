# =====================================================================================
# == FINAL STABLE PRODUCTION VERSION V30.0 - Optimized Excel Logic ==
# =====================================================================================

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import subprocess
from werkzeug.utils import secure_filename
import mimetypes
from docx import Document

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    safe_filename = secure_filename(filename)
    return os.path.join(UPLOAD_FOLDER, safe_filename)

# --- PDF to Word METHOD 1: High-Quality (LibreOffice) ---
@app.route('/pdf-to-word-premium', methods=['POST'])
def pdf_to_word_premium():
    if 'file' not in request.files: return jsonify({"error": "No file part."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Invalid file type. Please upload a PDF."}), 400
    return convert_with_libreoffice(file, "docx")

# --- PDF to Word METHOD 2: Basic Fallback (In-Memory) ---
@app.route('/pdf-to-word-basic', methods=['POST'])
def pdf_to_word_basic():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        word_doc = Document()
        for page in doc:
            word_doc.add_paragraph(page.get_text("text"))
        doc.close()
        doc_buffer = BytesIO()
        word_doc.save(doc_buffer)
        doc_buffer.seek(0)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(doc_buffer, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return jsonify({"error": f"An error occurred while creating the Word file: {str(e)}"}), 500

# --- Other Converters ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        all_pages_data = []

        for page in doc:
            # Get words and group them by lines more efficiently
            words = page.get_text("words")
            if not words: continue

            # Group words by their vertical position (y0)
            lines = {}
            for w in words:
                y0 = round(w[1])
                lines.setdefault(y0, []).append(w)

            # Sort lines by vertical position
            sorted_lines = sorted(lines.items())
            
            # Merge nearby lines
            merged_lines = []
            if sorted_lines:
                current_line_words = sorted_lines[0][1]
                last_y = sorted_lines[0][0]
                for y, words_in_line in sorted_lines[1:]:
                    if abs(y - last_y) < 5:  # Tolerance for merging lines
                        current_line_words.extend(words_in_line)
                    else:
                        merged_lines.append(sorted(current_line_words, key=lambda w: w[0]))
                        current_line_words = words_in_line
                    last_y = y
                merged_lines.append(sorted(current_line_words, key=lambda w: w[0]))

            # Process each merged line to form rows
            for line_words in merged_lines:
                if not line_words: continue
                row = []
                current_cell = line_words[0][4]
                last_x1 = line_words[0][2]
                for i in range(1, len(line_words)):
                    word = line_words[i]
                    space = word[0] - last_x1
                    if space > 15:  # Space threshold to define a new cell
                        row.append(current_cell)
                        current_cell = word[4]
                    else:
                        current_cell += " " + word[4]
                    last_x1 = word[2]
                row.append(current_cell)
                all_pages_data.append(row)

        doc.close()
        
        if not all_pages_data: return jsonify({"error": "No text data was extracted from the PDF."}), 400
        
        df = pd.DataFrame(all_pages_data)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_Data', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        # For debugging, you can print the error
        print(f"Error during Excel conversion: {str(e)}")
        return jsonify({"error": f"An error occurred during Excel conversion: {str(e)}"}), 500


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

# --- Universal LibreOffice Function ---
def convert_with_libreoffice(file, output_format):
    input_path = get_safe_filepath(file.filename)
    output_path = None
    output_dir = os.path.abspath(UPLOAD_FOLDER)
    # Define a unique user profile directory to avoid conflicts
    user_profile_dir = os.path.abspath(os.path.join(UPLOAD_FOLDER, 'libreoffice_profile'))
    if not os.path.exists(user_profile_dir):
        os.makedirs(user_profile_dir)
    
    user_profile_arg = f"-env:UserInstallation=file://{user_profile_dir}"
    
    command = ['soffice', user_profile_arg, '--headless', '--convert-to', output_format, '--outdir', output_dir, input_path]
    try:
        file.save(input_path)
        # Increased timeout for potentially large files on slow servers
        result = subprocess.run(command, check=True, timeout=180, capture_output=True, text=True)
        
        output_filename = os.path.splitext(os.path.basename(input_path))[0] + f'.{output_format.split(":")[0]}'
        output_path = get_safe_filepath(output_filename)
        
        if not os.path.exists(output_path):
            # Provide more detailed error logging on the server side
            print(f"LibreOffice command: {' '.join(command)}")
            print(f"LibreOffice stdout: {result.stdout}")
            print(f"LibreOffice stderr: {result.stderr}")
            raise Exception("Output file was not created by LibreOffice. Check server logs for details.")
            
        mimetype = mimetypes.guess_type(output_path)[0] or 'application/octet-stream'
        return send_file(output_path, as_attachment=True, download_name=output_filename, mimetype=mimetype)
    except subprocess.TimeoutExpired:
        return jsonify({"error": "The conversion process took too long and was timed out. The file might be too large or complex."}), 504
    except Exception as e:
        print(f"An unexpected server error occurred: {str(e)}")
        return jsonify({"error": f"An unexpected server error occurred. The server might be busy or the file is not supported."}), 500
    finally:
        try:
            if input_path and os.path.exists(input_path): os.remove(input_path)
            if output_path and os.path.exists(output_path): os.remove(output_path)
        except OSError as e:
            print(f"Error during file cleanup: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)