from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
import subprocess
import os
import uuid
import logging
import pandas as pd
import pdfplumber
from pdf2docx import Converter

# Basic setup
logging.basicConfig(level=logging.INFO)
app = Flask(__name__, template_folder='templates')

# *** IMPORTANT FIX: This updated CORS configuration gives the server permission to accept files. ***
CORS(app, resources={r"/*": {"origins": "*"}})

# Folder to temporarily store uploaded files
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# --- Helper Functions ---
def get_unique_filepath(original_filename):
    """Creates a unique, safe filepath for a file."""
    ext = os.path.splitext(original_filename)[1]
    if not ext:
        ext = '.tmp'
    filename = str(uuid.uuid4()) + ext
    return os.path.join(app.config['UPLOAD_FOLDER'], filename)

def cleanup_files(*args):
    """Safely deletes temporary files after use."""
    for file_path in args:
        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
                app.logger.info(f"Successfully cleaned up file: {file_path}")
            except Exception as e:
                app.logger.error(f"Error deleting file {file_path}: {e}")

# --- Route to Serve the Main Website Page ---
@app.route('/')
def home():
    """This function loads and displays your index.html file."""
    return render_template('index.html')

# --- Backend Tool Routes ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel_tool():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400

    input_path = get_unique_filepath(file.filename)
    excel_filepath = os.path.splitext(input_path)[0] + '.xlsx'
    
    try:
        file.save(input_path)
        with pdfplumber.open(input_path) as pdf:
            with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
                if not pdf.pages:
                    raise ValueError("PDF has no pages.")
                found_table = False
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    if tables:
                        found_table = True
                        for j, table in enumerate(tables):
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df.to_excel(writer, sheet_name=f'Page_{i+1}_Table_{j+1}', index=False)
                if not found_table:
                    app.logger.warning("No tables found, extracting raw text as a fallback.")
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        if text:
                            df = pd.DataFrame([line.split() for line in text.split('\n')])
                            df.to_excel(writer, sheet_name=f'Page_{i+1}_Text', index=False, header=False)
        if not os.path.exists(excel_filepath):
             raise FileNotFoundError("Conversion to Excel failed: Output file not created.")
        return send_file(excel_filepath, as_attachment=True, download_name=os.path.basename(excel_filepath))
    except Exception as e:
        app.logger.error(f"PDF-to-Excel error: {e}")
        return jsonify({"error": f"An error occurred during conversion: {e}"}), 500
    finally:
        cleanup_files(input_path, excel_filepath)

@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word_tool():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400
    
    pdf_path = get_unique_filepath(file.filename)
    docx_path = os.path.splitext(pdf_path)[0] + '.docx'

    try:
        file.save(pdf_path)
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        if not os.path.exists(docx_path):
             raise FileNotFoundError("Conversion to DOCX failed.")
        return send_file(docx_path, as_attachment=True, download_name=os.path.basename(docx_path))
    except Exception as e:
        app.logger.error(f"PDF-to-Word error: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        cleanup_files(pdf_path, docx_path)

@app.route('/office-to-pdf', methods=['POST'])
def office_to_pdf_tool():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400
    
    input_path, pdf_filepath = None, None
    try:
        input_path = get_unique_filepath(file.filename)
        file.save(input_path)
        output_dir = app.config['UPLOAD_FOLDER']
        command = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, input_path]
        subprocess.run(command, check=True, timeout=300) 
        pdf_filename = os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
        pdf_filepath = os.path.join(output_dir, pdf_filename)
        if not os.path.exists(pdf_filepath): 
            raise FileNotFoundError("Conversion to PDF failed. Ensure LibreOffice is installed on the server.")
        return send_file(pdf_filepath, as_attachment=True, download_name=pdf_filename)
    except subprocess.TimeoutExpired:
        app.logger.error("LibreOffice conversion timed out.")
        return jsonify({"error": "The file is too large or complex to convert in time."}), 500
    except Exception as e:
        app.logger.error(f"Office-to-PDF error: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        cleanup_files(input_path, pdf_filepath)

@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_wrapper():
    return office_to_pdf_tool()
    
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_wrapper():
    return office_to_pdf_tool()
    
@app.route('/powerpoint-to-pdf', methods=['POST'])
def powerpoint_to_pdf_wrapper():
    return office_to_pdf_tool()

# --- Main Execution ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)), debug=True)