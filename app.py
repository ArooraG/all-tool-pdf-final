from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
import subprocess
import os
import uuid
import convertapi
import logging
import camelot
import pandas as pd # Pandas کو امپورٹ کریں

logging.basicConfig(level=logging.INFO)

app = Flask(__name__, template_folder='templates', static_folder='static')
CORS(app)

convertapi.api_secret = os.environ.get('CONVERTAPI_SECRET') 

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True) 
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def get_unique_filepath(file):
    ext = os.path.splitext(file.filename)[1]
    filename = str(uuid.uuid4()) + ext
    return os.path.join(app.config['UPLOAD_FOLDER'], filename)

def cleanup_files(*args):
    for file_path in args:
        if file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                app.logger.error(f"Error deleting file {file_path}: {e}")

# --- HTML پیجز کے روٹس ---
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/blog/blogs.html')
def blog_page():
    return render_template('blog/blogs.html')

# --- بیک اینڈ ٹولز کے روٹس ---
def run_libreoffice_conversion(input_path, output_format, output_dir):
    command = ['libreoffice', '--headless', '--convert-to', output_format, '--outdir', output_dir, input_path]
    app.logger.info(f"Running command: {' '.join(command)}")
    process = subprocess.run(command, check=True, timeout=300, capture_output=True, text=True)
    app.logger.info(f"LibreOffice stdout: {process.stdout}")
    if process.stderr:
        app.logger.error(f"LibreOffice stderr: {process.stderr}")

@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf_tool():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400
    
    input_path, pdf_filepath = None, None
    try:
        input_path = get_unique_filepath(file)
        file.save(input_path)
        output_dir = app.config['UPLOAD_FOLDER']
        
        run_libreoffice_conversion(input_path, 'pdf', output_dir)
        
        pdf_filename = os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
        pdf_filepath = os.path.join(output_dir, pdf_filename)
        
        if not os.path.exists(pdf_filepath): raise FileNotFoundError("Conversion to PDF failed.")
        
        return send_file(pdf_filepath, as_attachment=True)
    except Exception as e:
        app.logger.error(f"Word-to-PDF error: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        cleanup_files(input_path, pdf_filepath)

@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word_tool():
    if not convertapi.api_secret:
        app.logger.error("ConvertAPI Secret is not configured on the server.")
        return jsonify({"error": "API Conversion failed: ConvertAPI is not configured on the server."}), 500

    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400
    
    input_path, docx_filepath = None, None
    try:
        input_path = get_unique_filepath(file)
        file.save(input_path)
        
        result = convertapi.convert('docx', {'File': input_path})
        docx_filepath = result.file.save(app.config['UPLOAD_FOLDER'])
        
        return send_file(docx_filepath, as_attachment=True)
    except Exception as e:
        app.logger.error(f"PDF-to-Word (API) error: {e}")
        return jsonify({"error": f"API Conversion failed: {str(e)}"}), 500
    finally:
        cleanup_files(input_path, docx_filepath)

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_tool():
    return word_to_pdf_tool()

# --- PDF to Excel کا نیا اور بہتر کوڈ ---
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel_tool():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400
    
    input_path, xlsx_filepath = None, None
    try:
        input_path = get_unique_filepath(file)
        file.save(input_path)
        
        # Camelot کے ذریعے تمام صفحات سے ٹیبلز نکالیں
        tables = camelot.read_pdf(input_path, pages='all', flavor='stream')
        
        if tables.n == 0:
            return jsonify({"error": "No tables found in this PDF file."}), 400

        # تمام ٹیبلز کو ایک لسٹ میں جمع کریں
        all_dfs = [table.df for table in tables]
        
        # Pandas کا استعمال کرکے تمام ٹیبلز کو ایک DataFrame میں جوڑ دیں
        combined_df = pd.concat(all_dfs, ignore_index=True)

        # ایکسل فائل کا نام تیار کریں
        xlsx_filename = os.path.splitext(os.path.basename(input_path))[0] + '.xlsx'
        xlsx_filepath = os.path.join(app.config['UPLOAD_FOLDER'], xlsx_filename)
        
        # ایکسل فائل بنائیں جس میں صرف ایک شیٹ ہو اور انڈیکس نمبر نہ ہوں
        combined_df.to_excel(xlsx_filepath, index=False, header=False, sheet_name='All_Pages_Data')
        
        if not os.path.exists(xlsx_filepath):
            raise FileNotFoundError("PDF to Excel conversion failed.")

        return send_file(xlsx_filepath, as_attachment=True)
    except Exception as e:
        app.logger.error(f"PDF-to-Excel error: {e}")
        return jsonify({"error": f"Something went wrong during PDF to Excel conversion: {e}"}), 500
    finally:
        cleanup_files(input_path, xlsx_filepath)
        
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))