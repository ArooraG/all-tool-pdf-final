from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
import subprocess
import os
import uuid
import logging

logging.basicConfig(level=logging.INFO)
app = Flask(__name__, template_folder='templates', static_folder='static')
CORS(app)

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

# --- HTML Pages ke Routes ---
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/blog/blogs.html')
def blog_page():
    return render_template('blog/blogs.html')

# --- Backend Conversion Tools ---
# Sirf halkay tools ab yahan hain (jaise Word to PDF)
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
        
        command = ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', output_dir, input_path]
        subprocess.run(command, check=True, timeout=300)
        
        pdf_filename = os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
        pdf_filepath = os.path.join(output_dir, pdf_filename)
        
        if not os.path.exists(pdf_filepath): raise FileNotFoundError("Conversion to PDF failed.")
        
        return send_file(pdf_filepath, as_attachment=True)
    except Exception as e:
        app.logger.error(f"Word-to-PDF error: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        cleanup_files(input_path, pdf_filepath)

@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf_tool():
    # Yeh tool bhi Word to PDF ki tarah LibreOffice se chalega
    return word_to_pdf_tool()
    
@app.route('/powerpoint-to-pdf', methods=['POST'])
def powerpoint_to_pdf_tool():
    # Yeh tool bhi Word to PDF ki tarah LibreOffice se chalega
    return word_to_pdf_tool()

if __name__ == '__main__':
    app.run(host='0.0.0.O', port=int(os.environ.get("PORT", 10000)))