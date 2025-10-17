--- START OF FILE app.py ---

# =====================================================================================
# == FINAL STABLE PRODUCTION VERSION V31.0 - Optimized Excel & Enhanced Word Logic ==
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
from docx.shared import Inches # For image size in Word

# --- NEW: Import Camelot for advanced PDF to Excel ---
try:
    import camelot
except ImportError:
    print("Warning: 'camelot-py' not installed. PDF to Excel Advanced will not work.")
    print("Please install with: pip install \"camelot-py[cv]\"")
    camelot = None


app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_safe_filepath(filename):
    safe_filename = secure_filename(filename)
    return os.path.join(UPLOAD_FOLDER, safe_filename)

# --- PDF to Word METHOD 1: High-Quality (LibreOffice) ---
# This method relies on external LibreOffice installation for best fidelity including images and formatting.
@app.route('/pdf-to-word-premium', methods=['POST'])
def pdf_to_word_premium():
    if 'file' not in request.files: return jsonify({"error": "No file part."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Invalid file type. Please upload a PDF."}), 400
    return convert_with_libreoffice(file, "docx")

# --- PDF to Word METHOD 2: Basic Fallback (In-Memory with Image Extraction) ---
# This method extracts text and images sequentially, but may not preserve complex layouts.
@app.route('/pdf-to-word-basic', methods=['POST'])
def pdf_to_word_basic():
    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF."}), 400
    try:
        pdf_bytes = file.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        word_doc = Document()

        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Extract and add text
            text = page.get_text("text")
            if text.strip():
                word_doc.add_paragraph(text)
            
            # Extract and add images
            images = page.get_images(full=True)
            for img_index, img in enumerate(images):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]

                # Save image to a temporary buffer and add to Word document
                try:
                    # Ensure we have a common image format, e.g., PNG or JPEG
                    if image_ext.lower() in ['jpeg', 'png', 'bmp', 'tiff', 'jpg']: # Added 'jpg'
                        img_stream = BytesIO(image_bytes)
                        # Adjust width for better fit, 6 inches is a common default
                        word_doc.add_picture(img_stream, width=Inches(6)) 
                        word_doc.add_paragraph(f"--- Image {img_index+1} from Page {page_num+1} ---")
                    else:
                        print(f"Skipping unsupported image format: {image_ext} on page {page_num+1}")
                        word_doc.add_paragraph(f"[Image {img_index+1} from Page {page_num+1} skipped (unsupported format: {image_ext})]")
                except Exception as img_e:
                    print(f"Error adding image {img_index+1} from page {page_num+1}: {img_e}")
                    word_doc.add_paragraph(f"[Could not insert image {img_index+1} from Page {page_num+1}]")

            # Add a page break in Word to visually separate content from different PDF pages
            if page_num < len(doc) - 1:
                word_doc.add_page_break()

        doc.close()
        doc_buffer = BytesIO()
        word_doc.save(doc_buffer)
        doc_buffer.seek(0)
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        return send_file(doc_buffer, as_attachment=True, download_name=docx_filename, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        print(f"Error during basic Word conversion: {str(e)}")
        return jsonify({"error": f"An error occurred while creating the Word file: {str(e)}"}), 500


# --- PDF to Excel (Advanced using Camelot) ---
# This new method uses Camelot for robust table extraction, preventing unintended cell merging.
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if not camelot:
        return jsonify({"error": "Camelot library not found. Please install with: pip install \"camelot-py[cv]\""}), 500

    if 'file' not in request.files: return jsonify({"error": "No file received."}), 400
    file = request.files['file']
    if not file or not file.filename.lower().endswith('.pdf'): return jsonify({"error": "Please upload a valid PDF."}), 400
    
    input_pdf_path = None
    try:
        input_pdf_path = get_safe_filepath(file.filename)
        file.save(input_pdf_path)

        # Use Camelot to extract tables
        # 'flavor': 'lattice' for PDFs with lines/grids, 'stream' for PDFs without lines
        # 'pages': 'all' to process all pages
        # You might need to experiment with 'flavor' and 'table_areas' for best results
        tables = camelot.read_pdf(input_pdf_path, flavor='lattice', pages='all') 
        
        if not tables:
            # Try 'stream' flavor as a fallback if 'lattice' finds nothing
            tables = camelot.read_pdf(input_pdf_path, flavor='stream', pages='all')
            if not tables:
                return jsonify({"error": "No tables were detected in the PDF by Camelot using both 'lattice' and 'stream' methods."}), 400
        
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            for i, table in enumerate(tables):
                # Ensure the sheet name is valid for Excel (max 31 chars)
                sheet_name = f'Table_Page_{table.page}_{i+1}'
                table.df.to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False) # header=False to match original request if desired
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '_extracted_tables.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        print(f"Error during advanced Excel conversion: {str(e)}")
        return jsonify({"error": f"An error occurred during Excel conversion: {str(e)}. Ensure Camelot and Ghostscript are installed."}), 500
    finally:
        if input_pdf_path and os.path.exists(input_pdf_path):
            os.remove(input_pdf_path)


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
    
    # Use 'soffice' for LibreOffice command. Adjust if 'libreoffice' is the command on your system.
    command = ['soffice', user_profile_arg, '--headless', '--convert-to', output_format, '--outdir', output_dir, input_path]
    
    try:
        file.save(input_path)
        # Increased timeout for potentially large files on slow servers
        result = subprocess.run(command, check=True, timeout=300, capture_output=True, text=True) # Timeout increased to 300 seconds (5 min)
        
        # LibreOffice sometimes saves files with different names (e.g., adding original extension)
        # We need to find the actual output file in the output_dir
        expected_output_prefix = os.path.splitext(os.path.basename(input_path))[0]
        actual_output_filename = None
        for f_name in os.listdir(output_dir):
            if f_name.startswith(expected_output_prefix) and f_name.endswith(f'.{output_format.split(":")[0]}'):
                actual_output_filename = f_name
                break

        if not actual_output_filename:
            print(f"LibreOffice command: {' '.join(command)}")
            print(f"LibreOffice stdout: {result.stdout}")
            print(f"LibreOffice stderr: {result.stderr}")
            raise Exception("Output file was not created by LibreOffice. Check server logs for details.")
            
        output_path = os.path.join(output_dir, actual_output_filename)
        mimetype = mimetypes.guess_type(output_path)[0] or 'application/octet-stream'
        return send_file(output_path, as_attachment=True, download_name=actual_output_filename, mimetype=mimetype)
    except subprocess.TimeoutExpired:
        return jsonify({"error": "The conversion process took too long and was timed out. The file might be too large or complex."}), 504
    except subprocess.CalledProcessError as e:
        print(f"LibreOffice conversion failed: {e.cmd}")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
        return jsonify({"error": f"LibreOffice conversion failed: {e.stderr}"}), 500
    except FileNotFoundError:
        return jsonify({"error": "LibreOffice ('soffice') command not found. Please ensure LibreOffice is installed and in your system's PATH."}), 500
    except Exception as e:
        print(f"An unexpected server error occurred: {str(e)}")
        return jsonify({"error": f"An unexpected server error occurred during conversion: {str(e)}"}), 500
    finally:
        try:
            if input_path and os.path.exists(input_path): os.remove(input_path)
            if output_path and os.path.exists(output_path): os.remove(output_path)
        except OSError as e:
            print(f"Error during file cleanup: {e}")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)