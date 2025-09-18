from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import requests
import pandas as pd
from docx import Document
import pdfplumber
from io import BytesIO
import subprocess

app = Flask(__name__)
CORS(app)  # Isse frontend aur backend connect ho paenge

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Helper function to get a safe filepath
def get_safe_filepath(filename):
    return os.path.join(UPLOAD_FOLDER, filename)

# 1. Word to PDF Tool
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and file.filename.endswith(('.doc', '.docx')):
        try:
            input_path = get_safe_filepath(file.filename)
            file.save(input_path)
            
            # LibreOffice command to convert Word to PDF
            # Yeh Render.com par kaam karega agar aap wahan LibreOffice install karwate hain
            # Iske liye aapko ek `render.yaml` file use karni hogi jisme build script ho
            subprocess.run(
                ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path],
                check=True
            )
            
            pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
            pdf_path = get_safe_filepath(pdf_filename)
            
            if not os.path.exists(pdf_path):
                 return jsonify({"error": "Conversion failed, output file not found."}), 500

            response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
            
            # Clean up files
            os.remove(input_path)
            os.remove(pdf_path)
            
            return response
        except Exception as e:
            # Error cleanup
            if 'input_path' in locals() and os.path.exists(input_path):
                os.remove(input_path)
            return jsonify({"error": f"An error occurred: {str(e)}"}), 500

    return jsonify({"error": "Invalid file type, please upload a .doc or .docx file."}), 400

# 2. Excel to PDF Tool
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    # Word to PDF jaisa hi logic yahan bhi use hoga LibreOffice ke saath
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and file.filename.endswith(('.xls', '.xlsx')):
        try:
            input_path = get_safe_filepath(file.filename)
            file.save(input_path)
            
            subprocess.run(
                ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', UPLOAD_FOLDER, input_path],
                check=True
            )
            
            pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
            pdf_path = get_safe_filepath(pdf_filename)

            if not os.path.exists(pdf_path):
                 return jsonify({"error": "Conversion failed, output file not found."}), 500

            response = send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
            os.remove(input_path)
            os.remove(pdf_path)
            return response
        except Exception as e:
            if 'input_path' in locals() and os.path.exists(input_path):
                os.remove(input_path)
            return jsonify({"error": f"An error occurred: {str(e)}"}), 500

    return jsonify({"error": "Invalid file type, please upload an .xls or .xlsx file."}), 400

# 3. PDF to Excel Tool
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file"}), 400

    try:
        input_path = get_safe_filepath(file.filename)
        file.save(input_path)
        
        # pdfplumber se table extract karein
        all_tables = []
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    all_tables.append(pd.DataFrame(table[1:], columns=table[0]))

        if not all_tables:
            return jsonify({"error": "No tables found in the PDF"}), 400

        # Excel file banayein
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            for i, df in enumerate(all_tables):
                df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
        
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        
        os.remove(input_path) # Clean up uploaded file
        
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, 
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        if 'input_path' in locals() and os.path.exists(input_path):
            os.remove(input_path)
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


# 4. PDF to Word (API ke zariye)
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    # Bhai, aapke diye gaye tokens kis service ke hain, yeh pehchaan nahi ho saki.
    # Main aapko ConvertAPI ka example de raha hoon, jo ek popular service hai.
    # Aap apni service ki documentation dekh kar URL aur headers badal sakte hain.
    
    API_SECRET = 'YOUR_CONVERTAPI_SECRET' # <-- Apna API Secret yahan daalein
    API_URL = f'https://v2.convertapi.com/convert/pdf/to/docx?Secret={API_SECRET}'
    
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '' or not file.filename.endswith('.pdf'):
        return jsonify({"error": "Please upload a valid PDF file"}), 400

    try:
        # File ko direct API par bhej dein
        response = requests.post(API_URL, files={'file': file.read()})
        response.raise_for_status() # Agar koi error ho toh ruk jaye
        
        converted_file_data = BytesIO(response.content)
        
        docx_filename = os.path.splitext(file.filename)[0] + '.docx'
        
        return send_file(converted_file_data, as_attachment=True, download_name=docx_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"API request failed: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)