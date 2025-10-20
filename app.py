# =====================================================================================
# == FINAL STABLE PRODUCTION VERSION V30.0 - Optimized Excel Logic (Table Structure Fix) ==
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
import numpy as np # Added for numerical operations in Excel conversion
from collections import defaultdict # Added for easier grouping in Excel conversion

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
    # LibreOffice generally handles text and images well.
    # Ensure LibreOffice is installed and configured correctly on the server.
    # For very complex PDFs, conversion quality may vary.
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
            # This basic method extracts text only, it does not handle images or complex layouts.
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

        for page_num, page in enumerate(doc):
            words = page.get_text("words") # [x0, y0, x1, y1, word, block_no, line_no, word_no]
            if not words: continue

            # --- Improved Line Grouping (Rows) ---
            line_groups = defaultdict(list)
            word_heights = [w[3] - w[1] for w in words if (w[3] - w[1]) > 0]
            avg_word_height = np.median(word_heights) if word_heights else 10 # Default median height
            
            line_clustering_threshold = max(3, avg_word_height * 0.5)

            for w in words:
                found_group = False
                for y_group_key in list(line_groups.keys()):
                    if abs(w[1] - y_group_key) < line_clustering_threshold:
                        line_groups[y_group_key].append(w)
                        found_group = True
                        break
                if not found_group:
                    line_groups[w[1]].append(w)

            consolidated_lines = []
            for y_key in sorted(line_groups.keys()):
                words_in_current_group = line_groups[y_key]
                if words_in_current_group:
                    avg_y0_of_group = sum([w[1] for w in words_in_current_group]) / len(words_in_current_group)
                    consolidated_lines.append((avg_y0_of_group, words_in_current_group))
            
            consolidated_lines.sort(key=lambda x: x[0])
            
            final_lines = []
            if consolidated_lines:
                current_line_words = consolidated_lines[0][1]
                last_avg_y0 = consolidated_lines[0][0]
                for avg_y0, words_in_line in consolidated_lines[1:]:
                    if abs(avg_y0 - last_avg_y0) < line_clustering_threshold:
                        current_line_words.extend(words_in_line)
                    else:
                        final_lines.append(sorted(current_line_words, key=lambda w: w[0]))
                        current_line_words = words_in_line
                    last_avg_y0 = avg_y0
                final_lines.append(sorted(current_line_words, key=lambda w: w[0]))
            # --- End Improved Line Grouping ---

            # --- Refined Column Detection ---
            x_coords_for_columns = set()
            for line_words in final_lines:
                line_words.sort(key=lambda w: w[0])
                for i, w in enumerate(line_words):
                    x_coords_for_columns.add(w[0]) # Start of word
                    x_coords_for_columns.add(w[2]) # End of word
                    if i < len(line_words) - 1:
                        gap = line_words[i+1][0] - w[2]
                        if gap > avg_word_height * 0.8: # Consider a significant gap as a potential column divider
                             x_coords_for_columns.add(w[2] + gap / 2) # Mid-point of the gap
            
            blocks = page.get_text("blocks")
            for b in blocks:
                 x_coords_for_columns.add(b[0])
                 x_coords_for_columns.add(b[2])

            x_coords_for_columns.add(page.rect.x0)
            x_coords_for_columns.add(page.rect.x1)

            sorted_x_candidates = sorted(list(x_coords_for_columns))
            
            column_boundary_clustering_threshold = max(5, avg_word_height * 0.7) # Cluster nearby x-coordinates into a single boundary

            final_column_boundaries = []
            if sorted_x_candidates:
                current_cluster_sum = sorted_x_candidates[0]
                current_cluster_count = 1
                for i in range(1, len(sorted_x_candidates)):
                    if sorted_x_candidates[i] - sorted_x_candidates[i-1] < column_boundary_clustering_threshold:
                        current_cluster_sum += sorted_x_candidates[i]
                        current_cluster_count += 1
                    else:
                        final_column_boundaries.append(current_cluster_sum / current_cluster_count)
                        current_cluster_sum = sorted_x_candidates[i]
                        current_cluster_count = 1
                final_column_boundaries.append(current_cluster_sum / current_cluster_count)

            final_column_boundaries = sorted(list(set(final_column_boundaries)))

            min_effective_column_width = max(10, avg_word_height * 1.5) # Minimum width for a valid column

            refined_column_boundaries = []
            if final_column_boundaries:
                refined_column_boundaries.append(final_column_boundaries[0])
                for i in range(1, len(final_column_boundaries)):
                    if final_column_boundaries[i] - refined_column_boundaries[-1] > min_effective_column_width:
                        refined_column_boundaries.append(final_column_boundaries[i])
            
            if len(refined_column_boundaries) < 2:
                if len(final_lines) > 0:
                     single_column_data = [[" ".join([w[4] for w in line_words])] for line_words in final_lines]
                     all_pages_data.extend(single_column_data)
                continue
            
            final_column_boundaries = refined_column_boundaries
            # --- End Refined Column Detection ---

            # Now process each line (row) using the determined column boundaries
            for line_words in final_lines:
                if not line_words: continue
                
                row = [""] * (len(final_column_boundaries) - 1)
                
                line_words.sort(key=lambda w: w[0])

                words_processed_in_line = [False] * len(line_words)

                # Assign words to the most appropriate column
                for col_idx in range(len(final_column_boundaries) - 1):
                    col_left_bound = final_column_boundaries[col_idx]
                    col_right_bound = final_column_boundaries[col_idx + 1]
                    
                    words_in_current_cell = []
                    for i, w in enumerate(line_words):
                        if words_processed_in_line[i]: continue

                        word_x0 = w[0]
                        word_x1 = w[2]
                        word_center_x = (word_x0 + word_x1) / 2

                        is_center_in_col = col_left_bound <= word_center_x <= col_right_bound
                        overlap_with_col = max(0, min(word_x1, col_right_bound) - max(word_x0, col_left_bound))
                        word_width = w[2] - w[0]
                        overlap_ratio = overlap_with_col / word_width if word_width > 0 else 0

                        # Strong criteria for primary assignment to a column
                        if is_center_in_col or overlap_ratio > 0.7:
                            words_in_current_cell.append(w)
                            words_processed_in_line[i] = True

                    if words_in_current_cell:
                        words_in_current_cell.sort(key=lambda x: x[0])
                        row[col_idx] = " ".join([w[4] for w in words_in_current_cell])

                # Fallback for any remaining unprocessed words: assign to the closest column
                for i, w in enumerate(line_words):
                    if not words_processed_in_line[i]:
                        word_center_x = (w[0] + w[2]) / 2

                        min_dist = float('inf')
                        closest_col_idx = -1

                        for col_idx in range(len(final_column_boundaries) - 1):
                            col_left_bound = final_column_boundaries[col_idx]
                            col_right_bound = final_column_boundaries[col_idx + 1]
                            
                            col_center = (col_left_bound + col_right_bound) / 2
                            dist = abs(word_center_x - col_center)

                            if dist < min_dist:
                                min_dist = dist
                                closest_col_idx = col_idx
                        
                        if closest_col_idx != -1:
                            if row[closest_col_idx]:
                                row[closest_col_idx] += " " + w[4]
                            else:
                                row[closest_col_idx] = w[4]
                        words_processed_in_line[i] = True

                all_pages_data.append(row)

        doc.close()
        
        if not all_pages_data: return jsonify({"error": "No text data or tables were extracted from the PDF. For complex tables, manual conversion is recommended."}), 400
        
        max_cols = max(len(row) for row in all_pages_data) if all_pages_data else 0
        normalized_data = [row + [''] * (max_cols - len(row)) for row in all_pages_data]

        df = pd.DataFrame(normalized_data)
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='All_Pages_Data', index=False, header=False)
        output_buffer.seek(0)
        
        excel_filename = os.path.splitext(file.filename)[0] + '.xlsx'
        return send_file(output_buffer, as_attachment=True, download_name=excel_filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.document')
    except Exception as e:
        print(f"Error during Excel conversion: {str(e)}")
        return jsonify({"error": f"An error occurred during Excel conversion: {str(e)}. For complex tables, manual conversion is recommended."}), 500


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
    user_profile_dir = os.path.abspath(os.path.join(UPLOAD_FOLDER, 'libreoffice_profile'))
    if not os.path.exists(user_profile_dir):
        os.makedirs(user_profile_dir)
    
    user_profile_arg = f"-env:UserInstallation=file://{user_profile_dir}"
    
    command = ['soffice', user_profile_arg, '--headless', '--convert-to', output_format, '--outdir', output_dir, input_path]
    try:
        file.save(input_path)
        result = subprocess.run(command, check=True, timeout=180, capture_output=True, text=True)
        
        output_filename = os.path.splitext(os.path.basename(input_path))[0] + f'.{output_format.split(":")[0]}'
        output_path = get_safe_filepath(output_filename)
        
        if not os.path.exists(output_path):
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