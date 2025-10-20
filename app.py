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

            # --- Improved Line Grouping ---
            line_groups = defaultdict(list)
            
            # Calculate a dynamic line threshold based on median word height on the page
            word_heights = [w[3] - w[1] for w in words if (w[3] - w[1]) > 0] # y1 - y0
            avg_word_height = np.median(word_heights) if word_heights else 10 # Default median height
            
            # Allow some vertical wiggle room, e.g., 50% of average word height. Min 3px.
            line_clustering_threshold = max(3, avg_word_height * 0.5)

            for w in words:
                found_group = False
                # Use y0 (top-left y) for line grouping
                for y_group_key in list(line_groups.keys()):
                    if abs(w[1] - y_group_key) < line_clustering_threshold:
                        line_groups[y_group_key].append(w)
                        found_group = True
                        break
                if not found_group:
                    line_groups[w[1]].append(w) # Start a new group with its y0

            # Consolidate line groups by averaging Y0 for keys and sorting
            # This handles cases where words on the same visual line have slightly different y0 values
            consolidated_lines = []
            for y_key in sorted(line_groups.keys()):
                words_in_current_group = line_groups[y_key]
                if words_in_current_group: # Ensure group is not empty
                    avg_y0_of_group = sum([w[1] for w in words_in_current_group]) / len(words_in_current_group)
                    consolidated_lines.append((avg_y0_of_group, words_in_current_group))
            
            # Sort all consolidated lines by their average y0
            consolidated_lines.sort(key=lambda x: x[0])
            
            # Final merge of very close lines (using the same threshold)
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

            # --- Improved Column Detection ---
            all_x0_coords = sorted([w[0] for line_words in final_lines for w in line_words])
            
            # If no words, or not enough words to form columns, handle as a single column (or skip page)
            if not all_x0_coords:
                if len(final_lines) > 0:
                     single_column_data = [[" ".join([w[4] for w in line_words])] for line_words in final_lines]
                     all_pages_data.extend(single_column_data)
                continue

            # Cluster x0 coordinates to find potential column start boundaries
            column_x_candidates = []
            # Dynamic tolerance for x-coordinates to be considered "aligned"
            # Using avg_word_height * 0.75 for x-alignment, minimum 5px
            dynamic_x_alignment_tolerance = max(5, avg_word_height * 0.75) 
            
            if all_x0_coords:
                current_cluster_sum = all_x0_coords[0]
                current_cluster_count = 1
                for i in range(1, len(all_x0_coords)):
                    if all_x0_coords[i] - all_x0_coords[i-1] < dynamic_x_alignment_tolerance:
                        current_cluster_sum += all_x0_coords[i]
                        current_cluster_count += 1
                    else:
                        column_x_candidates.append(current_cluster_sum / current_cluster_count)
                        current_cluster_sum = all_x0_coords[i]
                        current_cluster_count = 1
                column_x_candidates.append(current_cluster_sum / current_cluster_count) # Add the last cluster

            # Include page boundaries and sorted clustered candidates
            initial_column_boundaries = {page.rect.x0, page.rect.x1}
            for x_cand in column_x_candidates:
                # Add candidate if it's within the page and not too close to existing boundaries
                if page.rect.x0 < x_cand < page.rect.x1:
                    initial_column_boundaries.add(x_cand)

            final_column_boundaries = sorted(list(initial_column_boundaries))

            # Refine boundaries: remove very close ones to avoid super thin columns
            refined_column_boundaries = []
            # Dynamic minimum pixel width for a column to be considered distinct
            # Using avg_word_height * 1.0 for min column width, minimum 10px
            dynamic_min_column_width = max(10, avg_word_height * 1.0) 

            if final_column_boundaries:
                refined_column_boundaries.append(final_column_boundaries[0])
                for i in range(1, len(final_column_boundaries)):
                    if final_column_boundaries[i] - refined_column_boundaries[-1] > dynamic_min_column_width:
                        refined_column_boundaries.append(final_column_boundaries[i])
            
            # If after refinement, we still have less than 2 boundaries, fall back to single column
            if len(refined_column_boundaries) < 2:
                if len(final_lines) > 0:
                     single_column_data = [[" ".join([w[4] for w in line_words])] for line_words in final_lines]
                     all_pages_data.extend(single_column_data)
                continue
            
            # Use refined_column_boundaries from now on
            final_column_boundaries = refined_column_boundaries
            # --- End Improved Column Detection ---

            # Now process each line (row) using the determined column boundaries
            for line_words in final_lines:
                if not line_words: continue
                
                # Initialize row with empty strings for each potential cell
                row = [""] * (len(final_column_boundaries) - 1)
                
                # Sort words within the line by x0 to ensure correct reading order
                line_words.sort(key=lambda w: w[0])

                for w in line_words:
                    word_x0 = w[0]
                    word_x1 = w[2]
                    word_text = w[4]

                    best_col_idx = -1
                    max_overlap_ratio = 0.0
                    
                    # Iterate through potential columns to find the best fit for the word
                    for col_idx in range(len(final_column_boundaries) - 1):
                        col_left_bound = final_column_boundaries[col_idx]
                        col_right_bound = final_column_boundaries[col_idx + 1]

                        # Calculate overlap between word and column
                        overlap_start = max(word_x0, col_left_bound)
                        overlap_end = min(word_x1, col_right_bound)
                        current_overlap_width = max(0, overlap_end - overlap_start)

                        word_width = w[2] - w[0]
                        if word_width > 0:
                            overlap_ratio = current_overlap_width / word_width
                            
                            # Heuristic: Assign word if it largely overlaps (e.g., > 60%) with a column,
                            # or if its center is within the column and there's some overlap.
                            word_center_x = (word_x0 + word_x1) / 2
                            is_center_in_col = col_left_bound <= word_center_x <= col_right_bound

                            if overlap_ratio > 0.6 or (is_center_in_col and current_overlap_width > 0):
                                if overlap_ratio > max_overlap_ratio:
                                    max_overlap_ratio = overlap_ratio
                                    best_col_idx = col_idx
                            # Edge case: If word barely overlaps, but no other strong fit.
                            # Try to assign to the current column if no better fit found.
                            elif current_overlap_width > 0 and best_col_idx == -1: # if we have some overlap but no strong fit yet
                                best_col_idx = col_idx # temporarily assign
                    
                    # Assign the word to the determined best column
                    if best_col_idx != -1:
                        if row[best_col_idx]:
                            # Add a space before concatenating if the cell already has content
                            row[best_col_idx] += " " + word_text
                        else:
                            row[best_col_idx] = word_text
                    # Fallback for words that genuinely didn't fit strongly into any defined column
                    # Try to assign to the closest column if no strong overlap, to minimize skipped data
                    else:
                        min_dist_to_col = float('inf')
                        fallback_col_idx = -1
                        word_center_x = (word_x0 + word_x1) / 2
                        for col_idx in range(len(final_column_boundaries) - 1):
                            col_left_bound = final_column_boundaries[col_idx]
                            col_right_bound = final_column_boundaries[col_idx + 1]
                            
                            # Check if the word is entirely before or after the column
                            if word_x1 < col_left_bound: # Word is to the left of the column
                                dist = col_left_bound - word_x1
                            elif word_x0 > col_right_bound: # Word is to the right of the column
                                dist = word_x0 - col_right_bound
                            else: # Word is within or overlapping the column (already handled by main logic)
                                dist = 0 

                            if dist < min_dist_to_col:
                                min_dist_to_col = dist
                                fallback_col_idx = col_idx
                        
                        if fallback_col_idx != -1:
                             if row[fallback_col_idx]:
                                 row[fallback_col_idx] += " " + word_text
                             else:
                                 row[fallback_col_idx] = word_text

                all_pages_data.append(row)

        doc.close()
        
        if not all_pages_data: return jsonify({"error": "No text data or tables were extracted from the PDF. For complex tables, manual conversion is recommended."}), 400
        
        # Ensure all rows have the same number of columns for DataFrame creation
        # Find max_cols AFTER all_pages_data has been populated from all pages
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
        # For debugging, you can print the error
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