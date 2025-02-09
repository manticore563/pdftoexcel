import os
from flask import Flask, render_template, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from utils.pdf_processor import process_pdf_to_excel
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes
app.secret_key = "pdf2excel_secret_key"

# Configure upload settings
UPLOAD_FOLDER = 'temp'
ALLOWED_EXTENSIONS = {'pdf'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Create temp directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST', 'OPTIONS'])
def upload_file():
    # Add CORS headers to the response
    if request.method == 'OPTIONS':
        response = app.make_default_options_response()
        response.headers['Access-Control-Allow-Methods'] = 'POST'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return response

    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only PDF files are allowed'}), 400

    pdf_path = None
    excel_path = None

    try:
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 
                                f"{os.path.splitext(filename)[0]}.xlsx")

        # Save the uploaded PDF
        file.save(pdf_path)
        logging.debug(f"PDF saved at: {pdf_path}")

        # Process PDF to Excel
        process_pdf_to_excel(pdf_path, excel_path)
        logging.debug(f"Excel file created at: {excel_path}")

        # Return the Excel file with proper headers
        response = send_file(
            excel_path,
            as_attachment=True,
            download_name=f"{os.path.splitext(filename)[0]}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'
        return response

    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        return jsonify({'error': str(e)}), 500

    finally:
        # Cleanup temporary files
        try:
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
                logging.debug(f"Cleaned up PDF file: {pdf_path}")
            if excel_path and os.path.exists(excel_path):
                os.remove(excel_path)
                logging.debug(f"Cleaned up Excel file: {excel_path}")
        except Exception as e:
            logging.error(f"Error cleaning up temporary files: {str(e)}")

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File is too large. Maximum size is 16MB'}), 413

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)