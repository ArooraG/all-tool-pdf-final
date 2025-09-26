# ==========================================================
# == DIAGNOSTIC TEST CODE V1 - Sirf LibreOffice ko test karein ==
# ==========================================================
from flask import Flask, jsonify
import subprocess

app = Flask(__name__)

@app.route('/')
def home():
    return "Test server is running. Go to /test-libreoffice to check."

@app.route('/test-libreoffice')
def test_libreoffice():
    try:
        # Yeh command sirf LibreOffice ka version check karegi.
        # Agar yeh chal gayi, to iska matlab hai ke installation theek hai.
        result = subprocess.run(
            ['soffice', '--version'],
            check=True, timeout=30, capture_output=True, text=True
        )
        return jsonify({
            "status": "Success",
            "message": "LibreOffice is installed and responding.",
            "version_info": result.stdout.strip()
        })
    except FileNotFoundError:
        return jsonify({
            "status": "Error",
            "message": "LibreOffice command 'soffice' not found. Installation likely failed. Check your build.sh script."
        }), 500
    except Exception as e:
        return jsonify({
            "status": "Error",
            "message": "An unexpected error occurred while trying to run LibreOffice.",
            "error_details": str(e)
        }), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)