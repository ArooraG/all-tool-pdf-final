 #!/usr/bin/env bash
    # exit on error
    set -o errexit

    # Install dependencies from requirements.txt
    pip install -r requirements.txt

    # Install LibreOffice for Word/Excel to PDF conversion
    apt-get update && apt-get install -y libreoffice && apt-get clean
    