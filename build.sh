#!/usr/bin/env bash
# exit on error
set -o errexit

# Step 1: Install Python dependencies
echo "Installing Python dependencies..."
pip install -r requirements.txt

# Step 2: Install LibreOffice with ALL necessary components
# Hum 'pdfimport' library aur 'fonts-noto' (jo har zuban ko support karti hai) install kar rahe hain
echo "Updating packages and installing LibreOffice with all dependencies..."
apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-pdfimport \
    fonts-noto \
    && apt-get clean

# Step 3: Set a HOME environment variable. YEH SAB SE ZAROORI HAI.
# LibreOffice ko headless mode mein chalne ke liye ek writable HOME directory ki zaroorat hoti hai.
# Iske baghair, woh complex files par aksar silent fail ho jata hai.
export HOME=/tmp

echo "Build script completed successfully. Environment is ready."