#!/usr/bin/env bash
# exit on error
set -o errexit

# Install Python dependencies first
echo "Installing Python dependencies..."
pip install -r requirements.txt

# --- Final LibreOffice Installation for Render.com ---
echo "Updating packages and installing LibreOffice with all dependencies..."
apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-pdfimport \
    fonts-noto \
    && apt-get clean

# **THE CRITICAL FIX IS HERE**
# Create a dedicated, writable directory for the LibreOffice user profile.
# On Render's read-only filesystem, LibreOffice cannot create its config files.
# This command creates a directory it can actually write to.
echo "Creating a writable directory for LibreOffice user profile..."
mkdir -p /opt/render/.config/libreoffice

echo "Build script completed successfully. Environment is ready."