#!/usr/bin/env bash
# exit on error
set -o errexit

# Install Python dependencies first
pip install -r requirements.txt

# --- UPGRADED LIBREOFFICE INSTALLATION ---
# Update package list and install the full version of LibreOffice along with extra fonts.
# 'libreoffice-l10n-en-gb' helps with language packs.
# 'fonts-takao-gothic' provides essential Japanese fonts.
# 'fonts-noto' is a massive font collection from Google covering almost all languages.
# This makes our converter much more powerful for international documents.
apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-l10n-en-gb \
    fonts-takao-gothic \
    fonts-noto \
    && apt-get clean