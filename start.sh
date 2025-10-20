#!/bin/bash
set -e # Is line ko shamil kar len taake script mein koi bhi command fail ho to foran ruk jaye

# start.sh - Script to start the Gunicorn server

# UPLOAD_FOLDER create karein agar maujood nahi hai
mkdir -p /app/uploads

# Gunicorn server start karein
# --timeout: Camelot jaisi libraries ke liye process ko lamba chalne ki ijazat deta hai
gunicorn --timeout 300 --workers 4 --bind 0.0.0.0:$PORT app:app