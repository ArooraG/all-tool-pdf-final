#!/bin/bash
# start.sh - Script to start the Gunicorn server

# UPLOAD_FOLDER create karein agar maujood nahi hai
mkdir -p /app/uploads

# Gunicorn server start karein
# -w: workers ki tadad (2-4 * CPU cores recommended)
# -b: bind address (0.0.0.0:port)
# app:app: aapki app.py file aur uske andar 'app' variable ka naam
# --timeout: Camelot jaisi libraries ke liye process ko lamba chalne ki ijazat deta hai
gunicorn --timeout 300 --workers 4 --bind 0.0.0.0:$PORT app:app