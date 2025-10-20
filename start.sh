#!/bin/bash
set -e

mkdir -p /app/uploads

gunicorn --timeout 300 --workers 1 --bind 0.0.0.0:$PORT app:app