# Base image for your Python application
# Aap isko badal sakte hain agar aap koi aur base image use kar rahe hain (e.g., python:3.9-slim-buster)
FROM python:3.9-slim-buster

# Ya for example, agar aap python:3.9 use kar rahe hain
# FROM python:3.9

# System dependencies ko update aur install karein
# LibreOffice aur Ghostscript yahan install honge
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        fonts-dejavu-core \
        ghostscript \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Working directory set karein
WORKDIR /app

# requirements.txt ko copy karein aur Python dependencies install karein
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Baaqi application files copy karein
COPY . .

# Environment variable for Flask
ENV FLASK_APP=app.py
ENV FLASK_ENV=production # Ya development, jo aapka preference ho

# Port expose karein
EXPOSE 10000

# Application ko run karne ki command
CMD gunicorn -w 4 -b 0.0.0.0:10000 app:app
# Ya jo bhi aapki actual run command ho (e.g., flask run)