# Dockerfile for Python Flask application with LibreOffice, Ghostscript, and Camelot

# Use a specific Python base image (recommended for stability)
# python:3.9-slim-bullseye is a good choice for smaller image size with a newer Debian base
FROM python:3.9-slim-bullseye

# Update system packages and install external dependencies
# LibreOffice (for document conversions)
# fonts-dejavu-core (for better font rendering in LibreOffice conversions)
# ghostscript (required by Camelot for PDF processing)
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        fonts-dejavu-core \
        fonts-freefont-ttf \
        fonts-liberation \
        ghostscript \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set the working directory inside the container
WORKDIR /app

# Copy the requirements file and install Python dependencies
# --no-cache-dir reduces the image size
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code into the container
COPY . .

# Set environment variables for Flask
ENV FLASK_APP=app.py
# Use 'development' for development, 'production' for deployment
ENV FLASK_ENV=production

# Expose the port your application will listen on
# Render automatically injects the $PORT environment variable
EXPOSE 10000

# Command to run the application using Gunicorn
# Using $PORT here so Render can inject its dynamically assigned port
CMD gunicorn -w 4 -b 0.0.0.0:$PORT app:app