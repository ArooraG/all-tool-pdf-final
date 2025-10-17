--- START OF FILE Dockerfile ---
# Dockerfile for Python Flask application with LibreOffice, Ghostscript, and Camelot

# Use a specific Python base image (recommended for stability)
# python:3.9-slim-buster is a good choice for smaller image size
FROM python:3.9-slim-buster

# Update system packages and install external dependencies
# LibreOffice (for document conversions)
# fonts-dejavu-core (for better font rendering in LibreOffice conversions)
# ghostscript (required by Camelot for PDF processing)
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        fonts-dejavu-core \
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
ENV FLASK_ENV=production # Use 'development' for development, 'production' for deployment

# Expose the port your application will listen on
# Render automatically injects the $PORT environment variable
EXPOSE 10000

# Command to run the application using Gunicorn
# Using $PORT here so Render can inject its dynamically assigned port
CMD gunicorn -w 4 -b 0.0.0.0:$PORT app:app
--- END OF FILE Dockerfile ---