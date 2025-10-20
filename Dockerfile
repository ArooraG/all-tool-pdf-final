# Dockerfile for Python Flask application with LibreOffice, Ghostscript, and Camelot

# Use a specific Python base image (recommended for stability)
# Using Python 3.10 for newer environment
FROM python:3.10-slim-bullseye

# Update system packages and install external dependencies
# LibreOffice (for document conversions)
# fonts-dejavu-core (for better font rendering in LibreOffice conversions)
# ghostscript (required by Camelot for PDF processing)
# default-jre and libreoffice-java-common are added to resolve Java Runtime Environment issues with LibreOffice
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        fonts-dejavu-core \
        fonts-freefont-ttf \
        fonts-liberation \
        ghostscript \
        default-jre \
        libreoffice-java-common \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set the working directory inside the container
WORKDIR /app

# Configure LibreOffice to use more generic rendering (less reliance on specific display server features)
# And set up some environment variables for memory/logging
ENV SAL_USE_VCLPLUGIN=gen
ENV UNO_VERBOSE=true
ENV URE_BOOTSTRAP_LINES=20

# We removed the explicit Java setup from previous iteration as installing default-jre + libreoffice-java-common
# should ideally be enough for LibreOffice to find Java. If issues persist, logs will tell us more.


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