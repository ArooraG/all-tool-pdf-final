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

# Explicitly enable Java in LibreOffice
# This often helps LibreOffice detect and use the installed JRE
# The sleep commands give LibreOffice time to start and process the command
RUN libreoffice --headless --nologo --nofirststartwizard --norestore --accept='socket,host=localhost,port=2002;urp;StarOffice.ServiceManager' & \
    LO_PID=$! && sleep 10 && \
    libreoffice --headless "vnd.sun.star.script:ScriptForge.SF_Core.Basic.java_setup?language=Basic&location=application" & \
    sleep 5 && kill $LO_PID

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