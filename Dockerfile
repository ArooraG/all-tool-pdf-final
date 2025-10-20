# Base image for Python applications (Python 3.10 on Debian Bookworm)
FROM python:3.10-slim-bookworm

# Set the working directory inside the container
WORKDIR /app

# Install system dependencies required by Camelot and Ghostscript
# Ghostscript is crucial for Camelot to process PDFs
# libgl1-mesa-glx for headless OpenCV (part of camelot-py[cv])
# gcc and python3-dev for compiling some Python packages
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    ghostscript \
    libgl1-mesa-glx \
    gcc \
    python3-dev \
    # Clean up APT cache to reduce image size
    && rm -rf /var/lib/apt/lists/*

# Copy the requirements file into the container (Explicit Path)
COPY requirements.txt /app/requirements.txt

# Install Python dependencies
RUN pip install --no-cache-dir -r /app/requirements.txt

# Create a directory for uploaded files
RUN mkdir -p /app/uploads

# Copy your application code into the container (Explicit Paths)
COPY app.py /app/app.py
COPY start.sh /app/start.sh

# Make the start script executable (Explicit Path)
RUN chmod +x /app/start.sh

# Expose the port your Flask app will run on
EXPOSE 10000

# Define the command to run your Flask application using gunicorn (Explicit Path)
CMD ["/app/start.sh"]