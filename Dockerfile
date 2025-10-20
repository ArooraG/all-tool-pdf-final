# Base image for Python applications (Python 3.10 on Debian Bookworm)
FROM python:3.10-slim-bookworm

# Set the working directory inside the container
WORKDIR /app

# Install system dependencies required by Camelot and Ghostscript
# Ghostscript is crucial for Camelot to process PDFs
# libpq-dev is for PostgreSQL, include if you plan to use a DB
# gcc and python3-dev are for compiling some Python packages (e.g., pandas, numpy)
# libgl1 is for headless OpenCV if camelot-py[cv] pulls it in that way
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    ghostscript \
    libgl1-mesa-glx \
    gcc \
    python3-dev \
    # Clean up APT cache to reduce image size
    && rm -rf /var/lib/apt/lists/*

# Copy the requirements file into the container
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Create a directory for uploaded files (if needed, though Camelot uses temp files)
RUN mkdir -p /app/uploads

# Copy your application code into the container
COPY app.py .
COPY start.sh . # Copy the start script

# Make the start script executable
RUN chmod +x start.sh

# Expose the port your Flask app will run on
# Render will map this to an external port
EXPOSE 10000

# Define the command to run your Flask application using gunicorn
# Render will use the command specified in its service settings or this CMD
CMD ["./start.sh"]