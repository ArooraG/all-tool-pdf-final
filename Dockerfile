# Dockerfile for Python Flask application with LibreOffice, Ghostscript, and Python 3.10 on Ubuntu
# Using a more robust base image (Ubuntu) for better LibreOffice compatibility

# Use Ubuntu 22.04 as the base image for broader compatibility with LibreOffice
FROM ubuntu:22.04

# Set environment variables for non-interactive apt-get installs and Python
ENV DEBIAN_FRONTEND=noninteractive
ENV PYENV_ROOT="/opt/pyenv"
ENV PATH="$PYENV_ROOT/bin:$PYENV_ROOT/shims:$PATH"

# Update system packages and install prerequisites for pyenv, Python, LibreOffice, Ghostscript
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        build-essential \
        curl \
        git \
        libssl-dev \
        zlib1g-dev \
        libbz2-dev \
        libreadline-dev \
        libsqlite3-dev \
        wget \
        llvm \
        libncursesw5-dev \
        xz-utils \
        tk-dev \
        libxml2-dev \
        libxmlsec1-dev \
        libffi-dev \
        liblzma-dev \
        # LibreOffice and related dependencies
        libreoffice \
        fonts-dejavu-core \
        fonts-freefont-ttf \
        fonts-liberation \
        fonts-opensymbol \
        ghostscript \
        default-jre \
        libreoffice-java-common \
        locales \
    && rm -rf /var/lib/apt/lists/* && \
    # Generate en_US.UTF-8 locale for LibreOffice
    locale-gen en_US.UTF-8 && \
    update-locale LANG=en_US.UTF-8

ENV LANG=en_US.UTF-8
ENV LC_ALL=en_US.UTF-8

# Install pyenv
RUN curl https://pyenv.run | bash

# Install Python 3.10.12 (a stable version)
RUN pyenv install 3.10.12 && \
    pyenv global 3.10.12 && \
    pip install --upgrade pip

# Set the working directory inside the container
WORKDIR /app

# Configure LibreOffice to use more generic rendering (less reliance on specific display server features)
# And set up some environment variables for memory/logging
ENV SAL_USE_VCLPLUGIN=gen
ENV UNO_VERBOSE=true
ENV URE_BOOTSTRAP_LINES=20

# Copy the requirements file and install Python dependencies
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