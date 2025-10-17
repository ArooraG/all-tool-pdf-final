# Base image for your Python application
# Aap isko badal sakte hain agar aap koi aur base image use kar rahe hain
FROM python:3.9-slim-buster

# System dependencies ko update aur install karein
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
# Gunicorn ko bhi install karein agar aap CMD mein use kar rahe hain
RUN pip install gunicorn

# Baaqi application files copy karein
COPY . .

# Environment variables set karein
ENV FLASK_APP=app.py
ENV FLASK_ENV=production # Ya development, jo aapka preference ho

# Port expose karein
EXPOSE 10000

# Application ko run karne ki command
CMD gunicorn -w 4 -b 0.0.0.0:$PORT app:app
# Render $PORT environment variable use karta hai, 10000 nahi.
# Agar aap sirf Flask ka default run method use kar rahe hote, toh:
# CMD ["flask", "run", "--host=0.0.0.0", "--port", "10000"]
# Lekin gunicorn zyada production-ready hai.