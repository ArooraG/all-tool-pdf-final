# Debian Linux کے ایک مستحکم ورژن سے شروع کریں
FROM python:3.9-slim-bullseye

# ضروری سسٹم ٹولز، فونٹس، LibreOffice اور Java انسٹال کریں
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    default-jre \
    ghostscript \
    fonts-liberation \
    fonts-opensymbol \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# کام کرنے کے لیے ایک فولڈر بنائیں
WORKDIR /app

# پہلے صرف requirements فائل کاپی کریں
COPY requirements.txt .

# تمام Python لائبریریاں انسٹال کریں
RUN pip install --no-cache-dir -r requirements.txt

# باقی کا پورا پروجیکٹ کوڈ کاپی کریں
COPY . .

# سرور کو Gunicorn کے ساتھ چلائیں
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000", "--workers", "2", "--timeout", "300"]