FROM python:3.11-slim

# ---- LibreOffice + 中文字体 ----
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-core libreoffice-impress libreoffice-writer \
    fonts-noto-cjk fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY . .

RUN mkdir -p output uploads

EXPOSE 8080

CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "2", "--timeout", "120", "app:app"]
