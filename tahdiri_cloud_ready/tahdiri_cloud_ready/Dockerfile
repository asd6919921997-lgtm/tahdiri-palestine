# Use official lightweight Python image
FROM python:3.11-slim

# Install system dependencies: LibreOffice for PDF conversion + Arabic Fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer-nogui \
    fonts-noto-core \
    fonts-sil-scheherazade \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Environment variables
ENV PORT=8000
EXPOSE 8000

# Start application (respecting dynamic $PORT from cloud hosts like Render/Railway/Koyeb)
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
