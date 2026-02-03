FROM python:3.9-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-eng \
    poppler-utils \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY sms_scanner_app.py .

# Expose port
EXPOSE 8501

# Run application
CMD ["streamlit", "run", "sms_scanner_app.py", "--server.port=8501", "--server.address=0.0.0.0"]
