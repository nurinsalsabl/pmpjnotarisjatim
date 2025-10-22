FROM python:3.11-slim

# --- Install dependency sistem (poppler & tesseract) ---
RUN apt-get update && apt-get install -y \
    poppler-utils \
    tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

# --- Set working directory ---
WORKDIR /app

# --- Copy semua file ke container ---
COPY . /app

# --- Install dependensi Python ---
RUN pip install --no-cache-dir -r requirements.txt

# --- Jalankan Streamlit ---
EXPOSE 8501
CMD ["streamlit", "run", "stkanwil.py", "--server.port=8501", "--server.address=0.0.0.0"]
