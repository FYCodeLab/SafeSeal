FROM python:3.11-slim

# Install LibreOffice (headless) and fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# Render sets $PORT; default for local runs
ENV PORT=10000
EXPOSE 10000

CMD streamlit run app.py --server.port $PORT --server.address 0.0.0.0
