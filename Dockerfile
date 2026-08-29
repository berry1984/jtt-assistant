FROM python:3.11-slim

WORKDIR /app

# Copy requirements
COPY requirements.txt .
COPY TR账单自动生成/requirements.txt ./TR_requirements.txt

# Install fonts so PDF insert_text uses Arial-compatible metrics (not helv fallback)
RUN apt-get update && apt-get install -y --no-install-recommends fonts-liberation && \
    rm -rf /var/lib/apt/lists/*

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt && \
    pip install --no-cache-dir -r TR_requirements.txt

# Copy all application code
COPY . .

# Run from the app directory
WORKDIR /app/TR账单自动生成

CMD ["python", "app.py"]
