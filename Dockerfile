# Use Python 3.10 for compatibility with PySide6
FROM python:3.10-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=8000

# Install system dependencies
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        poppler-utils \
        curl \
        fonts-noto-cjk \
        fonts-ipafont \
        fonts-wqy-zenhei \
    && rm -rf /var/lib/apt/lists/*

# Set work directory
WORKDIR /app

# Copy requirements first for better caching
COPY requirements_web.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements_web.txt

# Copy application code
COPY . .

# Create uploads directory
RUN mkdir -p /app/uploads /app/temp

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8000/ || exit 1

# Run the application
CMD ["python", "web_app.py"]
