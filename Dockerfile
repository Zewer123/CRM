FROM python:3.11-slim

# Install system deps needed for psycopg2
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpq-dev gcc \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install psycopg2 separately with system libs available
RUN pip install --no-cache-dir psycopg2-binary==2.9.9

COPY . .

EXPOSE 8000
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--workers", "1", "--timeout", "120", "wsgi:app"]
