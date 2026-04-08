FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Chromium + ALL system dependencies automatically
# (playwright install-deps detects the exact OS/version and installs correct packages)
RUN playwright install --with-deps chromium

COPY web_app.py .
COPY templates/ templates/

EXPOSE 10000

CMD ["gunicorn", "web_app:app", "--bind", "0.0.0.0:10000", "--timeout", "300", "--workers", "1", "--threads", "8"]
