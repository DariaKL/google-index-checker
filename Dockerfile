FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY web_app.py .
COPY templates/ templates/

EXPOSE 10000

CMD ["gunicorn", "web_app:app", "--bind", "0.0.0.0:10000", "--timeout", "300", "--workers", "1", "--threads", "8"]
