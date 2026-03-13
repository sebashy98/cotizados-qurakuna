FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    fonts-liberation \
    --no-install-recommends \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /tmp/cotizaciones

EXPOSE 8080

ENV PORT=8080

CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--timeout", "120", "app:app"]
