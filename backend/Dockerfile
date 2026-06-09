# syntax=docker/dockerfile:1.7

FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends build-essential libpq-dev \
    && rm -rf /var/lib/apt/lists/*

COPY . .

RUN pip install --no-cache-dir --upgrade pip \
    && if [ -f /app/requirements.txt ]; then \
        pip install --no-cache-dir -r /app/requirements.txt; \
    elif [ -f /app/backend/requirements.txt ]; then \
        pip install --no-cache-dir -r /app/backend/requirements.txt; \
    else \
        echo "requirements.txt not found in /app or /app/backend"; \
        exit 1; \
    fi

EXPOSE 8000

CMD ["sh", "-c", "export DJANGO_SETTINGS_MODULE=${DJANGO_SETTINGS_MODULE:-config.settings.production}; if [ -f /app/manage.py ]; then gunicorn config.wsgi:application --bind 0.0.0.0:${PORT:-8000} --workers 3 --log-file - --access-logfile -; else cd /app/backend && gunicorn config.wsgi:application --bind 0.0.0.0:${PORT:-8000} --workers 3 --log-file - --access-logfile -; fi"]


