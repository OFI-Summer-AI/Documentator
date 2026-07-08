# syntax=docker/dockerfile:1.7

FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update \
    && apt-get upgrade -y \
    && apt-get install -y --no-install-recommends build-essential libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy only the backend source so the frontend never bloats the image
COPY backend/ .

RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# Collect static files at build time so ManifestStaticFilesStorage / whitenoise
# have the manifest ready before the first request.
# DATABASE_URL points nowhere — collectstatic does not need a real database.
RUN DJANGO_SETTINGS_MODULE=config.settings.production \
    SECRET_KEY=collectstatic-build-placeholder \
    DATABASE_URL=postgres://localhost/placeholder \
    python manage.py collectstatic --noinput

EXPOSE 8000

# migrate runs once at container start so the schema is always up to date,
# then gunicorn takes over.
CMD ["sh", "-c", \
  "DJANGO_SETTINGS_MODULE=${DJANGO_SETTINGS_MODULE:-config.settings.production} \
   python manage.py migrate --noinput && \
   gunicorn config.wsgi:application \
     --bind 0.0.0.0:${PORT:-8000} \
     --workers 3 \
     --timeout 120 \
     --log-file - \
     --access-logfile -"]
