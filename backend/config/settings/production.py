from __future__ import annotations

from .base import *  # noqa: F403,F401

env = environ.Env()

DEBUG = False

configured_allowed_hosts = env.list(
	"ALLOWED_HOSTS",
	default=[
		"127.0.0.1",
		"localhost",
	],
)
railway_allowed_hosts = [
	".up.railway.app",
	".railway.app",
	".railway.internal",
]
ALLOWED_HOSTS = list(dict.fromkeys([*configured_allowed_hosts, *railway_allowed_hosts]))

USE_SQLITE_FOR_DEMO = env.bool("USE_SQLITE_FOR_DEMO", default=False)
if USE_SQLITE_FOR_DEMO:
	DATABASES = {
		"default": {
			"ENGINE": "django.db.backends.sqlite3",
			"NAME": BASE_DIR / "db.sqlite3",
			"ATOMIC_REQUESTS": True,
		}
	}

SECURE_PROXY_SSL_HEADER = ("HTTP_X_FORWARDED_PROTO", "https")
SECURE_SSL_REDIRECT = False
SECURE_REDIRECT_EXEMPT = [r"^health/$"]
SESSION_COOKIE_SECURE = True
CSRF_COOKIE_SECURE = True
SECURE_CONTENT_TYPE_NOSNIFF = True
X_FRAME_OPTIONS = "DENY"

CORS_ALLOWED_ORIGINS = env.list(
	"CORS_ALLOWED_ORIGINS",
	default=["https://documentator-omega.vercel.app"],
)
CSRF_TRUSTED_ORIGINS = env.list(
	"CSRF_TRUSTED_ORIGINS",
	default=["https://documentator-omega.vercel.app"],
)

MIDDLEWARE = ["whitenoise.middleware.WhiteNoiseMiddleware", *MIDDLEWARE]
STATICFILES_STORAGE = "django.contrib.staticfiles.storage.ManifestStaticFilesStorage"
