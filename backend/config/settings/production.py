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

known_good_frontend_origins = ["https://documentator-omega.vercel.app"]
configured_cors_origins = env.list("CORS_ALLOWED_ORIGINS", default=[])
configured_csrf_origins = env.list("CSRF_TRUSTED_ORIGINS", default=[])

# Merge rather than replace: if the CORS_ALLOWED_ORIGINS/CSRF_TRUSTED_ORIGINS env vars are set on
# the host to a stale or incomplete value, the known-good frontend origin still gets through
# instead of silently breaking every cross-origin request (mirrors the ALLOWED_HOSTS merge above).
CORS_ALLOWED_ORIGINS = list(dict.fromkeys([*configured_cors_origins, *known_good_frontend_origins]))
CSRF_TRUSTED_ORIGINS = list(dict.fromkeys([*configured_csrf_origins, *known_good_frontend_origins]))

MIDDLEWARE = ["whitenoise.middleware.WhiteNoiseMiddleware", *MIDDLEWARE]
STATICFILES_STORAGE = "django.contrib.staticfiles.storage.ManifestStaticFilesStorage"
