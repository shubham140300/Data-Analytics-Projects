"""Settings for the local TechCompliance workspace."""
import os
from pathlib import Path
from django.core.management.utils import get_random_secret_key

os.environ["OPENPYXL_DEFUSEDXML"] = "True"

BASE_DIR = Path(__file__).resolve().parent.parent

SECRET_KEY = os.environ.get("DJANGO_SECRET_KEY", "").strip()
if not SECRET_KEY:
    secret_file = BASE_DIR / ".django_secret_key"
    try:
        SECRET_KEY = secret_file.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        generated_secret = get_random_secret_key()
        try:
            with secret_file.open("x", encoding="utf-8") as key_file:
                key_file.write(generated_secret)
            SECRET_KEY = generated_secret
        except FileExistsError:
            SECRET_KEY = secret_file.read_text(encoding="utf-8").strip()

DEBUG = os.environ.get("DJANGO_DEBUG", "0").lower() in {"1", "true", "yes"}
WORKSPACE_OWNER_ID = os.environ.get("WORKSPACE_OWNER_ID", "2798869").strip()
WORKSPACE_OWNER_USERNAME = os.environ.get("WORKSPACE_OWNER_USERNAME", WORKSPACE_OWNER_ID).strip().casefold()
WORKSPACE_OWNER_SESSION_AGE = int(os.environ.get("WORKSPACE_OWNER_SESSION_AGE", str(100 * 365 * 24 * 60 * 60)))
WORKSPACE_DEFAULT_PASSWORD = os.environ.get("WORKSPACE_DEFAULT_PASSWORD", "Password@123")
# ALLOWED_HOSTS = [
#     host.strip()
#     for host in os.environ.get("DJANGO_ALLOWED_HOSTS", "localhost,127.0.0.1").split(",")
#     if host.strip()
# ]
ALLOWED_HOSTS = ["*"]

INSTALLED_APPS = [
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "website.apps.WebsiteConfig",
]

MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    "django.middleware.csp.ContentSecurityPolicyMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "website.access.WorkspaceAccessMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

SECURE_CONTENT_TYPE_NOSNIFF = True
SECURE_REFERRER_POLICY = "same-origin"
X_FRAME_OPTIONS = "DENY"
SESSION_COOKIE_HTTPONLY = True
SESSION_COOKIE_SAMESITE = "Lax"
CSRF_COOKIE_HTTPONLY = True
CSRF_COOKIE_SAMESITE = "Lax"
SECURE_SSL_REDIRECT = os.environ.get("DJANGO_SECURE_SSL_REDIRECT", "0").lower() in {"1", "true", "yes"}
SECURE_COOKIES = os.environ.get("DJANGO_SECURE_COOKIES", "0").lower() in {"1", "true", "yes"}
SESSION_COOKIE_SECURE = SECURE_COOKIES
CSRF_COOKIE_SECURE = SECURE_COOKIES
SECURE_HSTS_SECONDS = int(os.environ.get("DJANGO_HSTS_SECONDS", "0"))
SECURE_HSTS_INCLUDE_SUBDOMAINS = SECURE_HSTS_SECONDS > 0
SECURE_HSTS_PRELOAD = False

# The app has no remote fonts, scripts, images, or APIs. Inline styles are used only
# for server-rendered chart bar heights.
SECURE_CSP = {
    "default-src": ["'self'"],
    "base-uri": ["'self'"],
    "form-action": ["'self'"],
    "frame-ancestors": ["'none'"],
    "frame-src": ["'none'"],
    "font-src": ["'self'"],
    "img-src": ["'self'", "data:"],
    "object-src": ["'none'"],
    "script-src": ["'self'"],
    "style-src": ["'self'", "'unsafe-inline'"],
}

ROOT_URLCONF = "config.urls"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [BASE_DIR / "templates"],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
                "website.context_processors.workspace_access",
            ],
        },
    },
]

WSGI_APPLICATION = "config.wsgi.application"
ASGI_APPLICATION = "config.asgi.application"

DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
    }
}

LANGUAGE_CODE = "en-us"
TIME_ZONE = "Asia/Kolkata"
USE_I18N = True
USE_TZ = True

STATIC_URL = "static/"
STATICFILES_DIRS = [BASE_DIR / "static"]
MEDIA_URL = "media/"
MEDIA_ROOT = BASE_DIR / "uploads"

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"
DATA_UPLOAD_MAX_MEMORY_SIZE = 25 * 1024 * 1024
FILE_UPLOAD_MAX_MEMORY_SIZE = 2_500_000
DATA_UPLOAD_MAX_NUMBER_FILES = 1
DATA_UPLOAD_MAX_NUMBER_FIELDS = 1000
