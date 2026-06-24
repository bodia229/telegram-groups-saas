"""Application configuration loaded from environment variables."""

import os

# Telegram API credentials — https://my.telegram.org
API_ID = int(os.getenv("TG_API_ID", "0"))
API_HASH = os.getenv("TG_API_HASH", "")
SESSION_NAME = os.getenv("TG_SESSION", "ru_groups_session")

# Storage
DB_PATH = os.getenv("DB_PATH", "telegram_groups.db")
CSV_PATH = os.getenv("CSV_PATH", "Telegram_groups_russia.csv")
XLSX_PATH = os.getenv("XLSX_PATH", "Telegram_groups_russia.xlsx")

# Collection scale
TARGET_GROUPS = int(os.getenv("TARGET_GROUPS", "100000"))
SEED_TARGET = int(os.getenv("SEED_TARGET", "5000"))

# Concurrency / rate limiting
WORKERS = int(os.getenv("WORKERS", "8"))
MIN_DELAY = float(os.getenv("MIN_DELAY", "1.0"))
MAX_DELAY = float(os.getenv("MAX_DELAY", "3.0"))
MESSAGES_PER_GROUP = int(os.getenv("MESSAGES_PER_GROUP", "200"))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "3"))

# API
API_KEY = os.getenv("API_KEY", "")  # optional bearer token to protect the SaaS API
