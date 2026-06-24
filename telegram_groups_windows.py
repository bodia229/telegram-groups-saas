#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Telegram Russian Groups Collector — SINGLE FILE / Windows 10 edition.

Запуск:
    1. Установи Python 3.10+  (https://www.python.org/downloads/  — галочка "Add to PATH")
    2. Двойной клик по файлу  ИЛИ  в консоли:  python telegram_groups_windows.py
    3. При первом запуске введи API_ID и API_HASH (бери на https://my.telegram.org),
       затем номер телефона и код из Telegram.

Зависимости ставятся автоматически. Результат: Telegram_groups_russia.csv / .xlsx
"""

import os
import sys

# ──────────────────────────────────────────────────────────────────────────────
# Windows: корректная кодировка консоли (кириллица) + авто-установка зависимостей
# ──────────────────────────────────────────────────────────────────────────────

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _ensure_deps():
    import importlib.util
    import subprocess
    needed = {
        "telethon": "telethon",
        "pandas": "pandas",
        "tqdm": "tqdm",
        "openpyxl": "openpyxl",
    }
    missing = [pip for mod, pip in needed.items()
               if importlib.util.find_spec(mod) is None]
    if missing:
        print(f"Устанавливаю зависимости: {', '.join(missing)} …")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
        print("Зависимости установлены.\n")


_ensure_deps()

import re
import csv
import json
import random
import asyncio
import logging
import sqlite3
import urllib.parse
import urllib.request
from datetime import datetime

import pandas as pd
from tqdm import tqdm

from telethon import TelegramClient, functions
from telethon.errors import (
    FloodWaitError,
    ChannelPrivateError,
    UsernameInvalidError,
    UsernameNotOccupiedError,
    RPCError,
)
from telethon.tl.types import Channel, Chat, User
from telethon.tl.functions.contacts import SearchRequest

# ──────────────────────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────────────────────
#
# API_ID / API_HASH уже вшиты (публичные креды Telegram Desktop) — ничего
# регистрировать и вводить не нужно. При желании можно подставить свои через
# переменные окружения TG_API_ID / TG_API_HASH (безопаснее: меньше риск бана).
#
# ВНИМАНИЕ: вход по номеру телефона и коду из Telegram ОБЯЗАТЕЛЕН — без этого
# ни один user-клиент Telegram работать не может, это требование самого Telegram.

DEFAULT_API_ID = 2040
DEFAULT_API_HASH = "b18441a1ff607e10a989891a5462e627"

API_ID = int(os.getenv("TG_API_ID", str(DEFAULT_API_ID)))
API_HASH = os.getenv("TG_API_HASH", DEFAULT_API_HASH)
SESSION_NAME = os.getenv("TG_SESSION", "ru_groups_session")

# TGStat API — токен из личного кабинета https://api.tgstat.ru (платный).
# Если пусто — интеграция просто пропускается.
TGSTAT_TOKEN = os.getenv("TGSTAT_TOKEN", "")
TGSTAT_LIMIT = int(os.getenv("TGSTAT_LIMIT", "50"))          # результатов на запрос
TGSTAT_MAX_QUERIES = int(os.getenv("TGSTAT_MAX_QUERIES", "600"))  # потолок запросов (квота!)

# Бесплатные методы поиска
SEED_SEARCH_MAX_QUERIES = int(os.getenv("SEED_SEARCH_MAX_QUERIES", "400"))  # глоб. поиск Telegram
SEED_PATTERN_MAX = int(os.getenv("SEED_PATTERN_MAX", "400"))                # перебор username-шаблонов

DB_PATH = "telegram_groups.db"
CSV_PATH = "Telegram_groups_russia.csv"
XLSX_PATH = "Telegram_groups_russia.xlsx"

TARGET_GROUPS = int(os.getenv("TARGET_GROUPS", "100000"))
SEED_TARGET = int(os.getenv("SEED_TARGET", "5000"))

WORKERS = int(os.getenv("WORKERS", "8"))         # 5–10 параллельных задач
MIN_DELAY = 1.0                                  # задержки 1–3 сек
MAX_DELAY = 3.0
MESSAGES_PER_GROUP = 200
MAX_RETRIES = 3
# Если Telegram просит ждать дольше этого (сек) — пропускаем запрос, а не висим
FLOOD_SKIP_THRESHOLD = int(os.getenv("FLOOD_SKIP_THRESHOLD", "300"))
# Максимум секунд на ОДНУ группу. Дольше — бросаем и идём к следующей
GROUP_TIMEOUT = int(os.getenv("GROUP_TIMEOUT", "90"))

class _TqdmLoggingHandler(logging.Handler):
    """Печатает логи через tqdm.write, чтобы не ломать прогресс-бар."""
    def emit(self, record):
        try:
            tqdm.write(self.format(record))
            self.flush()
        except Exception:
            self.handleError(record)


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler("collector.log", encoding="utf-8"),
        _TqdmLoggingHandler(),
    ],
)
log = logging.getLogger("collector")

# ──────────────────────────────────────────────────────────────────────────────
# SEED KEYWORDS
# ──────────────────────────────────────────────────────────────────────────────

SEED_KEYWORDS = [
    "Москва чат", "СПБ чат", "Питер чат", "Россия чат",
    "работа чат Москва", "работа чат Питер", "вакансии чат",
    "аренда чат Москва", "аренда квартир чат", "недвижимость чат",
    "объявления чат", "барахолка чат", "куплю продам чат",
    "знакомства чат", "общение чат", "флудилка чат",
    "бизнес чат", "стартап чат", "крипта чат", "инвестиции чат",
    "Казань чат", "Новосибирск чат", "Екатеринбург чат",
    "Нижний Новгород чат", "Челябинск чат", "Самара чат",
    "Омск чат", "Ростов чат", "Уфа чат", "Краснодар чат",
    "Воронеж чат", "Пермь чат", "Волгоград чат", "Сочи чат",
    "Тюмень чат", "Саратов чат", "Тольятти чат", "Ижевск чат",
    "Барнаул чат", "Иркутск чат", "Хабаровск чат", "Владивосток чат",
    "Ярославль чат", "Махачкала чат", "Томск чат", "Оренбург чат",
    "Кемерово чат", "Рязань чат", "Тула чат", "Липецк чат",
    "Калининград чат", "Ставрополь чат", "мамочки чат", "автолюбители чат",
    "путешествия чат", "ремонт чат", "строительство чат", "юристы чат",
    "фриланс чат", "дизайн чат", "айти чат", "разработка чат",
]

# Известные публичные чаты/каталоги — гарантированные точки входа для графа.
# Если какой-то не существует — get_entity вернёт None, это не страшно.
SEED_USERNAMES = [
    "ru_python", "pythonchatru", "ru_python_beginners",
    "natandev", "tproger_chat", "devschat", "JavaScript_ru",
    "moscow", "spb", "ru", "chat", "obyavleniya", "rabota",
    "arenda_msk", "rabota_moskva", "nedvizhimost_chat",
    "baraholka", "kupiprodai", "znakomstva_chat",
    "freelance", "smm_chat", "marketing_chat", "crypto_ru",
    "investing_chat", "auto_chat", "remont_chat", "mamochki_chat",
]

# ~100 крупнейших городов России — основа для генерации запросов и username-шаблонов
RU_CITIES = [
    "Москва", "Санкт-Петербург", "Новосибирск", "Екатеринбург", "Казань",
    "Нижний Новгород", "Челябинск", "Самара", "Уфа", "Ростов-на-Дону",
    "Краснодар", "Омск", "Воронеж", "Пермь", "Волгоград", "Саратов",
    "Тюмень", "Тольятти", "Барнаул", "Ижевск", "Ульяновск", "Иркутск",
    "Хабаровск", "Махачкала", "Ярославль", "Владивосток", "Томск",
    "Оренбург", "Кемерово", "Новокузнецк", "Рязань", "Астрахань",
    "Пенза", "Липецк", "Тула", "Киров", "Чебоксары", "Калининград",
    "Балашиха", "Курск", "Севастополь", "Сочи", "Ставрополь", "Улан-Удэ",
    "Тверь", "Магнитогорск", "Иваново", "Брянск", "Белгород", "Сургут",
    "Владимир", "Нижний Тагил", "Архангельск", "Чита", "Калуга",
    "Смоленск", "Волжский", "Якутск", "Саранск", "Череповец", "Курган",
    "Вологда", "Орёл", "Подольск", "Грозный", "Владикавказ", "Мурманск",
    "Тамбов", "Стерлитамак", "Петрозаводск", "Кострома", "Нижневартовск",
    "Новороссийск", "Йошкар-Ола", "Таганрог", "Комсомольск-на-Амуре",
    "Сыктывкар", "Нальчик", "Шахты", "Нижнекамск", "Дзержинск", "Братск",
    "Орск", "Ангарск", "Энгельс", "Благовещенск", "Старый Оскол",
    "Великий Новгород", "Бийск", "Прокопьевск", "Псков", "Балаково",
    "Армавир", "Рыбинск", "Северодвинск", "Абакан", "Норильск",
    "Сызрань", "Каменск-Уральский", "Новочеркасск",
]

# Темы, по которым обычно создают именно ЧАТЫ (а не каналы)
CHAT_TOPICS = [
    "чат", "общение", "объявления", "барахолка", "подслушано", "типичный",
    "работа", "вакансии", "аренда", "недвижимость", "знакомства",
    "куплю продам", "новости", "афиша", "мамочки", "автолюбители",
    "бизнес", "флудилка", "переезд", "туризм",
]


def build_queries(limit=None):
    """Запросы для поиска: SEED_KEYWORDS + города×темы + темы."""
    out, seen = [], set()

    def add(q):
        k = q.lower()
        if k not in seen:
            seen.add(k)
            out.append(q)

    for kw in SEED_KEYWORDS:
        add(kw)
    for city in RU_CITIES:
        add(f"{city} чат")
        add(f"{city} объявления")
        add(f"{city} барахолка")
        add(f"подслушано {city}")
        add(f"типичный {city}")
        add(f"{city} работа")
        add(f"{city} аренда")
        add(f"{city} знакомства")
    for t in CHAT_TOPICS:
        add(t)
    return out[:limit] if limit else out


_TRANSLIT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
    " ": "", "-": "",
}


def translit(s):
    return "".join(_TRANSLIT.get(c, c) for c in s.lower())


def build_username_candidates(limit=None):
    """Кандидаты публичных username по городам: moskvachat, chat_kazan, podslushano_perm …"""
    post = ["_chat", "chat", "_obyavleniya", "_baraholka", "_news", "_work", "_official"]
    pre = ["chat_", "podslushano_", "tipich_", "afisha_", "rabota_", "arenda_", "news_"]
    out, seen = [], set()

    def add(u):
        if u and 5 <= len(u) <= 32 and u.lower() not in seen:
            seen.add(u.lower())
            out.append(u)

    for city in RU_CITIES:
        t = translit(city)
        if not t:
            continue
        for s in post:
            add(f"{t}{s}")
        for s in pre:
            add(f"{s}{t}")
    return out[:limit] if limit else out

# ──────────────────────────────────────────────────────────────────────────────
# CITY DETECTION
# ──────────────────────────────────────────────────────────────────────────────

CITY_MAP = {
    "Moscow": ["москв", "moscow", "мск", "msk"],
    "Saint Petersburg": ["санкт-петербург", "спб", "питер", "petersburg", "spb"],
    "Kazan": ["казан", "kazan"],
    "Novosibirsk": ["новосибирск", "novosibirsk", "нск"],
    "Yekaterinburg": ["екатеринбург", "yekaterinburg", "екб"],
    "Nizhny Novgorod": ["нижний новгород", "нижнем новгороде", "nizhny"],
    "Chelyabinsk": ["челябинск", "chelyabinsk"],
    "Samara": ["самар", "samara"],
    "Omsk": ["омск", "omsk"],
    "Rostov-on-Don": ["ростов", "rostov"],
    "Ufa": ["уфа", "ufa"],
    "Krasnodar": ["краснодар", "krasnodar"],
    "Voronezh": ["воронеж", "voronezh"],
    "Perm": ["пермь", "perm"],
    "Volgograd": ["волгоград", "volgograd"],
    "Sochi": ["сочи", "sochi"],
    "Tyumen": ["тюмень", "tyumen"],
    "Saratov": ["саратов", "saratov"],
    "Tolyatti": ["тольятти", "tolyatti"],
    "Izhevsk": ["ижевск", "izhevsk"],
    "Barnaul": ["барнаул", "barnaul"],
    "Irkutsk": ["иркутск", "irkutsk"],
    "Khabarovsk": ["хабаровск", "khabarovsk"],
    "Vladivostok": ["владивосток", "vladivostok"],
    "Yaroslavl": ["ярославль", "yaroslavl"],
    "Makhachkala": ["махачкал", "makhachkala"],
    "Tomsk": ["томск", "tomsk"],
    "Orenburg": ["оренбург", "orenburg"],
    "Kemerovo": ["кемерово", "kemerovo"],
    "Ryazan": ["рязань", "ryazan"],
    "Tula": ["тула", "tula"],
    "Lipetsk": ["липецк", "lipetsk"],
    "Kaliningrad": ["калининград", "kaliningrad"],
    "Stavropol": ["ставрополь", "stavropol"],
}

USERNAME_RE = re.compile(r"(?:@|t\.me/|telegram\.me/)([A-Za-z][A-Za-z0-9_]{3,31})")
INVITE_RE = re.compile(r"(?:t\.me/joinchat/|t\.me/\+|telegram\.me/joinchat/)([A-Za-z0-9_-]{8,})")


def detect_city(*texts):
    blob = " ".join(t for t in texts if t).lower()
    for city, keys in CITY_MAP.items():
        for k in keys:
            if k in blob:
                return city
    return None


# ──────────────────────────────────────────────────────────────────────────────
# DATABASE
# ──────────────────────────────────────────────────────────────────────────────

class DB:
    def __init__(self, path):
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL;")
        self.conn.execute("PRAGMA synchronous=NORMAL;")
        self._init_schema()
        self._lock = asyncio.Lock()

    def _init_schema(self):
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS groups (
                id                  INTEGER PRIMARY KEY,
                title               TEXT,
                username            TEXT,
                type                TEXT,
                participants_count  INTEGER,
                description         TEXT,
                city                TEXT,
                source              TEXT,
                created_at          TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS queue (
                group_id  INTEGER PRIMARY KEY,
                username  TEXT,
                status    TEXT DEFAULT 'new'
            )
        """)
        cur.execute("CREATE INDEX IF NOT EXISTS idx_queue_status ON queue(status)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_groups_username ON groups(username)")
        self.conn.commit()

    async def upsert_group(self, g):
        async with self._lock:
            cur = self.conn.cursor()
            cur.execute("SELECT 1 FROM groups WHERE id = ?", (g["id"],))
            exists = cur.fetchone() is not None
            cur.execute("""
                INSERT INTO groups
                    (id, title, username, type, participants_count,
                     description, city, source, created_at)
                VALUES (:id, :title, :username, :type, :participants_count,
                        :description, :city, :source, :created_at)
                ON CONFLICT(id) DO UPDATE SET
                    title=excluded.title,
                    username=COALESCE(excluded.username, groups.username),
                    participants_count=COALESCE(excluded.participants_count, groups.participants_count),
                    description=COALESCE(excluded.description, groups.description),
                    city=COALESCE(excluded.city, groups.city)
            """, g)
            self.conn.commit()
            return not exists

    async def enqueue(self, group_id, username):
        async with self._lock:
            self.conn.execute(
                "INSERT OR IGNORE INTO queue (group_id, username, status) VALUES (?, ?, 'new')",
                (group_id, username),
            )
            self.conn.commit()

    async def next_batch(self, limit):
        async with self._lock:
            cur = self.conn.cursor()
            cur.execute(
                "SELECT group_id, username FROM queue WHERE status='new' LIMIT ?",
                (limit,),
            )
            rows = cur.fetchall()
            ids = [r[0] for r in rows]
            if ids:
                self.conn.executemany(
                    "UPDATE queue SET status='processing' WHERE group_id=?",
                    [(i,) for i in ids],
                )
                self.conn.commit()
            return rows

    async def mark_processed(self, group_id):
        async with self._lock:
            self.conn.execute(
                "UPDATE queue SET status='processed' WHERE group_id=?", (group_id,)
            )
            self.conn.commit()

    def count_groups(self):
        return self.conn.execute("SELECT COUNT(*) FROM groups").fetchone()[0]

    def count_new_queue(self):
        return self.conn.execute(
            "SELECT COUNT(*) FROM queue WHERE status='new'"
        ).fetchone()[0]

    def export(self):
        df = pd.read_sql_query(
            "SELECT * FROM groups ORDER BY participants_count DESC", self.conn
        )
        df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
        try:
            df.to_excel(XLSX_PATH, index=False, engine="openpyxl")
        except Exception as e:
            log.warning("XLSX export failed (%s); CSV saved.", e)
        return len(df)


# ──────────────────────────────────────────────────────────────────────────────
# ENTITY CLASSIFICATION
# ──────────────────────────────────────────────────────────────────────────────

def classify(entity):
    """Return 'group'/'supergroup' for keepable entities, else None."""
    if isinstance(entity, Chat):
        if getattr(entity, "deactivated", False):
            return None
        return "group"
    if isinstance(entity, Channel):
        if entity.broadcast:                      # канал — исключаем
            return None
        if entity.megagroup or getattr(entity, "gigagroup", False):
            return "supergroup"
        return None
    if isinstance(entity, User):                  # боты / пользователи — исключаем
        return None
    return None


def entity_to_row(entity, etype, source, description=None):
    title = getattr(entity, "title", None)
    username = getattr(entity, "username", None)
    participants = getattr(entity, "participants_count", None)
    return {
        "id": entity.id,
        "title": title,
        "username": username,
        "type": etype,
        "participants_count": participants,
        "description": description,
        "city": detect_city(title, username, description),
        "source": source,
        "created_at": datetime.utcnow().isoformat(),
    }


# ──────────────────────────────────────────────────────────────────────────────
# RATE-LIMIT / RETRY
# ──────────────────────────────────────────────────────────────────────────────

async def tgstat_search(token, query, country="ru", limit=50):
    """Поиск чатов/каналов в TGStat. Возвращает список item-словарей."""
    params = urllib.parse.urlencode({
        "token": token,
        "q": query,
        "country": country,
        "limit": limit,
    })
    url = "https://api.tgstat.ru/channels/search?" + params

    def _do():
        req = urllib.request.Request(url, headers={"User-Agent": "collector/1.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode("utf-8"))

    try:
        data = await asyncio.to_thread(_do)
    except Exception as e:
        log.warning("TGStat запрос упал (%s): %s", query, e)
        return []
    if data.get("status") != "ok":
        log.warning("TGStat ответ не ok по «%s»: %s", query, str(data)[:200])
        return []
    return data.get("response", {}).get("items", [])


def _username_from_item(it):
    u = (it.get("username") or "").lstrip("@")
    if not u and it.get("link"):
        u = it["link"].rstrip("/").split("/")[-1].lstrip("@")
    return u or None


async def jitter():
    await asyncio.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


async def safe_call(coro_func, *args, **kwargs):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return await coro_func(*args, **kwargs)
        except FloodWaitError as e:
            if e.seconds > FLOOD_SKIP_THRESHOLD:
                log.warning("FloodWait %ds > порога %ds — пропускаю запрос",
                            e.seconds, FLOOD_SKIP_THRESHOLD)
                return None
            wait = e.seconds + random.uniform(1, 5)
            log.warning("FloodWait: жду %.0fs", wait)
            await asyncio.sleep(wait)
        except (ChannelPrivateError, UsernameInvalidError, UsernameNotOccupiedError):
            return None
        except RPCError as e:
            log.debug("RPCError (%s) attempt %d", e, attempt)
            await asyncio.sleep(2 * attempt)
        except Exception as e:
            log.debug("Error %s attempt %d", e, attempt)
            await asyncio.sleep(2 * attempt)
    return None


# ──────────────────────────────────────────────────────────────────────────────
# COLLECTOR
# ──────────────────────────────────────────────────────────────────────────────

class Collector:
    def __init__(self, client, db):
        self.client = client
        self.db = db
        self.pbar = None

    async def register(self, entity, source, description=None):
        etype = classify(entity)
        if etype is None:
            return False
        row = entity_to_row(entity, etype, source, description)
        is_new = await self.db.upsert_group(row)
        await self.db.enqueue(entity.id, row["username"])
        if is_new:
            total = self.db.count_groups()
            uname = f"@{row['username']}" if row["username"] else "—"
            city = row["city"] or "?"
            log.info("➕ [%s] %s | %s | %s | source=%s | всего: %d",
                     etype, (row["title"] or "")[:40], uname, city, source, total)
            if self.pbar:
                self.pbar.n = total
                self.pbar.refresh()
        return is_new

    async def seed_search(self):
        queries = build_queries(SEED_SEARCH_MAX_QUERIES)
        log.info("🔎 Глоб. поиск Telegram (бесплатно): %d запросов", len(queries))
        for i, kw in enumerate(queries, 1):
            if self.db.count_groups() >= TARGET_GROUPS:
                break
            log.info("🔎 [%d/%d] Поиск по запросу: «%s»", i, len(queries), kw)
            result = await safe_call(self.client, SearchRequest(q=kw, limit=100))
            if not result:
                log.info("   ничего не найдено / лимит")
                await jitter()
                continue
            before = self.db.count_groups()
            raw = len(result.chats)
            channels = 0
            for chat in result.chats:
                added = await self.register(chat, source="seed")
                # глобальный поиск возвращает много каналов — берём их чат-обсуждение
                if isinstance(chat, Channel) and chat.broadcast:
                    channels += 1
                    await self.try_linked_group(chat, source="seed")
            log.info("   по «%s»: получено %d чатов, добавлено групп %d (каналов %d)",
                     kw, raw, self.db.count_groups() - before, channels)
            await jitter()
        log.info("✅ Seed-поиск завершён. Собрано групп: %d", self.db.count_groups())

    async def try_linked_group(self, channel, source):
        """У канала часто есть привязанная группа-обсуждение (supergroup) — берём её."""
        try:
            full = await safe_call(
                self.client,
                functions.channels.GetFullChannelRequest(channel=channel),
            )
            if not full:
                return
            linked_id = getattr(full.full_chat, "linked_chat_id", None)
            if not linked_id:
                return
            for ch in full.chats:
                if ch.id == linked_id and classify(ch):
                    await self.register(ch, source=source)
        except Exception:
            pass

    async def process_group(self, group_id, username):
        ref = username or group_id
        label = f"@{username}" if username else f"id={group_id}"
        log.info("🔄 Обрабатываю группу %s …", label)
        entity = await safe_call(self.client.get_entity, ref)
        if entity is None or classify(entity) is None:
            log.info("   пропуск %s (недоступна / не группа)", label)
            await self.db.mark_processed(group_id)
            return

        description = None
        try:
            if isinstance(entity, Channel):
                full = await safe_call(
                    self.client,
                    functions.channels.GetFullChannelRequest(channel=entity),
                )
                if full:
                    description = full.full_chat.about
                    await self.db.upsert_group(
                        entity_to_row(entity, classify(entity), "seed", description)
                    )
        except Exception:
            description = None

        found_usernames = set()
        invite_hashes = set()
        try:
            async for msg in self.client.iter_messages(entity, limit=MESSAGES_PER_GROUP):
                if msg.message:
                    for m in USERNAME_RE.findall(msg.message):
                        found_usernames.add(m.lower())
                    for h in INVITE_RE.findall(msg.message):
                        invite_hashes.add(h)
                fwd = msg.forward
                if fwd is not None:
                    fchat = getattr(fwd, "chat", None)
                    if fchat is not None:
                        await self.register(fchat, source="graph")
                    fname = getattr(fwd, "from_name", None)
                    if fname:
                        for m in USERNAME_RE.findall(fname):
                            found_usernames.add(m.lower())
        except FloodWaitError as e:
            if e.seconds > FLOOD_SKIP_THRESHOLD:
                log.warning("   FloodWait %ds при чтении %s — пропускаю", e.seconds, label)
            else:
                log.warning("   FloodWait %ds при чтении %s — жду", e.seconds, label)
                await asyncio.sleep(e.seconds + 3)
        except Exception as e:
            log.debug("iter_messages failed for %s: %s", ref, e)

        log.info("   %s: упоминаний %d, invite-ссылок %d",
                 label, len(found_usernames), len(invite_hashes))

        # invite-ссылки: резолвим, что доступно без вступления (уже участник)
        for h in invite_hashes:
            if self.db.count_groups() >= TARGET_GROUPS:
                break
            res = await safe_call(
                self.client, functions.messages.CheckChatInviteRequest(hash=h)
            )
            chat = getattr(res, "chat", None) if res is not None else None
            if chat is not None and classify(chat):
                await self.register(chat, source="invite")
            await asyncio.sleep(random.uniform(0.3, 0.8))

        for uname in found_usernames:
            if self.db.count_groups() >= TARGET_GROUPS:
                break
            new_entity = await safe_call(self.client.get_entity, uname)
            if new_entity is not None:
                if classify(new_entity):
                    await self.register(new_entity, source="graph")
                elif isinstance(new_entity, Channel) and new_entity.broadcast:
                    await self.try_linked_group(new_entity, source="graph")
            await asyncio.sleep(random.uniform(0.3, 1.0))

        await self.db.mark_processed(group_id)
        await jitter()

    async def worker(self, name):
        log.info("👷 Воркер #%d запущен", name)
        while self.db.count_groups() < TARGET_GROUPS:
            batch = await self.db.next_batch(1)
            if not batch:
                await asyncio.sleep(5)
                if self.db.count_new_queue() == 0:
                    log.info("👷 Воркер #%d: очередь пуста, завершаюсь", name)
                    return
                continue
            group_id, username = batch[0]
            try:
                await asyncio.wait_for(
                    self.process_group(group_id, username), timeout=GROUP_TIMEOUT
                )
            except asyncio.TimeoutError:
                log.warning("⏭ Группа %s дольше %dс — бросаю, иду к следующей",
                            username or group_id, GROUP_TIMEOUT)
                await self.db.mark_processed(group_id)
            except Exception as e:
                log.debug("worker %d error: %s", name, e)
                await self.db.mark_processed(group_id)

    async def seed_from_tgstat(self):
        if not TGSTAT_TOKEN:
            log.info("ℹ️ TGStat не подключён (нет TGSTAT_TOKEN) — пропускаю.")
            return
        queries = build_queries(TGSTAT_MAX_QUERIES)
        log.info("📡 TGStat: ищу чаты по %d запросам…", len(queries))
        before = self.db.count_groups()
        seen = set()
        for kw in queries:
            if self.db.count_groups() >= TARGET_GROUPS:
                break
            items = await tgstat_search(TGSTAT_TOKEN, kw, limit=TGSTAT_LIMIT)
            usernames = []
            for it in items:
                u = _username_from_item(it)
                if u and u.lower() not in seen:
                    seen.add(u.lower())
                    usernames.append(u)
            log.info("📡 TGStat «%s»: кандидатов %d", kw, len(usernames))
            for u in usernames:
                ent = await safe_call(self.client.get_entity, u)
                if ent is None:
                    continue
                if classify(ent):
                    await self.register(ent, source="tgstat")
                elif isinstance(ent, Channel) and ent.broadcast:
                    await self.try_linked_group(ent, source="tgstat")
                await asyncio.sleep(random.uniform(0.3, 0.8))
            await asyncio.sleep(1)
        log.info("📡 TGStat: добавлено групп %d", self.db.count_groups() - before)

    async def seed_from_usernames(self):
        log.info("🌱 Засев по списку известных чатов: %d шт.", len(SEED_USERNAMES))
        before = self.db.count_groups()
        for uname in SEED_USERNAMES:
            ent = await safe_call(self.client.get_entity, uname)
            if ent is None:
                continue
            if classify(ent):
                await self.register(ent, source="seed")
            elif isinstance(ent, Channel) and ent.broadcast:
                await self.try_linked_group(ent, source="seed")
            await asyncio.sleep(random.uniform(0.5, 1.2))
        log.info("🌱 По списку добавлено: %d", self.db.count_groups() - before)

    async def seed_from_patterns(self):
        cands = build_username_candidates(SEED_PATTERN_MAX)
        log.info("🧩 Перебор username-шаблонов (бесплатно): %d кандидатов", len(cands))
        before = self.db.count_groups()
        for u in cands:
            if self.db.count_groups() >= TARGET_GROUPS:
                break
            ent = await safe_call(self.client.get_entity, u)
            if ent is None:
                continue
            if classify(ent):
                await self.register(ent, source="pattern")
            elif isinstance(ent, Channel) and ent.broadcast:
                await self.try_linked_group(ent, source="pattern")
            await asyncio.sleep(random.uniform(0.4, 1.0))
        log.info("🧩 По шаблонам добавлено: %d", self.db.count_groups() - before)

    async def seed_from_dialogs(self):
        log.info("📂 Засев из твоих диалогов…")
        before = self.db.count_groups()
        try:
            async for dialog in self.client.iter_dialogs():
                ent = dialog.entity
                if classify(ent):
                    row = entity_to_row(ent, classify(ent), "seed")
                    if await self.db.upsert_group(row):
                        log.info("➕ [%s] %s (из диалогов)", row["type"], (row["title"] or "")[:40])
                    await self.db.enqueue(ent.id, row["username"])
        except Exception as e:
            log.debug("dialog seeding skipped: %s", e)
        log.info("📂 Из диалогов добавлено: %d", self.db.count_groups() - before)

    async def heartbeat(self):
        """Каждые 10 сек показывает, что процесс жив, даже если групп пока 0."""
        while True:
            await asyncio.sleep(10)
            log.info("💓 Жив: групп %d | в очереди %d",
                     self.db.count_groups(), self.db.count_new_queue())

    async def run(self):
        self.pbar = tqdm(total=TARGET_GROUPS, initial=self.db.count_groups(),
                         desc="Groups", unit="grp")
        log.info("🚀 Старт. Цель: %d групп. Воркеров: %d", TARGET_GROUPS, WORKERS)
        hb = asyncio.create_task(self.heartbeat())
        # ── Источники (бесплатные) ──
        await self.seed_from_dialogs()      # твои диалоги
        await self.seed_from_usernames()    # известные чаты
        await self.seed_from_patterns()     # перебор username-шаблонов
        await self.seed_search()            # глоб. поиск Telegram (города×темы)
        # ── Источник (платный, если задан TGSTAT_TOKEN) ──
        await self.seed_from_tgstat()
        log.info("🌐 Граф-расширение: запускаю %d параллельных воркеров "
                 "(в очереди %d групп)…", WORKERS, self.db.count_new_queue())
        workers = [asyncio.create_task(self.worker(i)) for i in range(WORKERS)]
        await asyncio.gather(*workers)
        hb.cancel()
        self.pbar.close()
        log.info("🏁 Сбор завершён. Итого групп: %d", self.db.count_groups())


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

async def main():
    db = DB(DB_PATH)
    client = TelegramClient(SESSION_NAME, API_ID, API_HASH)
    await client.start()
    log.info("Client authorized.")

    collector = Collector(client, db)
    try:
        await collector.run()
    finally:
        n = db.export()
        log.info("Exported %d groups -> %s / %s", n, CSV_PATH, XLSX_PATH)
        await client.disconnect()


if __name__ == "__main__":
    print("Telegram Russian Groups Collector")
    print("API_ID/API_HASH уже вшиты — нужен только вход по номеру телефона.\n")
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nОстановлено пользователем. Прогресс сохранён в базе.")
    input("\nГотово. Нажми Enter, чтобы закрыть окно…")
