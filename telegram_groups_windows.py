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
import random
import asyncio
import logging
import sqlite3
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
# CONFIG  — впиши сюда свои данные ИЛИ оставь пусто (спросит при запуске)
# ──────────────────────────────────────────────────────────────────────────────

API_ID = int(os.getenv("TG_API_ID", "0"))        # напр. 123456
API_HASH = os.getenv("TG_API_HASH", "")          # напр. "abcd1234..."
SESSION_NAME = os.getenv("TG_SESSION", "ru_groups_session")

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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler("collector.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
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

async def jitter():
    await asyncio.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


async def safe_call(coro_func, *args, **kwargs):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return await coro_func(*args, **kwargs)
        except FloodWaitError as e:
            wait = e.seconds + random.uniform(1, 5)
            log.warning("FloodWait: sleeping %.0fs", wait)
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
        if is_new and self.pbar:
            self.pbar.n = self.db.count_groups()
            self.pbar.refresh()
        return is_new

    async def seed_search(self):
        log.info("Seed search started (%d keywords)…", len(SEED_KEYWORDS))
        for kw in SEED_KEYWORDS:
            if self.db.count_groups() >= SEED_TARGET:
                break
            result = await safe_call(self.client, SearchRequest(q=kw, limit=100))
            if not result:
                await jitter()
                continue
            for chat in result.chats:
                await self.register(chat, source="seed")
            await jitter()
        log.info("Seed search done. Collected: %d", self.db.count_groups())

    async def process_group(self, group_id, username):
        ref = username or group_id
        entity = await safe_call(self.client.get_entity, ref)
        if entity is None or classify(entity) is None:
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
        try:
            async for msg in self.client.iter_messages(entity, limit=MESSAGES_PER_GROUP):
                if msg.message:
                    for m in USERNAME_RE.findall(msg.message):
                        found_usernames.add(m.lower())
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
            await asyncio.sleep(e.seconds + 3)
        except Exception as e:
            log.debug("iter_messages failed for %s: %s", ref, e)

        for uname in found_usernames:
            if self.db.count_groups() >= TARGET_GROUPS:
                break
            new_entity = await safe_call(self.client.get_entity, uname)
            if new_entity is not None:
                await self.register(new_entity, source="graph")
            await asyncio.sleep(random.uniform(0.3, 1.0))

        await self.db.mark_processed(group_id)
        await jitter()

    async def worker(self, name):
        while self.db.count_groups() < TARGET_GROUPS:
            batch = await self.db.next_batch(1)
            if not batch:
                await asyncio.sleep(5)
                if self.db.count_new_queue() == 0:
                    return
                continue
            group_id, username = batch[0]
            try:
                await self.process_group(group_id, username)
            except Exception as e:
                log.debug("worker %d error: %s", name, e)
                await self.db.mark_processed(group_id)

    async def seed_from_dialogs(self):
        try:
            async for dialog in self.client.iter_dialogs():
                ent = dialog.entity
                if classify(ent):
                    row = entity_to_row(ent, classify(ent), "seed")
                    await self.db.upsert_group(row)
                    await self.db.enqueue(ent.id, row["username"])
        except Exception as e:
            log.debug("dialog seeding skipped: %s", e)

    async def run(self):
        self.pbar = tqdm(total=TARGET_GROUPS, initial=self.db.count_groups(),
                         desc="Groups", unit="grp")
        await self.seed_from_dialogs()
        if self.db.count_groups() < SEED_TARGET:
            await self.seed_search()
        log.info("Graph expansion with %d workers…", WORKERS)
        workers = [asyncio.create_task(self.worker(i)) for i in range(WORKERS)]
        await asyncio.gather(*workers)
        self.pbar.close()


# ──────────────────────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────────────────────

def _prompt_credentials():
    global API_ID, API_HASH
    if not API_ID:
        try:
            API_ID = int(input("Введите API_ID (с https://my.telegram.org): ").strip())
        except ValueError:
            raise SystemExit("API_ID должен быть числом.")
    if not API_HASH:
        API_HASH = input("Введите API_HASH: ").strip()
    if not API_ID or not API_HASH:
        raise SystemExit("API_ID и API_HASH обязательны.")


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
    _prompt_credentials()
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nОстановлено пользователем. Прогресс сохранён в базе.")
    input("\nГотово. Нажми Enter, чтобы закрыть окно…")
