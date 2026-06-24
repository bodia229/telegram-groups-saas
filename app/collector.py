"""Telegram Russian groups collector engine."""

import re
import random
import asyncio
import logging
from datetime import datetime

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

from . import config
from .db import DB

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
# ENTITY CLASSIFICATION
# ──────────────────────────────────────────────────────────────────────────────

def classify(entity):
    """Return 'group'/'supergroup' for keepable entities, else None."""
    if isinstance(entity, Chat):
        if getattr(entity, "deactivated", False):
            return None
        return "group"
    if isinstance(entity, Channel):
        if entity.broadcast:                      # channel — excluded
            return None
        if entity.megagroup or getattr(entity, "gigagroup", False):
            return "supergroup"
        return None
    if isinstance(entity, User):                  # bots / users — excluded
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
# RATE-LIMIT / RETRY HELPERS
# ──────────────────────────────────────────────────────────────────────────────

async def jitter():
    await asyncio.sleep(random.uniform(config.MIN_DELAY, config.MAX_DELAY))


async def safe_call(coro_func, *args, **kwargs):
    for attempt in range(1, config.MAX_RETRIES + 1):
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
    def __init__(self, client, db, progress=None):
        self.client = client
        self.db = db
        self.progress = progress  # optional callback(new_count)

    async def register(self, entity, source, description=None):
        etype = classify(entity)
        if etype is None:
            return False
        row = entity_to_row(entity, etype, source, description)
        is_new = await self.db.upsert_group(row)
        await self.db.enqueue(entity.id, row["username"])
        if is_new and self.progress:
            self.progress(self.db.count_groups())
        return is_new

    async def seed_search(self):
        log.info("Seed search started (%d keywords)…", len(SEED_KEYWORDS))
        for kw in SEED_KEYWORDS:
            if self.db.count_groups() >= config.SEED_TARGET:
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
            async for msg in self.client.iter_messages(
                entity, limit=config.MESSAGES_PER_GROUP
            ):
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
            if self.db.count_groups() >= config.TARGET_GROUPS:
                break
            new_entity = await safe_call(self.client.get_entity, uname)
            if new_entity is not None:
                await self.register(new_entity, source="graph")
            await asyncio.sleep(random.uniform(0.3, 1.0))

        await self.db.mark_processed(group_id)
        await jitter()

    async def worker(self, name):
        while self.db.count_groups() < config.TARGET_GROUPS:
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
        await self.seed_from_dialogs()
        if self.db.count_groups() < config.SEED_TARGET:
            await self.seed_search()
        log.info("Graph expansion with %d workers…", config.WORKERS)
        workers = [asyncio.create_task(self.worker(i)) for i in range(config.WORKERS)]
        await asyncio.gather(*workers)


async def build_client():
    client = TelegramClient(config.SESSION_NAME, config.API_ID, config.API_HASH)
    await client.start()
    return client
