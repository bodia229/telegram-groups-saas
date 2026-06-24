"""SQLite persistence layer (async-safe via a single lock)."""

import csv
import asyncio
import sqlite3
from datetime import datetime

import pandas as pd

from . import config


class DB:
    def __init__(self, path=None):
        self.path = path or config.DB_PATH
        self.conn = sqlite3.connect(self.path, check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL;")
        self.conn.execute("PRAGMA synchronous=NORMAL;")
        self._init_schema()
        self._lock = asyncio.Lock()

    def _init_schema(self):
        cur = self.conn.cursor()
        cur.execute(
            """
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
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS queue (
                group_id  INTEGER PRIMARY KEY,
                username  TEXT,
                status    TEXT DEFAULT 'new'
            )
            """
        )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_queue_status ON queue(status)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_groups_username ON groups(username)")
        self.conn.commit()

    async def upsert_group(self, g):
        """Insert/update a group. Return True if newly inserted."""
        async with self._lock:
            cur = self.conn.cursor()
            cur.execute("SELECT 1 FROM groups WHERE id = ?", (g["id"],))
            exists = cur.fetchone() is not None
            cur.execute(
                """
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
                """,
                g,
            )
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

    def stats(self):
        cur = self.conn.cursor()
        total = self.count_groups()
        by_type = dict(cur.execute("SELECT type, COUNT(*) FROM groups GROUP BY type").fetchall())
        by_city = dict(
            cur.execute(
                "SELECT city, COUNT(*) FROM groups WHERE city IS NOT NULL GROUP BY city ORDER BY 2 DESC"
            ).fetchall()
        )
        return {"total": total, "by_type": by_type, "by_city": by_city,
                "queue_new": self.count_new_queue()}

    def export(self):
        df = pd.read_sql_query(
            "SELECT * FROM groups ORDER BY participants_count DESC", self.conn
        )
        df.to_csv(config.CSV_PATH, index=False, encoding="utf-8-sig",
                  quoting=csv.QUOTE_MINIMAL)
        try:
            df.to_excel(config.XLSX_PATH, index=False, engine="openpyxl")
        except Exception:
            pass
        return len(df)
