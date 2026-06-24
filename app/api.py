"""FastAPI SaaS layer: start collection jobs, query results, export."""

import asyncio
import logging

from fastapi import FastAPI, HTTPException, Depends, Header, Query
from fastapi.responses import FileResponse

from . import config
from .db import DB
from .collector import Collector, build_client

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s | %(levelname)s | %(message)s")

app = FastAPI(title="Telegram Russian Groups Collector", version="1.0.0")

_state = {"job": None, "running": False, "client": None}


def auth(authorization: str = Header(default="")):
    if config.API_KEY and authorization != f"Bearer {config.API_KEY}":
        raise HTTPException(status_code=401, detail="Unauthorized")


def get_db():
    return DB()


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/stats", dependencies=[Depends(auth)])
def stats(db: DB = Depends(get_db)):
    s = db.stats()
    s["job_running"] = _state["running"]
    return s


@app.post("/collect/start", dependencies=[Depends(auth)])
async def start():
    if _state["running"]:
        raise HTTPException(status_code=409, detail="Collection already running")

    async def _run():
        _state["running"] = True
        try:
            client = await build_client()
            _state["client"] = client
            db = DB()
            collector = Collector(client, db)
            await collector.run()
            db.export()
        except Exception as e:
            logging.exception("collection failed: %s", e)
        finally:
            _state["running"] = False
            if _state["client"]:
                await _state["client"].disconnect()
                _state["client"] = None

    _state["job"] = asyncio.create_task(_run())
    return {"status": "started"}


@app.post("/collect/stop", dependencies=[Depends(auth)])
async def stop():
    job = _state.get("job")
    if job and not job.done():
        job.cancel()
        _state["running"] = False
        return {"status": "stopping"}
    return {"status": "idle"}


@app.get("/groups", dependencies=[Depends(auth)])
def groups(
    db: DB = Depends(get_db),
    city: str = Query(default=None),
    type: str = Query(default=None),
    limit: int = Query(default=100, le=1000),
    offset: int = Query(default=0),
):
    q = "SELECT * FROM groups WHERE 1=1"
    params = []
    if city:
        q += " AND city = ?"
        params.append(city)
    if type:
        q += " AND type = ?"
        params.append(type)
    q += " ORDER BY participants_count DESC LIMIT ? OFFSET ?"
    params += [limit, offset]
    cur = db.conn.execute(q, params)
    cols = [c[0] for c in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]


@app.post("/export", dependencies=[Depends(auth)])
def export(db: DB = Depends(get_db)):
    n = db.export()
    return {"exported": n, "csv": config.CSV_PATH, "xlsx": config.XLSX_PATH}


@app.get("/export/csv", dependencies=[Depends(auth)])
def export_csv(db: DB = Depends(get_db)):
    db.export()
    return FileResponse(config.CSV_PATH, filename="Telegram_groups_russia.csv")


@app.get("/export/xlsx", dependencies=[Depends(auth)])
def export_xlsx(db: DB = Depends(get_db)):
    db.export()
    return FileResponse(config.XLSX_PATH, filename="Telegram_groups_russia.xlsx")
