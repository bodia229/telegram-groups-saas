#!/usr/bin/env python3
"""CLI entrypoint — run a full collection pass and export."""

import asyncio
import logging

from tqdm import tqdm

from app import config
from app.db import DB
from app.collector import Collector, build_client

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler("collector.log", encoding="utf-8"),
              logging.StreamHandler()],
)
log = logging.getLogger("main")


async def main():
    if not config.API_ID or not config.API_HASH:
        raise SystemExit(
            "Set TG_API_ID and TG_API_HASH (get them at https://my.telegram.org)."
        )

    db = DB()
    client = await build_client()
    log.info("Client authorized.")

    pbar = tqdm(total=config.TARGET_GROUPS, initial=db.count_groups(),
                desc="Groups", unit="grp")

    def progress(n):
        pbar.n = n
        pbar.refresh()

    collector = Collector(client, db, progress=progress)
    try:
        await collector.run()
    finally:
        pbar.close()
        n = db.export()
        log.info("Exported %d groups → %s / %s", n, config.CSV_PATH, config.XLSX_PATH)
        await client.disconnect()


if __name__ == "__main__":
    asyncio.run(main())
