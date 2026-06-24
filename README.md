# Telegram Russian Groups Collector (SaaS)

Production-ready system that collects **100 000+ Telegram groups, supergroups and
public chats** across Russia via Telethon — **excluding channels and bots**.

Ships as both a **CLI** and a **FastAPI SaaS service** (start/stop jobs, query,
export).

## Features

- **Type filtering** — keeps `group` / `supergroup`, drops `channel`, `bot`, `user`
- **Seed search** over 60+ Russian keywords + your existing dialogs
- **Graph expansion** — parses recent messages for `@mentions`, `t.me/` links and
  forwarded sources to grow the queue toward 100k
- **City detection** for 30+ Russian cities (Moscow, Saint Petersburg, Kazan, …)
- **SQLite** storage (`groups` + `queue` tables) with dedup
- **FloodWait handling**, retry logic, 1–3 s randomized delays
- **8 concurrent asyncio workers**
- Export to **CSV** and **XLSX** via pandas, progress via tqdm

## Architecture

```
app/
  config.py     # env-driven settings
  db.py         # SQLite persistence (groups, queue)
  collector.py  # Telethon engine: seed, classify, graph expansion
  api.py        # FastAPI SaaS endpoints
main.py         # CLI: one full collection pass + export
```

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env      # fill TG_API_ID / TG_API_HASH from https://my.telegram.org
```

## Run — CLI

```bash
python main.py
```

## Run — SaaS API

```bash
uvicorn app.api:app --host 0.0.0.0 --port 8000
# or
docker compose up --build
```

### Endpoints

| Method | Path            | Description                       |
|--------|-----------------|-----------------------------------|
| GET    | `/health`       | liveness                          |
| POST   | `/collect/start`| start background collection job   |
| POST   | `/collect/stop` | cancel running job                |
| GET    | `/stats`        | counts by type / city, queue size |
| GET    | `/groups`       | query groups (filter by city/type)|
| POST   | `/export`       | write CSV + XLSX                   |
| GET    | `/export/csv`   | download CSV                       |
| GET    | `/export/xlsx`  | download XLSX                      |

Set `API_KEY` in `.env` to protect endpoints with `Authorization: Bearer <key>`.

## Output

- `Telegram_groups_russia.csv`
- `Telegram_groups_russia.xlsx`

## Legal

Use only for collecting **public** group metadata and in compliance with
Telegram's Terms of Service and applicable data-protection law. You are
responsible for how you use the collected data.

## License

MIT — see [LICENSE](LICENSE).
