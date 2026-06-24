# Telegram Russian Groups Collector

Production-ready system that collects **100 000+ Telegram groups, supergroups and
public chats** across Russia via Telethon — **excluding channels and bots**.

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
main.py         # CLI: one full collection pass + export
```

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env      # fill TG_API_ID / TG_API_HASH from https://my.telegram.org
```

## Run

```bash
python main.py
```

Or with Docker:

```bash
docker compose up --build
```

On first run Telethon will ask for your phone number and login code to create the
`*.session` file. Subsequent runs resume from the SQLite state automatically.

## Output

- `Telegram_groups_russia.csv`
- `Telegram_groups_russia.xlsx`

## Database schema

`groups`: `id, title, username, type, participants_count, description, city, source, created_at`
`queue`: `group_id, username, status (new / processing / processed)`

## Legal

Use only for collecting **public** group metadata and in compliance with
Telegram's Terms of Service and applicable data-protection law. You are
responsible for how you use the collected data.

## License

MIT — see [LICENSE](LICENSE).
