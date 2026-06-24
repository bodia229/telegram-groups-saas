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

### Windows 10 — один файл

Не нужна структура `app/` и установка зависимостей вручную. Просто:

1. Установи **Python 3.10+** (https://www.python.org/downloads/ — поставь галочку *Add to PATH*).
2. Запусти `telegram_groups_windows.py` двойным кликом, либо двойной клик по `run_windows.bat`,
   либо в консоли: `python telegram_groups_windows.py`.
3. При первом запуске введи `API_ID` / `API_HASH` (с https://my.telegram.org), затем телефон и код.

Зависимости (`telethon`, `pandas`, `tqdm`, `openpyxl`) ставятся автоматически.
Кодировка консоли и кириллица настраиваются сами. Результат — `Telegram_groups_russia.csv` / `.xlsx`.

Or with Docker:

```bash
docker compose up --build
```

On first run Telethon will ask for your phone number and login code to create the
`*.session` file. Subsequent runs resume from the SQLite state automatically.

## TGStat (опционально)

Глобальный поиск Telegram отдаёт мало результатов. Чтобы быстрее набрать стартовую
базу чатов по России, можно подключить **TGStat API**:

1. Получи токен в личном кабинете https://api.tgstat.ru (платный сервис).
2. Задай переменную окружения `TGSTAT_TOKEN` (или впиши в `.env`).

Тогда на старте отработает `seed_from_tgstat()`: возьмёт чаты из TGStat по ключам,
прогонит через фильтр (группы оставит, у каналов возьмёт чат-обсуждение) и
закинет в очередь графа. Без токена шаг просто пропускается.

> TGStat не вступает в чаты и не отдаёт сообщения — он даёт список юзернеймов,
> которые затем резолвит и расширяет Telethon.

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
