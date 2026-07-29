# Asyncexchange

**Async Microsoft Exchange / Outlook library** — think [exchangelib](https://github.com/ecederstrand/exchangelib), but built on `async`/`await` and [httpx](https://www.python-httpx.org/).

Asyncexchange talks to Exchange via the EWS SOAP API. It covers a focused set of operations for mail and calendar. The API is designed to be extended; contributions and feature requests are welcome.

## Features

- **Email** — fetch inbox messages (filter by date / read status), mark as read
- **Calendar** — fetch meetings (busy time) from the user's calendar in a date range
- Async HTTP via `httpx`, typed models via Pydantic

## Requirements

- Python 3.13+
- Exchange server with EWS (Exchange Web Services) enabled

## Installation

```bash
pip install asyncexchange
```

## Quick start

```python
import asyncio
import datetime as dt

from asyncexchange.services.exchange.emails import EmailService

async def main():
    service = EmailService(
        username="user@example.com",
        password="secret",
        server_url="https://mail.example.com",
        tz=dt.timezone(dt.timedelta(hours=3)),
    )
    end = dt.datetime.now(tz=dt.timezone.utc)
    start = end - dt.timedelta(days=1)
    messages = await service.get_messages(start=start, end=end, is_read=False)
    for msg in messages:
        print(msg.subject, msg.author)
    await service.mark_as_read(messages)
    await service.aclose()

asyncio.run(main())
```

## Scope and contributing

Asyncexchange was built around a concrete set of use cases, so it does not aim to mirror the full Exchange API. If you need more operations (folders, send, etc.), open an issue or a PR.

## License

MIT License
