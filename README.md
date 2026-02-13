# Asyncexchange

**Async Microsoft Exchange / Outlook library** — think [exchangelib](https://github.com/ecederstrand/exchangelib), but built on `async`/`await` and [httpx](https://www.python-httpx.org/).

Asyncexchange talks to Exchange via the EWS SOAP API. It currently covers a focused set of operations (read mail, filter by date/read status, mark as read). The API is designed to be extended; contributions and feature requests are welcome.

## Requirements

- Python 3.11+
- Exchange server with EWS (Exchange Web Services) enabled

## Installation

```bash
pip install asyncexchange
```

## Scope and contributing

Asyncexchange was built around a concrete set of use cases, so it does not aim to mirror the full Exchange API. If you need more operations (folders, send, calendar, etc.), open an issue or a PR

## License

MIT License
