# Quant Platform

Quant Platform is a Python-based quantitative research and stock screening web application. It combines data collection, market-data import, strategy backtesting, weekly reporting, user authentication, and a Flask dashboard for reviewing market signals.

The project is designed as an open-source foundation for research workflows, not as an investment-advice or live-trading system.

## Features

- Flask web dashboard with login, registration, admin user management, and upload pages
- Daily market data pipeline with crawler scripts and import tools
- SQLite-backed market data cache for faster dashboards and API responses
- Backtest and weekly report views with date navigation
- Screener APIs for filtering market records by return, industry, score, and other metrics
- Excel import, repair, validation, and export utilities
- Docker and docker-compose deployment files
- Automated tests for database integration, notifications, and atomic report writes

## Repository Layout

```text
quant_web/
  app.py                         Flask routes, auth, admin pages, APIs
  data_service.py                Data loading and transformation helpers
  db_service.py                  SQLite cache and market-data queries
  run_pipeline.py                Daily pipeline entry point
  batch_backtest.py              Backtest report generation
  batch_weekly.py                Weekly report generation
  import_market_xlsx.py          Market workbook import utility
  validate_workbook.py           Workbook validation utility
  templates/                     Jinja templates
  static/                        CSS, Bootstrap, ECharts, chart scripts
  Dockerfile
  docker-compose.yml

spider/
  quant_spider.py                Authenticated crawler
  playwright_spider.py           Playwright crawler variant
  upload_backtest.py             Backtest upload helper

tools/
  fix_excel.py                   Workbook repair helper
  recover_market.py              Market workbook recovery helper
```

## Requirements

- Python 3.11+
- SQLite
- Optional: Docker and Docker Compose
- Optional crawler/runtime dependencies listed in `quant_web/requirements.txt`

## Local Setup

```bash
cd quant_web
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy ..\.env.example .env
```

Edit `.env` before running the application. Do not commit `.env` or any real credentials.

```bash
set FLASK_SECRET_KEY=change-this-to-a-long-random-value
set ADMIN_PASSWORD=change-me
set USER_PASSWORD=change-me
set INVITE_CODE=change-me
python app.py
```

Then open `http://127.0.0.1:5000`.

## Docker

```bash
cd quant_web
copy ..\.env.example .env
docker compose up --build
```

The compose file reads secrets and runtime settings from environment variables. Keep production credentials outside Git.

## Data Pipeline

The pipeline can crawl market data, run backtests, generate weekly reports, refresh the SQLite cache, and optionally send an email notification.

```bash
cd quant_web
python run_pipeline.py
```

Useful utilities:

```bash
python import_market_xlsx.py path\to\market.xlsx
python validate_workbook.py path\to\workbook.xlsx
python batch_backtest.py
python batch_weekly.py
```

## Configuration

Copy `.env.example` to `.env` and fill in local values.

| Variable | Purpose |
| --- | --- |
| `FLASK_SECRET_KEY` | Session signing key. Use a long random value in production. |
| `ADMIN_PASSWORD` | Initial admin password created on first startup. |
| `USER_PASSWORD` | Initial regular-user password created on first startup. |
| `INVITE_CODE` | Registration invite code. |
| `CRAWLER_URL` | Target URL for crawler scripts. |
| `CRAWLER_USER` | Crawler username. |
| `CRAWLER_PASS` | Crawler password. |
| `SMTP_USER` | Email sender account. |
| `SMTP_PASS` | Email sender password or app password. |
| `SMTP_HOST` | SMTP host. |
| `SMTP_PORT` | SMTP SSL port. |
| `NOTIFY_EMAIL` | Optional notification recipient. |

## Tests

Run focused tests from the repository root:

```bash
python -m unittest discover -s quant_web -p "test_*.py"
```

## Security Notes

- No real API keys, passwords, session secrets, cookies, or private keys should be committed.
- `.env`, databases, generated market files, logs, and Excel workbooks are ignored by default.
- The application reads credentials from environment variables.
- The default runtime secret is suitable only for local development; production deployments must set `FLASK_SECRET_KEY`.

## Disclaimer

This project is for research, education, and software experimentation. It does not provide financial advice, investment recommendations, brokerage services, or live trading. Users are responsible for validating data, assumptions, and strategies before using any output.

## Contributing

Issues and pull requests are welcome. Good contributions include:

- Data-source adapters
- More robust workbook validation
- Backtest methodology improvements
- Dashboard and chart improvements
- Security hardening
- Tests for data pipeline edge cases

## License

MIT License. See `LICENSE`.
