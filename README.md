# Korea NXT Scanner

Korean stock scanner for NXT-listed stocks using DART disclosures, Naver Finance quotes, simple news reason enrichment, and Telegram notifications.

## What it does

- Loads the current NXT universe first.
- Fetches same-day DART disclosures.
- Filters disclosure candidates by material keywords.
- Adds Naver Finance close price, change rate, volume, and trading value.
- Adds a simple news-based reason/source for stocks up 28% or more.
- Produces CSV, JSON, HTML, and XLSX reports as GitHub Actions artifacts.
- Sends Telegram summaries when Telegram secrets are configured.

## Required GitHub Secrets

Add these repository secrets:

- `DART_API_KEY`: required for the 06:00 daily DART scanner.
- `TELEGRAM_BOT_TOKEN`: optional but required for Telegram messages.
- `TELEGRAM_CHAT_ID`: optional but required for Telegram messages.

## Schedules

- `Daily NXT Disclosure Scanner`: every day at 06:00 KST. Telegram is sent as soon as the scan finishes, with enough buffer before 07:30 KST under normal GitHub Actions scheduling.
- `Morning NXT Upper Limit Telegram`: every day at 07:00 KST. It checks the latest available KRX listing data, filters stocks up at least 25%, keeps only NXT-listed stocks, and sends Telegram regardless of DART disclosure status.
- Manual runs start immediately and automatically use today's Korean date.

## Outputs

Daily scanner artifacts are retained for 30 days. Morning upper-limit checks also upload a JSON artifact for 30 days.
