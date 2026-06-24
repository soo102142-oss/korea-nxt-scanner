# Korea NXT Scanner

Korean stock scanner for NXT-listed stocks using DART disclosures, Naver Finance quotes, and simple news reason enrichment.

## What it does

- Loads the current NXT universe first.
- Fetches same-day DART disclosures.
- Filters disclosure candidates by material keywords.
- Adds Naver Finance close price, change rate, and trading value.
- Adds a simple news-based reason/source for stocks up 28% or more.
- Produces CSV, JSON, HTML, and XLSX reports as GitHub Actions artifacts.

## Required GitHub Secret

Add this repository secret:

- `DART_API_KEY`

## Run

GitHub Actions runs automatically at 19:00 KST every day. You can also run it manually from the Actions tab.
