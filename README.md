# Infra Stocks - YouTube Sentiment

Collects YouTube videos + top comments for 19 Indian infra stocks (2015-2026) and scores sentiment
into the `India_Infra_Budget_vs_Companies.xlsx` workbook.

## Setup
```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```
Be logged into YouTube in Chrome on this machine - the collector reads Chrome's YouTube cookies
(`cookiesfrombrowser`) so YouTube does not rate-limit it. Close Chrome if cookie reading fails.

## Run
```bash
# 1. collect (resumable - skips videos already in data/youtube_raw.json)
.venv/bin/python scripts/youtube_collect.py

# 2. score sentiment + add Sentiment Heatmap / by Year / Raw tabs to the workbook
.venv/bin/python scripts/sentiment_build.py
```

## Status
`data/youtube_raw.json` already holds UltraTech, JSW Steel, L&T, BKT, RVNL, IRB, NCC, KNR, PNC and HG Infra.
Rerunning step 1 continues from Ashoka Buildcon onwards (~10-12 min per company).

## Files
- `scripts/youtube_collect.py` - search per company per year (2015-2026), fetch metadata + up to 80 top comments
- `scripts/sentiment_build.py` - multilingual sentiment (cardiffnlp/twitter-xlm-roberta-base-sentiment-multilingual, English/Hindi/Hinglish)
- `data/youtube_raw.json` - collected data cache
- `India_Infra_Budget_vs_Companies.xlsx` - workbook the sentiment tabs are written into
