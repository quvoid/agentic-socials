#!/bin/bash
# Indian infrastructure stocks: all of Reddit, then Indian stock-market subreddits only.
cd "$(dirname "$0")/../.."
mapfile -t Q < reddit/jobs/infra_queries.txt
INDIA="india indian nse bse nifty sensex crore lakh rupee rupees inr ₹ budget"
.venv/bin/python -u reddit/reddit_scraper.py "${Q[@]}" --days 1830 --max-posts 150 --require $INDIA \
  --output reddit/output/infra_all_reddit.xlsx
.venv/bin/python -u reddit/reddit_scraper.py "${Q[@]}" --days 1830 --max-posts 150 \
  --subreddits IndianStockMarket IndiaInvestments IndianStreetBets DalalStreetTalks StockMarketIndia NSEbets \
  --output reddit/output/infra_stock_subreddits.xlsx
