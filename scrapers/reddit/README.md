# Reddit

`reddit_scraper.py` searches all of Reddit for a query, re-scores results locally for
relevance, and pulls comments for the threads it keeps. Output: Excel (or .csv/.json).

## Auth (required)
Reddit blocks anonymous requests. Log in to reddit.com in your browser, export cookies with
the **Cookie-Editor** extension (Export → JSON), and save the file as:

```
scrapers/reddit/cookies.json        # gitignored
```

The scraper also checks `~/.config/agent-reach/reddit_cookies.json`, `--cookies FILE`,
`$REDDIT_COOKIES_FILE`, and `$REDDIT_SESSION` (just the `reddit_session` cookie value).

## Run
```bash
.venv/bin/python reddit/reddit_scraper.py "motorola edge 70 fusion"
.venv/bin/python reddit/reddit_scraper.py "croma unboxed" "croma refurbished" --days 180 --max-posts 150
.venv/bin/python reddit/reddit_scraper.py "qualcomm x elite" --require india indian --output output/qc.xlsx
.venv/bin/python reddit/reddit_scraper.py --help
```

Default output: `reddit/output/reddit_relevance.xlsx`. Re-running with the same `--output`
resumes from the last checkpoint and merges.

## Indian infra job (2 datasets)
```bash
cd scrapers && python3 -m venv .venv && .venv/bin/pip install -r reddit/requirements.txt
# save Cookie-Editor JSON export as scrapers/reddit/cookies.json
bash reddit/jobs/run_infra.sh
```
Outputs land in `reddit/output/`. Re-run the same command to resume. To continue an earlier run, put its `*_checkpoint.xlsx` in `reddit/output/` first.
