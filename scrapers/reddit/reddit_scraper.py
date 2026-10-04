"""
Reddit query scraper — give it a search query, get the most relevant Reddit threads.

No subreddit lists. Every query is searched across all of Reddit, then the results are
re-scored locally so only threads that are actually about the query are kept.

How it works
  1. Collect  — each query is run as several search variants (exact phrase, all words,
                title-only) across several sort orders. Reddit caps every listing at ~250
                results, so the variants together reach far more of the relevant posts
                than a single search can. Duplicates are merged.
  2. Score    — every candidate gets a relevance score from where the query appears
                (exact phrase in the title beats scattered words in the body, etc.).
  3. Rank     — candidates are sorted best-first; --max-posts caps how many are enriched.
  4. Enrich   — comments are fetched for each kept thread (with "load more" expansion).
                Threads that only mention the query in their comments keep just those
                comments.

Reliability
  - honours Reddit's x-ratelimit-* headers and Retry-After, backs off on 429/5xx/network errors
  - falls back to old.reddit.com when www returns 403
  - checkpoints every N threads; re-running with the same --output resumes where it stopped

Auth (Reddit blocks anonymous .json requests): export your reddit.com cookies with the
Cookie-Editor browser extension and save them to scrapers/reddit/cookies.json or
~/.config/agent-reach/reddit_cookies.json (or pass --cookies FILE, or set REDDIT_SESSION). Cookie-Editor's JSON export, a plain
{"name": "value"} dict and rdt-cli's credential.json all work.

Usage
  python reddit_scraper.py "motorola edge 70 fusion"
  python reddit_scraper.py "croma unboxed" "croma refurbished" --days 180 --max-posts 150
  python reddit_scraper.py "qualcomm x elite" --require india indian --output qc.xlsx
  python reddit_scraper.py "pixel 10 battery" --depth deep --min-score 4 --no-comments
"""

import argparse
import json
import math
import os
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
import requests

import os as _os
HERE = _os.path.dirname(_os.path.abspath(__file__))
OUTPUT_DIR = _os.path.join(HERE, "output")
DATA_DIR = _os.path.join(HERE, "data")
_os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------------------------------------------------------- config / defaults
DAYS_BACK = 366
OUTPUT_EXCEL = _os.path.join(OUTPUT_DIR, "reddit_relevance.xlsx")
BATCH_SAVE_EVERY = 25          # checkpoint after every N enriched threads
MAX_MORE_COMMENT_IDS = 200     # how many hidden ("load more") comments to expand per thread
PROXIMITY_WINDOW = 12         # without the exact phrase, all query words must sit within this many words
MIN_REQUEST_GAP = 1.0          # seconds between requests, even when the rate limit allows more

COOKIE_LOCATIONS = [
    _os.path.join(HERE, "cookies.json"),   # scrapers/reddit/cookies.json (gitignored)
    "~/.config/agent-reach/reddit_cookies.json",
    "~/.config/rdt-cli/credential.json",
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.reddit.com/",
}

# How many search variants to run per query. Each (variant, sort) pair is one listing of up
# to ~250 posts. More listings = better recall, more requests.
DEPTHS = {
    "quick":  [("phrase", "relevance"), ("words", "relevance")],
    "normal": [("phrase", "relevance"), ("phrase", "top"), ("phrase", "new"),
               ("words", "relevance"), ("words", "top"), ("title", "relevance")],
    "deep":   [("phrase", "relevance"), ("phrase", "top"), ("phrase", "new"), ("phrase", "comments"),
               ("words", "relevance"), ("words", "top"), ("words", "new"), ("words", "comments"),
               ("title", "relevance"), ("title", "new")],
}

STOPWORDS = {
    "a", "an", "the", "and", "or", "of", "for", "to", "in", "on", "at", "by", "with", "is",
    "are", "was", "be", "vs", "v", "my", "me", "i", "it", "this", "that", "how", "what",
    "which", "should", "do", "does", "can", "from", "about",
}

# Post-scrape topic buckets (for the analysis step, NOT used by the scraper).
# Apply these to post_title + comment_body to tag each row with topics.
TOPIC_KEYWORDS = {
    # --- Croma Retail ---
    "croma_pricing":        ["price", "expensive", "cheaper", "discount", "offer", "deal", "mrp"],
    "croma_emi_finance":    ["emi", "no cost emi", "bajaj", "credit card", "cashback"],
    "croma_warranty":       ["warranty", "extended warranty", "protection plan", "zip care", "zipcare"],
    "croma_returns_refund": ["return", "refund", "replacement", "exchange", "dead on arrival", "doa"],
    "croma_store_exp":      ["store", "showroom", "staff", "salesman", "sales guy", "in-store", "instore"],
    "croma_delivery":       ["delivery", "delivered", "installation", "shipping", "late"],
    "croma_vs_online":      ["amazon", "flipkart", "online", "reliance digital", "vijay sales"],
    "croma_apple":          ["iphone", "macbook", "ipad", "airpods", "apple"],
    "croma_tv_appliance":   ["tv", "television", "oled", "fridge", "refrigerator", "washing machine", "ac ", "air conditioner"],
    # --- Croma Unboxed ---
    "unboxed_trust":        ["refurbished", "refurb", "open box", "openbox", "unboxed", "second hand", "used", "renewed"],
    "unboxed_condition":    ["scratch", "dent", "battery health", "condition", "grade", "like new"],
    "unboxed_warranty":     ["warranty", "guarantee", "return policy"],
    "unboxed_competitors":  ["cashify", "amazon renewed", "flipkart refurbished", "2gud", "olx"],
    # --- Qualcomm ---
    "qc_performance":       ["performance", "benchmark", "antutu", "geekbench", "fps", "gaming", "throttle", "throttling"],
    "qc_heating_battery":   ["heat", "heating", "overheat", "thermal", "battery", "efficiency", "power draw"],
    "qc_vs_mediatek":       ["dimensity", "mediatek", "exynos", "tensor", "bionic", "apple silicon"],
    "qc_laptops_pc":        ["x elite", "x plus", "windows on arm", "arm laptop", "copilot+", "surface"],
    "qc_phones":            ["oneplus", "samsung", "xiaomi", "realme", "iqoo", "poco", "nothing phone", "motorola"],
    "qc_price_value":       ["price", "expensive", "overpriced", "value", "budget", "mid-range", "midrange", "flagship"],
    "qc_launch_news":       ["launch", "announced", "summit", "unveiled", "leak", "rumour", "rumor"],
    "qc_ai_features":       ["ai", "npu", "on-device", "gen ai", "copilot"],
    "qc_modem_connectivity":["5g", "modem", "wifi 7", "bluetooth", "x75", "x80"],
    "qc_brand_community":   ["insiders", "snapdragon insiders", "ama", "giveaway"],
    "qc_careers_hiring":    ["offer", "salary", "package", "interview", "hiring", "referral", "refferal", "intern", "internship", "oa", "wlb", "work life", "layoff", "layoffs", "wfh", "promo", "promotion", "switch", "join", "role", "fresher", "ctc"],
    "qc_india_business":    ["qualcomm india", "india labs", "tape-out", "tape out", "deep tech", "investment", "invest", "startup", "make in india", "semiconductor", "6g", "payments device", "adas", "hyderabad", "bangalore", "bengaluru", "chennai", "noida"],
}


def ts(t) -> str:
    return datetime.fromtimestamp(t, timezone.utc).strftime("%Y-%m-%d %H:%M:%S") if t else ""


# ---------------------------------------------------------------- cookies
class AuthError(RuntimeError):
    pass


def parse_cookie_file(path: Path) -> dict:
    """Accepts Cookie-Editor JSON export, a {name: value} dict, or rdt-cli credential.json."""
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, list):                                   # Cookie-Editor export
        return {c["name"]: c["value"] for c in data
                if isinstance(c, dict) and "name" in c and "reddit" in c.get("domain", "reddit")}
    if isinstance(data, dict) and isinstance(data.get("cookies"), dict):   # rdt-cli
        return {k: v for k, v in data["cookies"].items() if v}
    if isinstance(data, dict):
        return {k: v for k, v in data.items() if isinstance(v, str)}
    return {}


def load_cookies(explicit: str | None) -> tuple[dict, str]:
    candidates = [explicit] if explicit else [os.environ.get("REDDIT_COOKIES_FILE"), *COOKIE_LOCATIONS]
    for c in candidates:
        if not c:
            continue
        p = Path(c).expanduser()
        if p.is_file():
            cookies = parse_cookie_file(p)
            if cookies.get("reddit_session"):
                return cookies, str(p)
        elif explicit:
            raise AuthError(f"cookie file not found: {p}")
    if os.environ.get("REDDIT_SESSION"):
        return {"reddit_session": os.environ["REDDIT_SESSION"]}, "$REDDIT_SESSION"
    return {}, ""


# Kept for scripts that import COOKIES from this module (e.g. qualcomm_reddit_size.py).
try:
    COOKIES, _ = load_cookies(None)
except Exception:
    COOKIES = {}


# ---------------------------------------------------------------- HTTP client
class RedditClient:
    """GETs Reddit .json endpoints with rate-limit pacing, retries and host fallback."""

    HOSTS = ["https://www.reddit.com", "https://old.reddit.com"]

    def __init__(self, cookies: dict, max_retries: int = 6):
        self.s = requests.Session()
        self.s.headers.update(HEADERS)
        self.s.cookies.update(cookies)
        self.max_retries = max_retries
        self.host_idx = 0
        self.last_request = 0.0
        self.pause_until = 0.0
        self.requests_made = 0

    @property
    def host(self) -> str:
        return self.HOSTS[self.host_idx]

    def _pace(self):
        now = time.time()
        wait = max(self.pause_until - now, self.last_request + MIN_REQUEST_GAP - now)
        if wait > 0:
            if wait > 5:
                print(f"    ⏳ rate-limit pause {wait:.0f}s")
            time.sleep(wait)

    def _read_ratelimit(self, r: requests.Response):
        try:
            remaining = float(r.headers.get("x-ratelimit-remaining", "nan"))
            reset = float(r.headers.get("x-ratelimit-reset", "nan"))
        except ValueError:
            return
        if not math.isnan(remaining) and not math.isnan(reset) and remaining < 3:
            self.pause_until = time.time() + reset + 1

    def get(self, path: str, params: dict | None = None):
        """path is either '/search.json' style (host added) or a full reddit URL."""
        backoff = 5.0
        tried_fallback = False
        for attempt in range(self.max_retries):
            url = path if path.startswith("http") else self.host + path
            self._pace()
            try:
                r = self.s.get(url, params=params, timeout=25)
            except requests.RequestException as e:
                print(f"    ⚠️ network error ({e.__class__.__name__}), retry in {backoff:.0f}s")
                time.sleep(backoff)
                backoff = min(backoff * 2, 120)
                continue
            finally:
                self.last_request = time.time()
                self.requests_made += 1

            self._read_ratelimit(r)
            if r.status_code == 200:
                try:
                    return r.json()
                except ValueError:
                    # HTML interstitial instead of JSON — treat like a block
                    pass
            if r.status_code == 429:
                wait = float(r.headers.get("retry-after") or r.headers.get("x-ratelimit-reset") or backoff)
                print(f"    ⏳ 429 rate-limited, sleeping {wait:.0f}s")
                time.sleep(wait + 1)
                backoff = min(backoff * 2, 120)
                continue
            if r.status_code in (403, 200) and not tried_fallback and not path.startswith("http"):
                tried_fallback = True
                self.host_idx = (self.host_idx + 1) % len(self.HOSTS)
                print(f"    ↪️ blocked (HTTP {r.status_code}), switching to {self.host}")
                continue
            if r.status_code in (401, 403):
                raise AuthError(f"HTTP {r.status_code} from Reddit — cookies missing or expired")
            if r.status_code == 404:
                return None
            if r.status_code >= 500 or r.status_code == 200:
                print(f"    ⚠️ HTTP {r.status_code}, retry in {backoff:.0f}s")
                time.sleep(backoff)
                backoff = min(backoff * 2, 120)
                continue
            print(f"    ⚠️ HTTP {r.status_code} for {url}")
            return None
        print(f"    ❌ gave up after {self.max_retries} attempts: {path}")
        return None

    def whoami(self) -> str | None:
        data = self.get("/api/me.json")
        return (data or {}).get("data", {}).get("name") if isinstance(data, dict) else None


# ---------------------------------------------------------------- relevance
@dataclass
class Query:
    text: str
    terms: list[str] = field(default_factory=list)

    def __post_init__(self):
        words = re.findall(r"[\w+#.&-]+", self.text.lower())
        words = [w.strip(".-") for w in words]
        self.terms = [w for w in words if w and w not in STOPWORDS] or [w for w in words if w]


def _term_re(term: str) -> re.Pattern:
    # Word-start match that tolerates simple suffixes: "review" matches "reviews", "reviewed".
    # Short terms / model numbers ("5g", "x80", "70") must match as whole words.
    tail = r"(?!\w)" if len(term) < 4 or any(ch.isdigit() for ch in term) else r"\w{0,3}(?!\w)"
    return re.compile(r"(?<!\w)" + re.escape(term) + tail, re.I)


def _phrase_re(text: str) -> re.Pattern:
    parts = [re.escape(p) for p in text.lower().split()]
    return re.compile(r"(?<!\w)" + r"[\s\-_]*".join(parts) + r"(?!\w)", re.I)


class Matcher:
    def __init__(self, q: Query):
        self.q = q
        self.phrase = _phrase_re(q.text)
        self.term_res = {t: _term_re(t) for t in q.terms}

    def terms_in(self, text: str) -> set[str]:
        return {t for t, rx in self.term_res.items() if rx.search(text or "")}

    def has_phrase(self, text: str) -> bool:
        return bool(self.phrase.search(text or ""))

    def near(self, text: str, window: int = PROXIMITY_WINDOW) -> bool:
        """Every term occurs inside some span of at most `window` words."""
        if len(self.q.terms) < 2:
            return bool(self.terms_in(text))
        hits = []
        for i, w in enumerate(re.findall(r"[\w+#.&-]+", text or "")):
            for t, rx in self.term_res.items():
                if rx.fullmatch(w.strip(".-")) or rx.match(w):
                    hits.append((i, t))
        need, have, lo = len(self.q.terms), {}, 0
        for hi in range(len(hits)):
            have[hits[hi][1]] = have.get(hits[hi][1], 0) + 1
            while len(have) == need:
                if hits[hi][0] - hits[lo][0] < window:
                    return True
                t = hits[lo][1]
                have[t] -= 1
                if not have[t]:
                    del have[t]
                lo += 1
        return False

    def mentions(self, text: str) -> bool:
        """Text is about the query: exact phrase, or every term close together."""
        return self.has_phrase(text) or self.near(text)

    def score_post(self, title: str, body: str) -> tuple[float, str]:
        n = len(self.q.terms)
        t_terms, b_terms = self.terms_in(title), self.terms_in(body)
        score = 0.0
        if self.has_phrase(title):
            score += 6
        elif self.has_phrase(body):
            score += 3
        score += 3 * len(t_terms) / n + 1.5 * len(b_terms) / n
        covered = len(t_terms | b_terms) / n
        score *= 0.5 + 0.5 * covered           # partial matches are pushed down hard
        if self.has_phrase(title) or t_terms == set(self.q.terms):
            where = "title"
        elif self.mentions(title + "\n" + body):
            where = "body"
            score += 1                           # words sit together even without the exact phrase
        else:
            where = ""
            score *= 0.5                         # terms scattered far apart: likely a different topic
        return round(score, 2), where


def engagement_boost(post: dict) -> float:
    return math.log10(1 + max(post.get("score") or 0, 0)) * 0.3 + math.log10(1 + (post.get("num_comments") or 0)) * 0.3


# ---------------------------------------------------------------- search
def time_window(days: int) -> str:
    for limit, window in ((1, "day"), (7, "week"), (31, "month"), (366, "year")):
        if days <= limit:
            return window
    return "all"


def build_q(q: Query, variant: str) -> str:
    if variant == "phrase":
        return f'"{q.text}"'
    if variant == "title":
        return f'title:"{q.text}"' if len(q.terms) > 1 else f"title:{q.text}"
    return q.text


@dataclass
class Candidate:
    post: dict
    query: str
    variants: set[str] = field(default_factory=set)
    best_rank: int = 10**6
    score: float = 0.0
    matched_in: str = ""


def collect(client: RedditClient, q: Query, plan: list, days: int, min_ts: int,
            max_pages: int, include_nsfw: bool, subreddits: list[str] | None = None) -> dict[str, Candidate]:
    found: dict[str, Candidate] = {}
    variants = plan if len(q.terms) > 1 else [(v, s) for v, s in plan if v != "words"]
    for variant, sort in variants:
        qs = build_q(q, variant)
        # "new" is already date ordered; everything else gets a coarse Reddit time window
        params = {"q": qs, "sort": sort, "t": time_window(days), "type": "link",
                  "limit": 100, "raw_json": 1, "include_over_18": "on" if include_nsfw else "off"}
        path = "/search.json"
        if subreddits:
            path = f"/r/{'+'.join(subreddits)}/search.json"
            params["restrict_sr"] = 1
        after, rank, got, fresh = None, 0, 0, 0
        for _ in range(max_pages):
            if after:
                params["after"] = after
            data = client.get(path, params)
            children = (data or {}).get("data", {}).get("children", []) if isinstance(data, dict) else []
            if not children:
                break
            stop = False
            for ch in children:
                rank += 1
                p = ch.get("data", {})
                pid = p.get("id")
                if not pid:
                    continue
                if (p.get("created_utc") or 0) < min_ts:
                    if sort == "new":
                        stop = True        # everything after this is older
                    continue
                if p.get("over_18") and not include_nsfw:
                    continue
                got += 1
                c = found.get(pid)
                if c is None:
                    c = found[pid] = Candidate(post=p, query=q.text)
                    fresh += 1
                c.variants.add(f"{variant}/{sort}")
                c.best_rank = min(c.best_rank, rank)
            after = data.get("data", {}).get("after")
            if stop or not after:
                break
        print(f"  🔎 {qs:<40} sort={sort:<9} {got:>4} hits, {fresh:>4} new  (total {len(found)})")
    return found


# ---------------------------------------------------------------- comments
def fetch_comments(client: RedditClient, permalink: str, post_id: str, max_more: int) -> list[dict]:
    data = client.get(f"{permalink.rstrip('/')}.json",
                      {"limit": 500, "depth": 10, "raw_json": 1, "sort": "top"})
    out: list[dict] = []
    more_ids: list[str] = []
    if isinstance(data, list) and len(data) > 1:
        _walk(data[1], out, more_ids)
    more_ids = more_ids[:max_more]
    for i in range(0, len(more_ids), 100):
        chunk = more_ids[i:i + 100]
        resp = client.get("/api/morechildren.json", {
            "link_id": f"t3_{post_id}", "children": ",".join(chunk),
            "api_type": "json", "raw_json": 1, "limit_children": "false",
        })
        things = (resp or {}).get("json", {}).get("data", {}).get("things", []) if isinstance(resp, dict) else []
        for t in things:
            _walk(t, out, [])       # nested "more" inside a "more" is not expanded again
    seen, uniq = set(), []
    for c in out:
        if c["comment_id"] not in seen:
            seen.add(c["comment_id"])
            uniq.append(c)
    return uniq


def _walk(node, out: list, more_ids: list):
    if isinstance(node, list):
        for n in node:
            _walk(n, out, more_ids)
        return
    if not isinstance(node, dict):
        return
    kind = node.get("kind")
    if kind == "Listing":
        for ch in node.get("data", {}).get("children", []):
            _walk(ch, out, more_ids)
    elif kind == "t1":
        d = node.get("data", {})
        body = d.get("body") or ""
        if body not in ("[deleted]", "[removed]"):
            out.append({
                "comment_id": d.get("id"),
                "parent_id": (d.get("parent_id") or "").split("_", 1)[-1],
                "author": d.get("author", ""),
                "body": body,
                "score": d.get("score"),
                "created_utc": ts(d.get("created_utc", 0)),
            })
        if d.get("replies"):
            _walk(d["replies"], out, more_ids)
    elif kind == "more":
        more_ids.extend(node.get("data", {}).get("children", []))


# ---------------------------------------------------------------- storage
def read_existing(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        if path.endswith(".csv"):
            return pd.read_csv(path)
        if path.endswith(".json"):
            return pd.read_json(path)
        return pd.read_excel(path, sheet_name=0)
    except Exception as e:
        print(f"    ⚠️ could not read {path}: {e}")
        return pd.DataFrame()


def write_rows(df: pd.DataFrame, path: str):
    base, ext = os.path.splitext(path)
    tmp = f"{base}.tmp{ext}"
    if ext == ".csv":
        df.to_csv(tmp, index=False)
    elif ext == ".json":
        df.to_json(tmp, orient="records", force_ascii=False, indent=1)
    else:
        with pd.ExcelWriter(tmp) as xw:
            df.to_excel(xw, sheet_name="rows", index=False)
            post_cols = [c for c in df.columns if c.startswith("post_") or c in (
                "subreddit", "search_query", "relevance_score", "matched_in", "search_variants")]
            posts = (df[post_cols].drop_duplicates("post_id")
                     .sort_values(["search_query", "relevance_score"], ascending=[True, False]))
            posts.to_excel(xw, sheet_name="posts", index=False)
    os.replace(tmp, path)


def merge(existing: pd.DataFrame, rows: list[dict]) -> pd.DataFrame:
    df = pd.concat([existing, pd.DataFrame(rows)], ignore_index=True)
    df["comment_id"] = df["comment_id"].astype(str)
    return df.drop_duplicates(subset=["post_id", "comment_id"], keep="last")


# ---------------------------------------------------------------- main pipeline
def scrape(queries: list[str], *, days: int = DAYS_BACK, depth: str = "normal", max_pages: int = 3,
           max_posts: int = 200, min_score: float = 2.0, require: list[str] | None = None,
           comments: str = "auto", include_nsfw: bool = False, output: str = OUTPUT_EXCEL,
           cookies_path: str | None = None, subreddits: list[str] | None = None) -> pd.DataFrame:
    cookies, source = load_cookies(cookies_path)
    if not cookies:
        raise AuthError("no Reddit cookies found. Export them with Cookie-Editor to "
                        "scrapers/reddit/cookies.json or ~/.config/agent-reach/reddit_cookies.json (or use --cookies / REDDIT_SESSION).")
    client = RedditClient(cookies)
    user = client.whoami()
    print(f"👤 {'logged in as u/' + user if user else 'cookies loaded but not logged in'} (cookies: {source})")

    min_ts = int((datetime.now(timezone.utc) - timedelta(days=days)).timestamp())
    require_res = [_phrase_re(w) for w in (require or [])]
    checkpoint = os.path.splitext(output)[0] + "_checkpoint" + os.path.splitext(output)[1]

    # Resume: threads already enriched in a previous (interrupted) run are skipped.
    done_df = read_existing(checkpoint)
    done = set(zip(done_df["search_query"], done_df["post_id"])) if {"search_query", "post_id"} <= set(done_df.columns) else set()
    if done:
        print(f"♻️  resuming — {len(done)} threads already in {checkpoint}")
    rows: list[dict] = []
    since_save = 0

    def save_checkpoint():
        write_rows(merge(done_df, rows), checkpoint)
        print(f"    💾 checkpoint -> {checkpoint}")

    for qtext in queries:
        q = Query(qtext)
        m = Matcher(q)
        print(f"\n━━━ \"{q.text}\"  terms={q.terms}  depth={depth}  last {days}d")
        found = collect(client, q, DEPTHS[depth], days, min_ts, max_pages, include_nsfw, subreddits)

        # score + rank
        ranked: list[Candidate] = []
        for c in found.values():
            s, where = m.score_post(c.post.get("title") or "", c.post.get("selftext") or "")
            c.score = round(s + engagement_boost(c.post) + len(c.variants) * 0.2, 2)
            c.matched_in = where
            ranked.append(c)
        ranked.sort(key=lambda c: (c.score, -c.best_rank), reverse=True)

        strong = [c for c in ranked if c.matched_in and c.score >= min_score]
        weak = [c for c in ranked if not c.matched_in]      # query may still be in comments
        print(f"  📊 {len(found)} candidates → {len(strong)} match title/body, {len(weak)} need a comment check")

        to_enrich = strong[:max_posts]
        if comments == "auto":
            to_enrich += weak[:max(0, max_posts - len(to_enrich))]

        kept = 0
        for i, c in enumerate(to_enrich, 1):
            p = c.post
            pid = p["id"]
            if (q.text, pid) in done:
                kept += 1
                continue
            title, body = p.get("title") or "", p.get("selftext") or ""
            permalink = p.get("permalink", "")
            thread_comments = []
            if comments != "none":
                try:
                    thread_comments = fetch_comments(client, permalink, pid, MAX_MORE_COMMENT_IDS)
                except AuthError:
                    raise
                except Exception as e:
                    print(f"    ⚠️ comments failed for {pid}: {e}")

            if c.matched_in:
                keep = thread_comments
            else:
                keep = [cm for cm in thread_comments if m.mentions(cm["body"])]
                if not keep:
                    continue
                c.matched_in = "comments"

            if require_res:
                text = "\n".join([title, body, p.get("subreddit", "")] + [cm["body"] for cm in keep])
                if not any(rx.search(text) for rx in require_res):
                    continue

            kept += 1
            print(f"  📄 {i:>3}. [{c.score:>5}] r/{p.get('subreddit', ''):<20} {title[:70]}"
                  f"  ({len(keep)} comments, via {c.matched_in})")
            base = {
                "search_query": q.text,
                "search_phrase": q.text,                 # old column name, kept for existing reports
                "relevance_score": c.score,
                "matched_in": c.matched_in,
                "search_variants": ", ".join(sorted(c.variants)),
                "relevance_rank": c.best_rank,
                "post_id": pid,
                "subreddit": p.get("subreddit", ""),
                "post_title": title,
                "post_body": body,
                "post_author": p.get("author", ""),
                "post_created_utc": ts(p.get("created_utc", 0)),
                "post_url": f"https://www.reddit.com{permalink}",
                "post_link": p.get("url", ""),
                "post_score": p.get("score"),
                "post_upvote_ratio": p.get("upvote_ratio"),
                "post_num_comments": p.get("num_comments"),
                "post_flair": p.get("link_flair_text") or "",
                "phrase_in_title": "; ".join(sorted(m.terms_in(title))),
                "phrase_in_body": "; ".join(sorted(m.terms_in(body))),
            }
            if keep:
                for cm in keep:
                    rows.append({**base,
                                 "comment_id": cm["comment_id"], "comment_parent_id": cm["parent_id"],
                                 "comment_author": cm["author"], "comment_body": cm["body"],
                                 "comment_score": cm["score"], "comment_created_utc": cm["created_utc"],
                                 "phrase_in_comment": "; ".join(sorted(m.terms_in(cm["body"])))})
            else:
                rows.append({**base, "comment_id": "No Comments", "comment_parent_id": "",
                             "comment_author": "", "comment_body": "", "comment_score": "",
                             "comment_created_utc": "", "phrase_in_comment": ""})

            since_save += 1
            if since_save >= BATCH_SAVE_EVERY:
                save_checkpoint()
                since_save = 0
        print(f"  ✅ \"{q.text}\": {kept} relevant threads kept")
        if rows:
            save_checkpoint()
            since_save = 0

    print(f"\n🌐 {client.requests_made} requests made")
    new = merge(done_df, rows) if (rows or not done_df.empty) else pd.DataFrame()
    if new.empty:
        print("No relevant threads found.")
        return new
    final = merge(read_existing(output), new.to_dict("records"))
    write_rows(final, output)
    if os.path.exists(checkpoint):
        os.remove(checkpoint)
    print(f"✅ {new['post_id'].nunique()} threads / {len(new)} rows this run → {output} ({len(final)} rows total)")
    return final


def main():
    ap = argparse.ArgumentParser(description="Search all of Reddit for a query and keep the most relevant threads.")
    ap.add_argument("queries", nargs="*", help='search queries, e.g. "motorola edge 70 fusion"')
    ap.add_argument("--phrases", nargs="+", help="alias for the positional queries (old flag)")
    ap.add_argument("--days", type=int, default=DAYS_BACK, help="only posts from the last N days (default 366)")
    ap.add_argument("--depth", choices=list(DEPTHS), default="normal",
                    help="how many search variants per query: quick | normal | deep (default normal)")
    ap.add_argument("--max-pages", type=int, default=3, help="pages per search variant, 100 posts each (default 3)")
    ap.add_argument("--max-posts", type=int, default=200, help="max threads to keep per query (default 200)")
    ap.add_argument("--min-score", type=float, default=2.0,
                    help="minimum relevance score for title/body matches (default 2; raise for stricter results)")
    ap.add_argument("--require", nargs="+", default=[],
                    help="keep only threads that also mention one of these words, e.g. --require india indian")
    ap.add_argument("--comments", choices=["auto", "matched", "none"], default="auto",
                    help="auto: fetch comments and also check threads that only mention the query in comments; "
                         "matched: only threads whose title/body match; none: posts only (fastest)")
    ap.add_argument("--no-comments", dest="comments", action="store_const", const="none", help="same as --comments none")
    ap.add_argument("--subreddits", nargs="+",
                    help="optional: search only inside these subreddits (default: all of Reddit)")
    ap.add_argument("--nsfw", action="store_true", help="include NSFW posts")
    ap.add_argument("--cookies", help="cookie file (Cookie-Editor JSON export or {name: value})")
    ap.add_argument("--output", default=OUTPUT_EXCEL, help="output .xlsx / .csv / .json (merged if it exists)")
    args = ap.parse_args()

    queries = args.queries or args.phrases
    if not queries:
        ap.error("give at least one query, e.g.  python reddit_scraper.py \"croma unboxed\"")
    print(f"🚀 {len(queries)} quer{'y' if len(queries) == 1 else 'ies'} | all of Reddit | last {args.days} days -> {args.output}")
    try:
        scrape(queries, days=args.days, depth=args.depth, max_pages=args.max_pages, max_posts=args.max_posts,
               min_score=args.min_score, require=args.require, comments=args.comments,
               include_nsfw=args.nsfw, output=args.output, cookies_path=args.cookies,
               subreddits=args.subreddits)
    except AuthError as e:
        raise SystemExit(f"🔐 {e}")
    except KeyboardInterrupt:
        raise SystemExit("\n⏹️  interrupted — re-run the same command to resume from the checkpoint")


if __name__ == "__main__":
    main()
