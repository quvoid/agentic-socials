"""Collect YouTube videos + comments about Indian infra stocks (resumable cache)."""
import json, os, sys, time
import yt_dlp

OUT = os.path.join(os.path.dirname(__file__), "..", "data")
os.makedirs(OUT, exist_ok=True)
CACHE = os.path.join(OUT, "youtube_raw.json")

COMPANIES = {
    "ULTRACEMCO": "UltraTech Cement", "JSWSTEEL": "JSW Steel", "LT": "Larsen and Toubro",
    "BALKRISIND": "Balkrishna Industries BKT", "RVNL": "RVNL", "IRB": "IRB Infrastructure",
    "NCC": "NCC Ltd", "KNRCON": "KNR Constructions", "PNCINFRA": "PNC Infratech",
    "HGINFRA": "HG Infra Engineering", "ASHOKA": "Ashoka Buildcon", "IRCON": "IRCON International",
    "RITES": "RITES Ltd", "KPIL": "Kalpataru Projects", "NBCC": "NBCC India",
    "ENGINERSIN": "Engineers India", "BHEL": "BHEL", "ADANIPORTS": "Adani Ports", "POWERGRID": "Power Grid Corporation",
}
PER_QUERY, MAX_COMMENTS = 8, 80
YEARS = range(2015, 2027)

data = json.load(open(CACHE)) if os.path.exists(CACHE) else {}
flat = yt_dlp.YoutubeDL({"quiet": True, "no_warnings": True, "extract_flat": True, "skip_download": True})
full = yt_dlp.YoutubeDL({"quiet": True, "no_warnings": True, "cookiesfrombrowser": ("chrome",), "sleep_interval_requests": 1, "sleep_interval": 2, "max_sleep_interval": 5, "skip_download": True, "getcomments": True, "ignoreerrors": True,
                         "extractor_args": {"youtube": {"max_comments": [str(MAX_COMMENTS), "all", "0"],
                                                        "comment_sort": ["top"]}}})
for tk, name in COMPANIES.items():
    d = data.setdefault(tk, {"videos": {}})
    ids = []
    for q in [f"{name} share {y}" for y in YEARS]:
        try:
            res = flat.extract_info(f"ytsearch{PER_QUERY}:{q}", download=False)
            ids += [e["id"] for e in res.get("entries", []) if e and e.get("id")]
        except Exception as e:
            print("search fail", tk, e, file=sys.stderr)
    for vid in dict.fromkeys(ids):
        if vid in d["videos"]: continue
        try:
            info = full.extract_info(f"https://www.youtube.com/watch?v={vid}", download=False) or {}
        except Exception as e:
            info = {}
        if not info.get("title"): continue  # failed fetch: do not cache, retry next run
        d["videos"][vid] = {
            "title": info.get("title"), "channel": info.get("channel"), "upload_date": info.get("upload_date"),
            "views": info.get("view_count"), "likes": info.get("like_count"),
            "description": (info.get("description") or "")[:500],
            "comments": [{"text": c.get("text"), "ts": c.get("timestamp"), "likes": c.get("like_count")}
                         for c in (info.get("comments") or [])],
        }
        json.dump(data, open(CACHE, "w"))
    print(tk, len(d["videos"]), "videos", sum(len(v["comments"]) for v in d["videos"].values()), "comments", flush=True)
print("DONE")
