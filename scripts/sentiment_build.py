"""Score YouTube + X text for infra stocks and add sentiment tabs to the workbook."""
import json, os
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from transformers import pipeline

BASE = os.path.join(os.path.dirname(__file__), "..")
DATA = os.path.join(BASE, "data")
XLSX = os.path.join(BASE, "India_Infra_Budget_vs_Companies.xlsx")
NAMES = {"ULTRACEMCO": "UltraTech Cement", "JSWSTEEL": "JSW Steel", "LT": "Larsen & Toubro",
         "BALKRISIND": "BKT Tyres", "RVNL": "Rail Vikas Nigam", "IRB": "IRB Infrastructure", "NCC": "NCC Ltd",
         "KNRCON": "KNR Constructions", "PNCINFRA": "PNC Infratech", "HGINFRA": "HG Infra Engineering",
         "ASHOKA": "Ashoka Buildcon", "IRCON": "IRCON International", "RITES": "RITES", "KPIL": "Kalpataru Projects",
         "NBCC": "NBCC (India)", "ENGINERSIN": "Engineers India", "BHEL": "BHEL", "ADANIPORTS": "Adani Ports & SEZ",
         "POWERGRID": "Power Grid Corp"}
YEARS = list(range(2015, 2027))

# ---------- gather ----------
items = []  # (ticker, year, source, kind, text, url, likes)
yt = json.load(open(os.path.join(DATA, "youtube_raw.json"))) if os.path.exists(os.path.join(DATA, "youtube_raw.json")) else {}
for tk, d in yt.items():
    for vid, v in d["videos"].items():
        y = int((v.get("upload_date") or "0")[:4] or 0)
        url = f"https://www.youtube.com/watch?v={vid}"
        if v.get("title") and y in YEARS:
            items.append((tk, y, "YouTube", "Video title", v["title"], url, v.get("likes") or 0))
        for c in v.get("comments", []):
            cy = __import__("datetime").datetime.utcfromtimestamp(c["ts"]).year if c.get("ts") else y
            if c.get("text") and cy in YEARS:
                items.append((tk, cy, "YouTube", "Comment", c["text"][:400], url, c.get("likes") or 0))
tw_path = os.path.join(DATA, "twitter_raw.json")
tw = json.load(open(tw_path)) if os.path.exists(tw_path) else {}
for key, lst in tw.items():
    tk, y = key.rsplit("_", 1)
    for t in lst:
        items.append((tk, int(y), "X (Twitter)", "Tweet", t["t"], "https://x.com" + t["u"], t.get("lk") or 0))
print(len(items), "items")

# ---------- score ----------
clf = pipeline("sentiment-analysis", model="cardiffnlp/twitter-xlm-roberta-base-sentiment-multilingual",
               truncation=True, max_length=256)
preds = clf([i[4] for i in items], batch_size=32)
VAL = {"positive": 1, "neutral": 0, "negative": -1}
scored = [(*it, p["label"], VAL[p["label"].lower()], p["score"]) for it, p in zip(items, preds)]
json.dump(scored, open(os.path.join(DATA, "sentiment_scored.json"), "w"))

agg = defaultdict(lambda: {"n": 0, "pos": 0, "neg": 0, "neu": 0, "yt": 0, "tw": 0, "best": None, "worst": None})
for tk, y, src, kind, text, url, lk, lab, v, conf in scored:
    a = agg[(tk, y)]; a["n"] += 1; a["yt" if src == "YouTube" else "tw"] += 1
    a[{1: "pos", -1: "neg", 0: "neu"}[v]] += 1
    if v == 1 and (not a["best"] or conf > a["best"][0]): a["best"] = (conf, text, url)
    if v == -1 and (not a["worst"] or conf > a["worst"][0]): a["worst"] = (conf, text, url)
net = lambda a: (a["pos"] - a["neg"]) / a["n"] if a["n"] else None

# ---------- workbook ----------
wb = load_workbook(XLSX)
for n in ("Sentiment Heatmap", "Sentiment by Year", "Sentiment Raw"):
    if n in wb.sheetnames: del wb[n]
NAVY, BLUE, GREY = "1F3864", "DDEBF7", "F2F2F2"
thin = Side(style="thin", color="D9D9D9"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
B = Font(name="Calibri", size=10, bold=True, color="000000"); N = Font(name="Calibri", size=10, color="000000")
L = Font(name="Calibri", size=10, color="000000", underline="single")

def head(ws, r, vals):
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v); c.font = B; c.fill = PatternFill("solid", fgColor=BLUE); c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")

def titled(name, title, sub):
    ws = wb.create_sheet(name, 3); ws.sheet_view.showGridLines = False
    ws["A1"] = title; ws["A1"].font = Font(bold=True, size=16, color=NAVY)
    ws["A2"] = sub; ws["A2"].font = Font(italic=True, size=10, color="595959")
    return ws

# Heatmap
ws = titled("Sentiment Heatmap", "Net Sentiment by Company and Year (YouTube + X)",
            "Net sentiment = (% positive - % negative). Green = bullish, red = bearish. Blank = no posts found. n = number of posts in brackets on 'Sentiment by Year'.")
head(ws, 4, ["Company"] + [f"FY{str(y)[2:]}*" for y in YEARS] + ["All years"])
r = 5
for i, (tk, nm) in enumerate(NAMES.items()):
    ws.cell(r, 1, nm).font = B; ws.cell(r, 1).border = BOX
    tot = {"n": 0, "pos": 0, "neg": 0}
    for j, y in enumerate(YEARS, 2):
        a = agg.get((tk, y))
        c = ws.cell(r, j, round(net(a), 2) if a and a["n"] >= 3 else None); c.number_format = "+0.00;-0.00;0.00"
        c.border = BOX; c.font = N; c.alignment = Alignment(horizontal="center")
        if a:
            for k in tot: tot[k] += a[k]
    c = ws.cell(r, len(YEARS) + 2, round((tot["pos"] - tot["neg"]) / tot["n"], 2) if tot["n"] else None)
    c.font = B; c.border = BOX; c.number_format = "+0.00;-0.00;0.00"; r += 1
ws.conditional_formatting.add(f"B5:{chr(65 + len(YEARS) + 1)}{r - 1}",
                              ColorScaleRule(start_type="num", start_value=-0.5, start_color="F8696B",
                                             mid_type="num", mid_value=0, mid_color="FFFFFF",
                                             end_type="num", end_value=0.5, end_color="63BE7B"))
r += 1
ws.cell(r, 1, "Posts found (YouTube titles + comments + tweets)").font = B; r += 1
head(ws, r, ["Company"] + [f"FY{str(y)[2:]}*" for y in YEARS] + ["Total"]); r += 1
for tk, nm in NAMES.items():
    ws.cell(r, 1, nm).font = N
    cnt = [agg[(tk, y)]["n"] if (tk, y) in agg else 0 for y in YEARS]
    for j, v in enumerate(cnt + [sum(cnt)], 2):
        c = ws.cell(r, j, v); c.font = N; c.border = BOX; c.alignment = Alignment(horizontal="center")
    r += 1
ws.cell(r + 1, 1, "* Calendar year of the post (Jan-Dec). Years with fewer than 3 posts are left blank.").font = Font(italic=True, size=9)
ws.column_dimensions["A"].width = 26
for j in range(2, len(YEARS) + 3): ws.column_dimensions[chr(64 + j)].width = 9
ws.freeze_panes = "B5"

# By year detail
ws = titled("Sentiment by Year", "Sentiment Detail - Company x Year",
            "Model: cardiffnlp/twitter-xlm-roberta-base-sentiment-multilingual (handles English, Hindi, Hinglish).")
cols = ["Company", "Year", "Posts", "YouTube", "X", "% Positive", "% Neutral", "% Negative", "Net", "Most positive post", "Most negative post"]
head(ws, 4, cols); r = 5
for i, (tk, nm) in enumerate(NAMES.items()):
    for y in YEARS:
        a = agg.get((tk, y))
        if not a: continue
        vals = [nm, y, a["n"], a["yt"], a["tw"], a["pos"] / a["n"], a["neu"] / a["n"], a["neg"] / a["n"], net(a)]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v); c.font = N; c.border = BOX; c.alignment = Alignment(vertical="top", wrap_text=True)
            if i % 2: c.fill = PatternFill("solid", fgColor=GREY)
            if j in (6, 7, 8): c.number_format = "0%"
            if j == 9: c.number_format = "+0.00;-0.00;0.00"
        for j, k in ((10, "best"), (11, "worst")):
            c = ws.cell(r, j)
            if a[k]:
                c.value = a[k][1][:250]; c.hyperlink = a[k][2]; c.font = L
            c.border = BOX; c.alignment = Alignment(vertical="top", wrap_text=True)
            if i % 2: c.fill = PatternFill("solid", fgColor=GREY)
        r += 1
for col, w in zip("ABCDEFGHIJK", [22, 7, 7, 9, 6, 10, 10, 10, 8, 60, 60]): ws.column_dimensions[col].width = w
ws.freeze_panes = "C5"

# Raw
ws = titled("Sentiment Raw", "All Scored Posts", "Every YouTube title/comment and tweet with its sentiment label. Click a link to open the original.")
head(ws, 4, ["Company", "Year", "Source", "Type", "Text", "Sentiment", "Confidence", "Likes", "Link"]); r = 5
for tk, y, src, kind, text, url, lk, lab, v, conf in sorted(scored, key=lambda x: (x[0], x[1])):
    for j, val in enumerate([NAMES.get(tk, tk), y, src, kind, text, lab.title(), round(conf, 2), lk], 1):
        c = ws.cell(r, j, val); c.font = N; c.alignment = Alignment(vertical="top", wrap_text=(j == 5))
    c = ws.cell(r, 9, "open"); c.hyperlink = url; c.font = L; r += 1
for col, w in zip("ABCDEFGHI", [22, 7, 12, 12, 90, 11, 11, 8, 8]): ws.column_dimensions[col].width = w
ws.auto_filter.ref = f"A4:I{r - 1}"; ws.freeze_panes = "A5"

wb.save(XLSX)
print("saved", len(scored), "scored;", len(agg), "company-years")
