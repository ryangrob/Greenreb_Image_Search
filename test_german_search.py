"""Measures German search on the real library.

Run in C:\\ImageSearchBuild-v2.4.1:   py test_german_search.py

For each pair it searches the German words as they are, the German words
translated, and the English words, and saves a contact sheet per pair into
german_test\\ plus a summary (how many of the top 24 match the English search).
Takes about a minute and a half. Changes nothing in the app.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
os.environ.setdefault("HF_HOME", os.path.join(HERE, "model_cache"))

from PIL import Image, ImageDraw, ImageFont  # noqa: E402
import numpy as np  # noqa: E402
from search_engine import ImageSearchEngine, to_english  # noqa: E402

PAIRS = [
    ("Burger mit Pommes", "burger with fries"),
    ("Nachtisch", "dessert"),
    ("Freunde stoßen mit Getränken an", "friends toasting with drinks"),
    ("Golfplatz bei Nacht", "golf range at night"),
    ("Eventraum für eine Feier", "event room set up for a party"),
    ("Weihnachtsfeier", "christmas party"),
    ("Kindergeburtstag", "children's birthday party"),
    ("Golfhandschuh", "golf glove"),
    ("Gruppe lachender Kollegen", "group of laughing colleagues"),
    ("Cocktails an der Bar", "cocktails at the bar"),
    ("Familie mit Kindern beim Golfen", "family with children playing golf"),
    ("Sonnenuntergang über der Driving Range", "sunset over the driving range"),
    ("leerer Veranstaltungsraum", "empty event room"),
    ("Paar beim Essen", "couple eating"),
    ("Dachterrasse im Sommer", "rooftop terrace in summer"),
    ("Bier und Burger", "beer and burger"),
    ("Schokoladen-Dessert", "chocolate dessert"),
    ("Junggesellenabschied", "bachelor party"),
    ("Firmenfeier mit DJ", "company party with dj"),
    ("Weihnachtsmarkt ohne Menschen", "christmas market without people"),
]
TOP = 24
SHOW = 8


def load_library():
    base = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ImageSearch",
                        "sharepoint_cache", "FotosVideos")
    embeddings = np.load(os.path.join(base, "index_embeddings.npy"))
    with open(os.path.join(base, "index_meta.json"), encoding="utf-8") as f:
        meta = json.load(f)
    return embeddings, meta["paths"], meta["meta"]


def sheet(rows, path):
    cell, label_w, gap = 150, 230, 6
    width = label_w + SHOW * (cell + gap)
    img = Image.new("RGB", (width, len(rows) * (cell + gap) + gap), (9, 11, 10))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(os.path.join(HERE, "assets", "fonts", "Inter-Medium.ttf"), 15)
    except Exception:
        font = ImageFont.load_default()
    for r, (label, results) in enumerate(rows):
        y = gap + r * (cell + gap)
        draw.multiline_text((10, y + 10), label, fill=(230, 230, 230), font=font, spacing=4)
        for c, (p, _s, _m) in enumerate(results[:SHOW]):
            try:
                with Image.open(p) as t:
                    t = t.convert("RGB")
                    t.thumbnail((cell, cell))
                    img.paste(t, (label_w + c * (cell + gap), y))
            except Exception:
                pass
    img.save(path, quality=85)


def wrap(text, n=26):
    words, lines, line = text.split(), [], ""
    for w in words:
        if len(line) + len(w) + 1 > n:
            lines.append(line)
            line = w
        else:
            line = f"{line} {w}".strip()
    lines.append(line)
    return "\n".join(lines)


def main():
    out = os.path.join(HERE, "german_test")
    os.makedirs(out, exist_ok=True)
    engine = ImageSearchEngine()
    engine.embeddings, engine.paths, engine.meta = load_library()
    print(f"Library: {len(engine.paths):,} photos. Loading the model...")
    engine.load_model()
    lines = ["German | as English | raw German vs English | translated vs English (overlap of top 24)"]
    raw_total = tr_total = 0
    for n, (german, english) in enumerate(PAIRS, start=1):
        ref = engine.search_text(english, TOP, translate=False)
        raw = engine.search_text(german, TOP, translate=False)
        tr = engine.search_text(german, TOP, translate=True)
        ref_set = {p for p, _s, _m in ref}
        raw_hit = len(ref_set & {p for p, _s, _m in raw})
        tr_hit = len(ref_set & {p for p, _s, _m in tr})
        raw_total += raw_hit
        tr_total += tr_hit
        lines.append(f"{german} | {to_english(german)} | {raw_hit}/{TOP} | {tr_hit}/{TOP}")
        sheet([(wrap(f"German as typed:\n{german}"), raw),
               (wrap(f"Translated:\n{to_english(german)}"), tr),
               (wrap(f"English:\n{english}"), ref)],
              os.path.join(out, f"{n:02d}.jpg"))
        print(lines[-1])
    lines.append(f"TOTAL | | {raw_total}/{TOP * len(PAIRS)} | {tr_total}/{TOP * len(PAIRS)}")
    with open(os.path.join(out, "summary.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(lines[-1])
    print(f"Contact sheets and summary in {out}")


if __name__ == "__main__":
    main()
