"""CLIP-based image search engine: builds an embedding index for a folder
of images and ranks them against a text query or a reference image."""

import json
import os
import re
import threading
import time

import numpy as np
from PIL import Image

# torch is imported on first use, in load_model(), not here. It is by far the
# heaviest import in the app - seconds on its own - and importing it at the
# top of this file put all of that in front of the first window. Showing the
# cached library needs only numpy; only searching needs the model.
torch = None

MODEL_NAME = "ViT-B-32"
PRETRAINED = "openai"
MODEL_TAG = f"{MODEL_NAME}/{PRETRAINED}"
CACHE_FILENAME = ".imagesearch_cache.json"
# Cosine similarity above which two images are treated as the same shot and
# the later one is demoted. Chosen from measured data on this library:
# different photos of the same subject cluster around 0.65-0.80, while true
# near-duplicates exceed 0.92. Set to 1.0 to disable demotion entirely.
NEAR_DUPLICATE_THRESHOLD = 0.92
# Caption-shaped phrasings a query is averaged over (see embed_query). The
# bare "{}" keeps the result anchored to the literal query.
QUERY_TEMPLATES = ("a photo of {}", "a picture of {}", "{}")
# Words that start an exclusion, in English and German. CLIP's text encoder
# has no notion of negation - it reads "no christmas" as containing
# "christmas" and pulls results towards it - so exclusions are handled by
# splitting the query and subtracting the unwanted direction instead.
NEGATION_MARKERS = (
    "without", "no ", "not ", "except", "excluding",
    "ohne", "kein ", "keine ", "nicht ",
)
# How strongly an excluded concept is penalised. Applied to standardised
# scores rather than raw similarity: CLIP similarities sit in a narrow band
# (roughly 0.19-0.23 across this library), so subtracting raw values barely
# reorders anything. Measured on real data, 1.5 removes 10-24% of the
# unwanted concept while leaving the wanted one essentially intact.
NEGATION_WEIGHT = 1.5
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff"}


def _standardize(v):
    """Rescales scores to zero mean / unit spread so two different concepts
    can be compared and combined meaningfully."""
    std = v.std()
    return (v - v.mean()) / (std if std else 1.0)



# ---------- German queries ----------
#
# CLIP reads English: a German query ("Burger mit Pommes") lands somewhere
# vague in its space. So a query that looks German is turned into English
# words first ("burger with fries"), from a vocabulary made for this photo
# library - food, drinks, events, golf, the venue, people, moods, seasons,
# colours. Compound words are split ("Weihnachtsfeier" -> "christmas
# party"), endings are tried off ("Getränken" -> "getraenk"), and anything
# unknown - names, brands, English words - passes through unchanged.

_DE_GLOSSARY = """
burger=burger|pommes=fries|fritten=fries|essen=food|gericht=dish|speise=dish|speisen=dishes
teller=plate|nachtisch=dessert|nachspeise=dessert|dessert=dessert|kuchen=cake|torte=cake
eis=ice cream|eisbecher=ice cream sundae|schokolade=chocolate|schoko=chocolate|pizza=pizza
salat=salad|suppe=soup|fleisch=meat|steak=steak|haehnchen=chicken|huhn=chicken|fluegel=wings
chickenwings=chicken wings|fisch=fish|nudel=pasta|nudeln=pasta|brot=bread|broetchen=bun
sandwich=sandwich|wrap=wrap|taco=taco|tacos=tacos|nachos=nachos|kaese=cheese|snack=snack
fingerfood=finger food|haeppchen=canapes|fruehstueck=breakfast|brunch=brunch
mittagessen=lunch|abendessen=dinner|buffet=buffet|bueffet=buffet|grill=grill|grillen=barbecue
vegan=vegan|vegetarisch=vegetarian|obst=fruit|frucht=fruit|fruechte=fruits|erdbeere=strawberry
beere=berries|beeren=berries|popcorn=popcorn|sosse=sauce|sauce=sauce|dip=dip|bowl=bowl
donut=donut|waffel=waffle|pfannkuchen=pancakes|kartoffel=potato|tomate=tomato|zitrone=lemon
limette=lime|minze=mint|kraeuter=herbs|scharf=spicy|suess=sweet|lecker=delicious
getraenk=drink|getraenke=drinks|drink=drink|cocktail=cocktail|cocktails=cocktails|bier=beer
wein=wine|rotwein=red wine|weisswein=white wine|sekt=sparkling wine|champagner=champagne
prosecco=prosecco|glas=glass|glaeser=glasses|flasche=bottle|kaffee=coffee|tee=tea|saft=juice
limonade=lemonade|wasser=water|softdrink=soft drink|cola=cola|longdrink=long drink|shot=shot
shots=shots|bar=bar|theke=bar counter|tresen=bar counter|barkeeper=bartender|eiswuerfel=ice cubes
anstossen=toasting|prost=cheers|trinken=drinking|trinkend=drinking|zapfhahn=beer tap
feier=party|party=party|fest=celebration|feiern=celebrating|feiernd=celebrating|geburtstag=birthday
firma=company|firmen=company|weihnacht=christmas|weihnachten=christmas|markt=market
silvester=new year's eve|neujahr=new year|ostern=easter|halloween=halloween|hochzeit=wedding
junggesellenabschied=bachelor party|jga=bachelor party|veranstaltung=event|event=event
teambuilding=team building|konferenz=conference|besprechung=meeting|meeting=meeting
tagung=conference|praesentation=presentation|empfang=reception|gala=gala|turnier=tournament
wettbewerb=competition|konzert=concert|dj=dj|musik=music|tanzen=dancing|tanz=dance
jubilaeum=anniversary|eroeffnung=opening|feierabend=after work|weihnachtsmarkt=christmas market
mensch=person|menschen=people|leute=people|person=person|personen=people|gruppe=group
freund=friend|freunde=friends|freundin=friend|freundinnen=friends|familie=family
kind=child|kinder=children|jugendliche=teenagers|paar=couple|frau=woman|frauen=women
mann=man|maenner=men|maedchen=girl|junge=boy|jungs=boys|kollege=colleague|kollegen=colleagues
team=team|mitarbeiter=staff|gast=guest|gaeste=guests|kunde=customer|kunden=customers
personal=staff|kellner=waiter|kellnerin=waitress|koch=chef|koeche=chefs|golfer=golfer
golferin=golfer|spieler=player|spielerin=player|trainer=coach|publikum=crowd|zuschauer=spectators
lachen=laughing|lach=laughing|lachend=laughing|laecheln=smiling|laechelnd=smiling
gluecklich=happy|froehlich=cheerful|spass=fun|freude=joy|jubeln=cheering|jubelnd=cheering
spielen=playing|spielend=playing|spiel=game|schlagen=hitting|abschlag=tee shot|schwung=swing
zielen=aiming|gewinnen=winning|sieger=winner|entspannt=relaxed|entspannen=relaxing
gemuetlich=cozy|sitzen=sitting|sitzend=sitting|stehen=standing|stehend=standing|reden=talking
unterhalten=talking|umarmen=hugging|umarmung=hug|posieren=posing|aufregend=exciting
stimmung=atmosphere|atmosphaere=atmosphere|golfen=playing golf|golf=golf|schlaeger=golf club
golfschlaeger=golf club|ball=ball|baelle=balls|golfball=golf ball|golfbaelle=golf balls
handschuh=glove|golfhandschuh=golf glove|abschlagplatz=hitting bay|bucht=bay|range=range
ziel=target|zielscheibe=target|gruen=green|rasen=grass|golfplatz=golf course|platz=place
loch=hole|fahne=flag|bildschirm=screen|monitor=screen|fernseher=tv|leinwand=big screen
terrasse=terrace|dach=roof|dachterrasse=rooftop terrace|lounge=lounge|sofa=sofa|couch=couch
sessel=armchair|tisch=table|tische=tables|stuhl=chair|stuehle=chairs|raum=room|saal=hall
eventraum=event room|flaeche=space|restaurant=restaurant|kueche=kitchen|eingang=entrance
gebaeude=building|fassade=facade|aussen=outside|aussenansicht=exterior|innen=inside
innenansicht=interior|parkplatz=parking lot|auto=car|autos=cars|logo=logo|schild=sign
schriftzug=lettering|dekoration=decoration|deko=decoration|luftballon=balloon
luftballons=balloons|ballon=balloon|ballons=balloons|lichter=lights|beleuchtung=lighting
licht=light|neon=neon|lichterkette=string lights|kerze=candle|kerzen=candles|blume=flower
blumen=flowers|weihnachtsbaum=christmas tree|tannenbaum=christmas tree|geschenk=present
geschenke=presents|schnee=snow|nacht=night|nachts=at night|abend=evening|abends=in the evening
tag=day|tagsueber=daytime|morgen=morning|morgens=in the morning|sonnenuntergang=sunset
sonnenaufgang=sunrise|sonne=sun|sonnig=sunny|sommer=summer|winter=winter|herbst=autumn
fruehling=spring|regen=rain|himmel=sky|wolke=cloud|wolken=clouds|dunkel=dark|hell=bright
rot=red|blau=blue|gelb=yellow|orange=orange|lila=purple|violett=purple|rosa=pink|pink=pink
schwarz=black|weiss=white|grau=grey|braun=brown|gold=gold|silber=silver|bunt=colorful
farbenfroh=colorful|gross=big|klein=small|viele=many|voll=full|leer=empty|modern=modern
schoen=beautiful|frisch=fresh|warm=warm|kalt=cold|heiss=hot|nahaufnahme=close-up
detail=detail|draussen=outdoors|drinnen=indoors|luftaufnahme=aerial view|drohne=drone
hochformat=portrait|querformat=landscape|foto=photo|bild=picture|kamera=camera|handy=phone
selfie=selfie|gruppenfoto=group photo|teamfoto=team photo|portraet=portrait|ohne=without
mit=with|und=and|oder=or|im=in|in=in|am=at|an=at|auf=on|bei=at|beim=at|vom=from|von=of
zum=for|zur=for|fuer=for|unter=under|ueber=over|vor=in front of|hinter=behind|neben=next to
zwischen=between|aus=from|zusammen=together|allein=alone|einzeln=single|viel=a lot
stossen=toasting|stoessen=toasting|stosst=toasting|stoesst=toasting|kindergeburtstag=children's birthday party
"""
# "beim Essen", "zum Golfen": after these, a verb-noun is an activity.
_DE_ACTIVITY = {"essen": "eating", "golfen": "playing golf", "trinken": "drinking",
                "spielen": "playing", "feiern": "celebrating", "tanzen": "dancing",
                "lachen": "laughing", "anstossen": "toasting", "grillen": "having a barbecue"}
# Particles a German verb leaves at the end ("stossen ... an").
_DE_TRAILING_PARTICLES = {"an", "auf", "zu", "mit", "ab", "aus"}

_DE_PHRASES = {
    "von oben": "from above", "bei nacht": "at night", "in der nacht": "at night",
    "stossen an": "toasting", "stoessen an": "toasting", "driving range": "driving range",
    "live musik": "live music", "high five": "high five", "happy hour": "happy hour",
}
_DE_DROP = {"der", "die", "das", "den", "dem", "des", "ein", "eine", "einer", "einem", "einen",
            "eines", "sich", "ist", "sind", "wird", "werden", "einem"}
# English words that would make an English query look German - never used
# to decide that a query is German.
_DE_AMBIGUOUS = {"an", "in", "am", "so", "was", "will", "hat", "tag", "rot", "bad", "hell",
                 "kind", "art", "see", "rock", "gift", "die", "fest", "bar", "spiel", "team",
                 "golf", "range", "ball", "event", "party", "shot", "dip", "drink", "grill",
                 "lounge", "sofa", "couch", "logo", "neon", "gala", "dj", "warm", "modern",
                 "detail", "winter", "orange", "pink", "gold", "foto", "selfie", "bowl", "brunch",
                 "buffet", "cola", "pizza", "steak", "burger", "sandwich", "wrap", "taco",
                 "tacos", "nachos", "snack", "donut", "vegan", "restaurant", "monitor", "trainer",
                 "golfer", "person", "meeting", "dessert", "sauce", "cocktail", "cocktails",
                 "shots", "popcorn", "prosecco", "halloween", "sekt"}


def _de_norm(text):
    text = text.lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        text = text.replace(a, b)
    return text


_DE_WORDS = {}
for _line in _DE_GLOSSARY.strip().splitlines():
    for _pair in _line.split("|"):
        if "=" in _pair:
            _k, _v = _pair.split("=", 1)
            _DE_WORDS[_de_norm(_k.strip())] = _v.strip()
_DE_MARKERS = {k for k in _DE_WORDS if k not in _DE_AMBIGUOUS} | {"mit", "und", "oder", "beim",
               "vom", "zum", "zur", "fuer", "auf", "bei", "der", "das", "ein", "eine", "im", "dem"}
_DE_ENDINGS = ("end", "en", "er", "es", "em", "nd", "e", "n", "s", "t")


def _de_word(token):
    if token in _DE_DROP:
        return ""
    if token in _DE_WORDS:
        return _DE_WORDS[token]
    for ending in _DE_ENDINGS:
        stem = token[: -len(ending)]
        if token.endswith(ending) and len(stem) >= 3 and stem in _DE_WORDS:
            return _DE_WORDS[stem]
    return None


def _de_compound(token, depth=0):
    """A word, or a compound of known words with the longest known ending
    first: "golfhandschuh" -> "golf" + "handschuh"."""
    found = _de_word(token)
    if found is not None:
        return found
    if depth > 2 or len(token) < 6:
        return None
    for i in range(3, len(token) - 2):
        head, tail = token[:i], token[i:]
        right = _de_word(tail)
        if not right:
            continue
        for left_part in (head, head[:-1] if head.endswith("s") else None,
                          head[:-2] if head.endswith("es") else None,
                          head[:-1] if head.endswith("n") else None,
                          head[:-2] if head.endswith("en") else None):
            if not left_part or len(left_part) < 3:
                continue
            left = _de_compound(left_part, depth + 1)
            if left:
                return f"{left} {right}"
    return None


def looks_german(text):
    if any(ch in text for ch in "äöüÄÖÜß"):
        return True
    tokens = re.findall(r"[a-z]+", _de_norm(text))
    return any(t in _DE_MARKERS for t in tokens) or any(
        _de_compound(t) and t not in _DE_AMBIGUOUS and len(t) >= 6 and _de_compound(t) != t
        for t in tokens)


def to_english(text):
    """The query in English words when it looks German, else unchanged."""
    if not text or not looks_german(text):
        return text
    norm = _de_norm(text).replace("-", " ")
    for de, en in _DE_PHRASES.items():
        norm = re.sub(rf"\b{re.escape(de)}\b", en.replace(" ", "_"), norm)
    tokens = re.findall(r"[a-z0-9_']+", norm)
    if len(tokens) > 2 and tokens[-1] in _DE_TRAILING_PARTICLES:
        tokens = tokens[:-1]
    words = []
    previous = ""
    for token in tokens:
        if previous in ("beim", "zum") and token in _DE_ACTIVITY:
            if words and words[-1] in ("at", "for"):
                words.pop()
            words.append(_DE_ACTIVITY[token])
            previous = token
            continue
        previous = token
        if "_" in token:
            words.append(token.replace("_", " "))
            continue
        english = _de_compound(token)
        if english is None:
            words.append(token)          # names, brands, English words
        elif english:
            words.append(english)
    out = []
    for w in words:
        if not out or out[-1] != w:
            out.append(w)
    if len(out) > 1 and out[0] == "group" and out[1] not in ("with", "and", "at", "in"):
        out[0] = "group of"
    return " ".join(out) or text


class ImageSearchEngine:
    def __init__(self):
        self.model = None
        self.preprocess = None
        self.tokenizer = None
        self.device = None  # decided in load_model, once torch is imported
        # Startup warms the model in the background while the user may
        # already be searching; without this both threads would load it.
        self._model_lock = threading.Lock()
        self.folder = None
        self.paths = []
        self.meta = []  # parallel to paths: None for local-folder items, dict for SharePoint items
        self.embeddings = None  # np.ndarray, shape (N, D), L2-normalized
        self.last_translation = None  # the English a German query was searched as

    def load_model(self, status_callback=None):
        if self.model is not None:
            return
        with self._model_lock:
            if self.model is not None:  # another thread finished while we waited
                return
            if status_callback:
                status_callback("Loading the search model...")
            global torch
            if torch is None:
                import torch as _torch
                torch = _torch
            import open_clip

            self.device = "cuda" if torch.cuda.is_available() else "cpu"
            model, _, preprocess = open_clip.create_model_and_transforms(
                MODEL_NAME, pretrained=PRETRAINED
            )
            model.eval().to(self.device)
            self.preprocess = preprocess
            self.tokenizer = open_clip.get_tokenizer(MODEL_NAME)
            # Assigned last: other threads treat a non-None model as "ready",
            # so it must not be visible before the tokenizer and preprocess.
            self.model = model
        if status_callback:
            status_callback("Model loaded.")

    def _cache_path(self, folder):
        return os.path.join(folder, CACHE_FILENAME)

    def _load_cache(self, folder):
        path = self._cache_path(folder)
        if not os.path.exists(path):
            return {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if data.get("model") != MODEL_TAG:
                return {}
            return data.get("images", {})
        except (json.JSONDecodeError, OSError):
            return {}

    def _save_cache(self, folder, images):
        path = self._cache_path(folder)
        data = {"model": MODEL_TAG, "images": images}
        tmp_path = path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        os.replace(tmp_path, path)

    def _scan_images(self, folder):
        found = []
        for root, _dirs, files in os.walk(folder):
            for name in files:
                ext = os.path.splitext(name)[1].lower()
                if ext in IMAGE_EXTENSIONS:
                    found.append(os.path.join(root, name))
        return found

    def embed_image_file(self, path):
        self.load_model()
        img = Image.open(path).convert("RGB")
        tensor = self.preprocess(img).unsqueeze(0).to(self.device)
        with torch.no_grad():
            feat = self.model.encode_image(tensor)
            feat = feat / feat.norm(dim=-1, keepdim=True)
        return feat.squeeze(0).cpu().numpy()

    def embed_image_files(self, paths):
        """Batch version of embed_image_file - runs one forward pass for
        several images at once, which is meaningfully faster per-image on
        CPU than embedding one at a time (better utilizes vectorized/
        multi-core execution than many tiny batch-of-1 calls).

        Each image is opened/preprocessed individually first so one corrupt
        file doesn't fail the whole batch - paths that can't be read are
        simply omitted from the result rather than raising. Returns
        {path: embedding} covering only the paths that succeeded.
        """
        self.load_model()
        valid_paths = []
        tensors = []
        for path in paths:
            try:
                img = Image.open(path).convert("RGB")
                tensors.append(self.preprocess(img))
                valid_paths.append(path)
            except Exception:
                continue
        if not tensors:
            return {}
        batch = torch.stack(tensors).to(self.device)
        with torch.no_grad():
            feats = self.model.encode_image(batch)
            feats = feats / feats.norm(dim=-1, keepdim=True)
        feats = feats.cpu().numpy()
        return {path: feats[i] for i, path in enumerate(valid_paths)}

    def embed_text(self, text):
        self.load_model()
        tokens = self.tokenizer([text]).to(self.device)
        with torch.no_grad():
            feat = self.model.encode_text(tokens)
            feat = feat / feat.norm(dim=-1, keepdim=True)
        return feat.squeeze(0).cpu().numpy()

    @staticmethod
    def split_negation(text):
        """Splits "family with kids in bay no christmas" into
        ("family with kids in bay", "christmas").

        Recognises English and German exclusion words, plus a leading minus
        ("-christmas"). Returns (wanted, unwanted); unwanted is "" when the
        query doesn't exclude anything.
        """
        lowered = text.lower()
        cut = None
        for marker in NEGATION_MARKERS:
            idx = lowered.find(" " + marker if not marker.endswith(" ") else " " + marker)
            if idx != -1 and (cut is None or idx < cut[0]):
                cut = (idx, len(marker) + 1)
        if cut is not None:
            start, offset = cut
            wanted = text[:start].strip()
            unwanted = text[start + offset:].strip()
            if wanted and unwanted:
                return wanted, unwanted

        # "-christmas" style exclusions
        positives, negatives = [], []
        for token in text.split():
            (negatives if token.startswith("-") and len(token) > 1 else positives).append(
                token.lstrip("-")
            )
        if negatives and positives:
            return " ".join(positives), " ".join(negatives)
        return text, ""

    def embed_query(self, text, translate=True):
        """Embeds a search query as an average over several caption-shaped
        phrasings rather than the bare words.

        CLIP was trained on image captions, so it responds better to text
        shaped like one ("a photo of beer") than to a bare noun ("beer").
        Averaging a few templates also damps the effect of any single
        awkward phrasing. The plain "{}" template stays in the mix so this
        never strays far from the raw query.

        Returns (wanted_vec, unwanted_vec); unwanted_vec is None unless the
        query excludes something. The exclusion is applied at ranking time
        rather than folded in here, because it needs the score distribution
        across the library to have any effect.

        Query-side only - image embeddings are untouched, so this needs no
        re-indexing and stays compatible with every existing index.
        """
        wanted, unwanted = self.split_negation(text)
        if translate:
            english_wanted, english_unwanted = to_english(wanted), to_english(unwanted)
            changed = english_wanted != wanted or english_unwanted != unwanted
            self.last_translation = (
                english_wanted + (f" without {english_unwanted}" if english_unwanted else "")
                if changed else None)
            wanted, unwanted = english_wanted, english_unwanted
        else:
            self.last_translation = None
        return (
            self._embed_phrasings(wanted),
            self._embed_phrasings(unwanted) if unwanted else None,
        )

    def _embed_phrasings(self, text):
        feats = [self.embed_text(t.format(text)) for t in QUERY_TEMPLATES]
        avg = np.mean(feats, axis=0)
        norm = np.linalg.norm(avg)
        return avg / norm if norm else avg

    def build_index(self, folder, progress_callback=None, status_callback=None, cancel_check=None):
        """Scans `folder`, reusing cached embeddings for unchanged files.
        Calls progress_callback(done, total) and status_callback(str) as it goes.
        """
        self.load_model(status_callback)

        cache = self._load_cache(folder)
        all_paths = self._scan_images(folder)
        total = len(all_paths)

        new_cache = {}
        paths = []
        embeddings = []

        for i, path in enumerate(all_paths):
            if cancel_check and cancel_check():
                break
            rel = os.path.relpath(path, folder)
            mtime = os.path.getmtime(path)
            cached = cache.get(rel)
            if cached and abs(cached.get("mtime", -1) - mtime) < 1e-6:
                emb = np.array(cached["embedding"], dtype=np.float32)
            else:
                try:
                    emb = self.embed_image_file(path)
                except Exception as exc:  # unreadable/corrupt image file
                    if status_callback:
                        status_callback(f"Skipping {rel}: {exc}")
                    if progress_callback:
                        progress_callback(i + 1, total)
                    continue
            new_cache[rel] = {"mtime": mtime, "embedding": emb.tolist()}
            paths.append(path)
            embeddings.append(emb)
            if progress_callback:
                progress_callback(i + 1, total)

        self._save_cache(folder, new_cache)

        self.folder = folder
        self.paths = paths
        self.meta = [None] * len(paths)
        self.embeddings = np.stack(embeddings) if embeddings else np.zeros((0, 512), dtype=np.float32)
        return len(paths)

    def load_cached_index(self, folder):
        """Loads only what's already cached, without scanning/embedding anything new."""
        cache = self._load_cache(folder)
        paths, embeddings = [], []
        for rel, entry in cache.items():
            full = os.path.join(folder, rel)
            if os.path.exists(full):
                paths.append(full)
                embeddings.append(np.array(entry["embedding"], dtype=np.float32))
        self.folder = folder
        self.paths = paths
        self.meta = [None] * len(paths)
        self.embeddings = np.stack(embeddings) if embeddings else np.zeros((0, 512), dtype=np.float32)
        return len(paths)

    def load_sp_items(self, items):
        """Loads a SharePoint-sourced index: each item is
        {"path": local_thumbnail_path, "embedding": np.ndarray, "meta": {...}}.
        """
        self.folder = None
        self.paths = [item["path"] for item in items]
        self.meta = [item["meta"] for item in items]
        embeddings = [item["embedding"] for item in items]
        self.embeddings = np.stack(embeddings) if embeddings else np.zeros((0, 512), dtype=np.float32)
        return len(self.paths)

    def _rank(self, query_emb, top_k, near_duplicate_threshold=NEAR_DUPLICATE_THRESHOLD,
              bonus=None, subset=None, exclude_emb=None):
        """Ranks purely by relevance to query_emb, demoting only genuine
        near-duplicates (e.g. burst shots of the same moment) to the end.

        `bonus` is an optional array parallel to self.paths, added to the
        similarity before ranking. It carries signals the image embedding
        cannot know about - the name of the SharePoint folder an image
        lives in, and which images people previously opened for this query.
        Kept out of this class deliberately: the caller owns those signals,
        this method only applies them.

        `subset` optionally restricts ranking to specific indices - used to
        search only within a saved collection rather than the whole library.

        Deliberately NOT a Maximal Marginal Relevance / "diversity" ranking.
        MMR penalizes every result by how similar it is to already-picked
        ones, which cannot tell "20 near-identical frames of one person"
        apart from "50 legitimately different photos of beer" - so a query
        like "beer" got its best matches pushed out in favour of less
        relevant but more dissimilar images. Measured on this library,
        different photos of the same subject sit around 0.65-0.80 cosine
        similarity while true near-duplicates sit above 0.92, so a hard
        threshold separates the two cleanly and leaves normal relevance
        ranking untouched.

        Near-duplicates are appended after the unique results rather than
        dropped, so a full top_k is still returned when the library
        genuinely contains many similar shots.
        """
        if len(self.paths) == 0:
            return []
        sims = self.embeddings @ query_emb
        if exclude_emb is not None:
            # Standardise both sides before combining: raw CLIP similarities
            # occupy too narrow a range for a subtraction to reorder much.
            excl = self.embeddings @ exclude_emb
            sims = _standardize(sims) - NEGATION_WEIGHT * _standardize(excl)
        if bonus is not None:
            sims = sims + bonus

        candidates = np.arange(len(self.paths))
        if subset is not None:
            candidates = np.asarray(sorted(subset), dtype=int)
            if len(candidates) == 0:
                return []

        # Only the strongest matches can ever appear, so near-duplicate
        # checking runs over a small pool rather than the whole library.
        pool_size = min(len(candidates), max(top_k * 6, 200))
        pool_order = candidates[np.argsort(-sims[candidates])[:pool_size]]
        pool_embs = self.embeddings[pool_order]

        k = min(top_k, len(pool_order))
        selected, deferred = [], []
        max_sim_to_selected = np.full(len(pool_order), -1.0, dtype=np.float32)

        for i in range(len(pool_order)):
            if len(selected) >= k:
                break
            if max_sim_to_selected[i] > near_duplicate_threshold:
                deferred.append(i)  # near-identical to something already shown
                continue
            selected.append(i)
            max_sim_to_selected = np.maximum(max_sim_to_selected, pool_embs @ pool_embs[i])

        if len(selected) < k:
            selected.extend(deferred[: k - len(selected)])

        order = pool_order[selected]
        return [(self.paths[i], float(sims[i]), self.meta[i]) for i in order]

    def embed_concept(self, prompts):
        """Averages several phrasings into one direction in embedding space,
        used to describe a collection ("people playing golf at night")
        rather than a single search query."""
        v = np.stack([self.embed_text(p) for p in prompts]).mean(axis=0)
        norm = np.linalg.norm(v)
        return v / norm if norm else v

    def score_against(self, concept_vec):
        """Similarity of every indexed image to a concept direction."""
        if self.embeddings is None or len(self.paths) == 0:
            return np.zeros(0, dtype=np.float32)
        return self.embeddings @ concept_vec

    def search_image(self, path, top_k=24, status_callback=None, subset=None):
        """Ranks the library by how much each photo looks like the one at
        `path`.

        CLIP puts images and text in the same space, so this is the same
        ranking as a text search with the query coming from a picture
        instead of words - useful when the thing you want is easier to show
        than to describe ("more like this shot from last year's campaign").
        """
        self.load_model(status_callback)
        query_emb = self.embed_image_file(path)
        return self._rank(query_emb, top_k, subset=subset)

    def search_text(self, text, top_k=24, status_callback=None, bonus=None, subset=None,
                    translate=True):
        self.load_model(status_callback)
        query_emb, exclude_emb = self.embed_query(text, translate=translate)
        return self._rank(
            query_emb, top_k, bonus=bonus, subset=subset, exclude_emb=exclude_emb,
        )
