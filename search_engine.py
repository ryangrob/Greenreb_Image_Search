"""CLIP-based image search engine: builds an embedding index for a folder
of images and ranks them against a text query or a reference image."""

import json
import os
import time

import numpy as np
import torch
from PIL import Image

MODEL_NAME = "ViT-B-32"
PRETRAINED = "openai"
MODEL_TAG = f"{MODEL_NAME}/{PRETRAINED}"
CACHE_FILENAME = ".imagesearch_cache.json"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff"}


class ImageSearchEngine:
    def __init__(self):
        self.model = None
        self.preprocess = None
        self.tokenizer = None
        self.device = "cuda" if torch.cuda.is_available() else "cpu"
        self.folder = None
        self.paths = []
        self.meta = []  # parallel to paths: None for local-folder items, dict for SharePoint items
        self.embeddings = None  # np.ndarray, shape (N, D), L2-normalized

    def load_model(self, status_callback=None):
        if self.model is not None:
            return
        if status_callback:
            status_callback("Loading CLIP model (first run downloads ~350MB)...")
        import open_clip

        model, _, preprocess = open_clip.create_model_and_transforms(
            MODEL_NAME, pretrained=PRETRAINED
        )
        model.eval().to(self.device)
        self.model = model
        self.preprocess = preprocess
        self.tokenizer = open_clip.get_tokenizer(MODEL_NAME)
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
        tokens = self.tokenizer([text]).to(self.device)
        with torch.no_grad():
            feat = self.model.encode_text(tokens)
            feat = feat / feat.norm(dim=-1, keepdim=True)
        return feat.squeeze(0).cpu().numpy()

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

    def _rank(self, query_emb, top_k, diversity=0.3):
        """Ranks by relevance to query_emb, then greedily re-orders using
        Maximal Marginal Relevance so near-duplicate photos of the same
        subject (which score almost identically) don't crowd out variety -
        each pick is penalized by its similarity to already-picked results,
        not just scored against the query. `diversity` in [0, 1]: 0 is pure
        relevance ranking (today's old behavior), higher values suppress
        near-duplicates more aggressively. Displayed scores stay pure
        query-relevance - only the ordering/selection is diversity-aware.
        """
        if len(self.paths) == 0:
            return []
        sims = self.embeddings @ query_emb

        # Rank a wider pool by pure relevance first, then diversify just
        # that pool down to top_k - cheap (a few hundred candidates) and
        # keeps genuinely irrelevant images from ever being considered.
        pool_size = min(len(self.paths), max(top_k * 6, 200))
        pool_order = np.argsort(-sims)[:pool_size]
        pool_embs = self.embeddings[pool_order]
        pool_sims = sims[pool_order]

        n = len(pool_order)
        k = min(top_k, n)
        selected = []
        max_sim_to_selected = np.full(n, -1.0, dtype=np.float32)
        remaining = np.ones(n, dtype=bool)

        for _ in range(k):
            mmr_scores = (1 - diversity) * pool_sims - diversity * max_sim_to_selected
            mmr_scores[~remaining] = -np.inf
            best = int(np.argmax(mmr_scores))
            selected.append(best)
            remaining[best] = False
            sim_to_new = pool_embs @ pool_embs[best]
            max_sim_to_selected = np.maximum(max_sim_to_selected, sim_to_new)

        order = pool_order[selected]
        return [(self.paths[i], float(sims[i]), self.meta[i]) for i in order]

    def search_text(self, text, top_k=24, status_callback=None):
        self.load_model(status_callback)
        query_emb = self.embed_text(text)
        return self._rank(query_emb, top_k)
