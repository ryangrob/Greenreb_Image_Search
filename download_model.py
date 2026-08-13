"""One-time helper: pre-downloads the CLIP model weights into a local
model_cache/ folder so build.bat can bundle them into the .exe, letting end
users skip the ~350MB first-run download entirely."""

import os

CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "model_cache")
os.environ["HF_HOME"] = CACHE_DIR
os.environ.setdefault("TORCH_HOME", CACHE_DIR)

from search_engine import MODEL_NAME, PRETRAINED  # noqa: E402

import open_clip  # noqa: E402

print(f"Downloading {MODEL_NAME}/{PRETRAINED} into {CACHE_DIR} ...")
open_clip.create_model_and_transforms(MODEL_NAME, pretrained=PRETRAINED)
print("Done.")
