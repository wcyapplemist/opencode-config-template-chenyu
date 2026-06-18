"""
image_resolver.py
=================
Resolves image placeholders (``image_prompt`` / ``image_query``) into a local
``image_path`` consumable by ``ppt_builder._add_image_to_slide()`` (#18).

Graceful by contract: any failure (missing config, network error, provider
returns nothing) logs a warning and returns the slide **unchanged** — the deck
is always produced, only the affected image is omitted.

Provider selection (``config["image"]``):
    * ``fetch_fn``  — injectable callable(query, config) -> bytes | None
                      (primary path for tests / custom providers).
    * ``provider``  — ``"pexels"`` / ``"unsplash"`` / ``"ai"``; uses HTTP when an
                      API key env var is configured.
"""

from __future__ import annotations

import hashlib
import logging
import os
from pathlib import Path
from typing import Any, Callable, Dict, Optional

logger = logging.getLogger(__name__)

_DEFAULT_CACHE_DIR = Path("output") / "_image_cache"
_IMAGE_PLACEHOLDER_KEYS = ("image_prompt", "image_query", "image_source")


def _cache_path(query: str, config: Dict[str, Any]) -> Path:
    cache_dir = Path(config.get("image", {}).get("cache_dir") or _DEFAULT_CACHE_DIR)
    cache_dir.mkdir(parents=True, exist_ok=True)
    digest = hashlib.md5(query.encode("utf-8")).hexdigest()[:12]
    return cache_dir / f"img_{digest}.png"


def _fetch_pexels(query: str, api_key: str) -> Optional[bytes]:
    import requests  # local import keeps the module importable without requests
    resp = requests.get(
        "https://api.pexels.com/v1/search",
        params={"query": query, "per_page": 1},
        headers={"Authorization": api_key},
        timeout=15,
    )
    resp.raise_for_status()
    photos = resp.json().get("photos") or []
    if not photos:
        return None
    img_url = photos[0]["src"]["large"]
    img_resp = requests.get(img_url, timeout=30)
    img_resp.raise_for_status()
    return img_resp.content


def _fetch_unsplash(query: str, api_key: str) -> Optional[bytes]:
    import requests
    resp = requests.get(
        "https://api.unsplash.com/search/photos",
        params={"query": query, "per_page": 1},
        headers={"Authorization": f"Client-ID {api_key}"},
        timeout=15,
    )
    resp.raise_for_status()
    results = resp.json().get("results") or []
    if not results:
        return None
    img_resp = requests.get(results[0]["urls"]["regular"], timeout=30)
    img_resp.raise_for_status()
    return img_resp.content


def _resolve_image_bytes(query: str, config: Dict[str, Any]) -> Optional[bytes]:
    """Return image bytes for ``query`` or ``None`` (never raises)."""
    img_cfg = config.get("image", {})

    # Injectable hook (tests / custom providers).
    fetch_fn: Optional[Callable[[str, Dict[str, Any]], Optional[bytes]]] = img_cfg.get("fetch_fn")
    if callable(fetch_fn):
        try:
            return fetch_fn(query, config)
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("image fetch_fn failed for '%s': %s", query, exc)
            return None

    provider = img_cfg.get("provider", "pexels")
    key_env = img_cfg.get("api_key_env", "PEXELS_API_KEY")
    api_key = os.environ.get(key_env, "")
    if not api_key:
        logger.warning(
            "Image resolver: no API key (%s) for provider '%s' — skipping image",
            key_env, provider,
        )
        return None

    try:
        if provider == "pexels":
            return _fetch_pexels(query, api_key)
        if provider == "unsplash":
            return _fetch_unsplash(query, api_key)
        logger.warning("Image resolver: unknown provider '%s' — skipping", provider)
    except Exception as exc:
        logger.warning("Image resolver failed for '%s': %s", query, exc)
    return None


def resolve(slide_data: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    """Resolve image placeholders on a single slide (returns a new dict)."""
    out = dict(slide_data)

    has_placeholder = any(k in out and out[k] for k in _IMAGE_PLACEHOLDER_KEYS)
    if not has_placeholder:
        return out  # nothing to do

    query = (out.get("image_query") or out.get("image_prompt") or "").strip()
    if not query:
        logger.warning("Image resolver: placeholder present but empty query — skipping")
        return out

    image_bytes = _resolve_image_bytes(query, config)
    if not image_bytes:
        return out

    dest = _cache_path(query, config)
    try:
        dest.write_bytes(image_bytes)
    except Exception as exc:
        logger.warning("Image resolver: cannot write cache file %s: %s", dest, exc)
        return out

    # Set concrete value, keep placeholder keys out of the renderer's way.
    out["image_path"] = str(dest)
    if not out.get("image_position"):
        default_pos = config.get("image", {}).get("default_position")
        if default_pos:
            out["image_position"] = default_pos
    for k in _IMAGE_PLACEHOLDER_KEYS:
        out.pop(k, None)
    logger.info("Image resolver: '%s' -> %s", query, dest)
    return out
