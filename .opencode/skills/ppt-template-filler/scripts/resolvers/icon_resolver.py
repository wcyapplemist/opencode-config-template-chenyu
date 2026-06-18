"""
icon_resolver.py
================
Resolves ``icon_query`` (a semantic keyword) into a local ``icon_path``
(SVG/PNG) matched against a local icon library.

Matching strategy (in order of availability):
    1. ``match_fn``  — injectable callable(query, config) -> Path | None
                       (tests / custom indexers).
    2. Keyword match against the file names of an on-disk icon library
       (``config["icon"]["path"]``), e.g. a Phosphor export.

Graceful by contract: no library configured, no match, or any error -> the
slide is returned unchanged (icon omitted), never raising.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Dict, Optional

logger = logging.getLogger(__name__)

_ICON_EXTENSIONS = (".svg", ".png", ".jpg", ".jpeg", ".webp")


def _keyword_index(library_dir: Path) -> Dict[str, Path]:
    """Map lowercased, tokenized icon file names -> file path."""
    index: Dict[str, Path] = {}
    if not library_dir.exists():
        return index
    for path in library_dir.rglob("*"):
        if path.suffix.lower() in _ICON_EXTENSIONS:
            # e.g. "arrow-up-right.svg" -> tokens "arrow up right"
            tokens = path.stem.lower().replace("-", " ").replace("_", " ")
            index[tokens] = path
    return index


def _match(query: str, config: Dict[str, Any]) -> Optional[Path]:
    icon_cfg = config.get("icon", {})

    match_fn: Optional[Callable[[str, Dict[str, Any]], Optional[Path]]] = icon_cfg.get("match_fn")
    if callable(match_fn):
        try:
            return match_fn(query, config)
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("icon match_fn failed for '%s': %s", query, exc)
            return None

    library = icon_cfg.get("path")
    if not library:
        logger.warning("Icon resolver: no icon library path configured — skipping")
        return None

    library_dir = Path(library)
    index = _keyword_index(library_dir)
    if not index:
        logger.warning("Icon resolver: library '%s' has no icons — skipping", library_dir)
        return None

    q_tokens = query.lower().replace("-", " ").replace("_", " ").strip()
    # Exact token match first, then substring containment.
    if q_tokens in index:
        return index[q_tokens]
    for key, path in index.items():
        if q_tokens and (q_tokens in key or key in q_tokens):
            return path
    # Partial: any query token present in an icon name.
    q_words = [w for w in q_tokens.split() if len(w) > 2]
    for key, path in index.items():
        if any(w in key for w in q_words):
            return path
    logger.warning("Icon resolver: no match for '%s'", query)
    return None


def resolve(slide_data: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    """Resolve an ``icon_query`` placeholder (returns a new dict)."""
    out = dict(slide_data)
    query = (out.get("icon_query") or "").strip()
    if not query:
        return out

    icon_path = _match(query, config)
    if icon_path is None:
        return out

    out["icon_path"] = str(icon_path)
    out.pop("icon_query", None)
    logger.info("Icon resolver: '%s' -> %s", query, icon_path)
    return out
