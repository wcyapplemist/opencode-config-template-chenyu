"""
outline_store.py
================
Persist / load the multi-stage generation **outline artifact** (Phase 1 Track C,
#21 / #24).

The outline (Stage 1 output, plain text) is saved to the OS temp dir so it can
be surfaced to the user and edited in the interactive-checkpoint flow (#24),
without polluting the repository. An optional ``deck_id`` disambiguates
concurrent runs.

Public API::

    save_outline(outline_text, deck_id=None) -> Path
    load_outline(path) -> str
"""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

_TEMP_DIR = Path(__file__).resolve().parent.parent.parent.parent.parent / ".opencode" / "tmp"
_PREFIX = "pptx_outline"


def _ensure_dir() -> Path:
    _TEMP_DIR.mkdir(parents=True, exist_ok=True)
    return _TEMP_DIR


def save_outline(outline_text: str, deck_id: Optional[str] = None) -> Path:
    """Persist ``outline_text`` and return its path."""
    stem = deck_id or f"{int(time.time() * 1000)}"
    path = _ensure_dir() / f"{_PREFIX}_{stem}.md"
    path.write_text(outline_text or "", encoding="utf-8")
    logger.info("Outline artifact saved: %s", path)
    return path


def load_outline(path) -> str:
    """Read a previously saved outline artifact."""
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Outline artifact not found: {p}")
    return p.read_text(encoding="utf-8")


def latest_outline() -> Optional[Path]:
    """Return the most recently saved outline artifact, if any."""
    if not _TEMP_DIR.exists():
        return None
    candidates = sorted(_TEMP_DIR.glob(f"{_PREFIX}_*.md"))
    return candidates[-1] if candidates else None
