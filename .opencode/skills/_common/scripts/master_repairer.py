"""US-4.8 Phase 1 — Master repair cascade for masterless PPTX files.

When a user-supplied ``.pptx`` has no slide master (Scenario A), this module
repairs it via a **three-level cascade** (Chain of Responsibility):

- **Level 1** — salvage ``ppt/theme/theme1.xml`` from the zip (exact fidelity).
- **Level 2** — scavenge explicit styles from slide XML (best-effort fidelity).
- **Level 3** — fallback to ``default.pptx``'s theme (no user styling).

The repair strategy is **NOT** cross-file ``SlideMasterPart`` injection (which
is fragile OOXML surgery with no python-pptx API). Instead it:

1. Copies ``default.pptx`` (which already has a valid master + all 35 layouts)
   to a derived file.
2. Optionally replaces the derived file's ``ppt/theme/theme1.xml`` with the
   salvaged/scavenged theme content.

The result is a valid template with default's master/layout structure + the
user's theme (colors + fonts). The engine removes all slides before rendering,
so the user's original slides (which referenced orphaned layouts) don't matter.
"""
from __future__ import annotations

import logging
import os
import shutil
import tempfile
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Tuple

from lxml import etree
from io import BytesIO

logger = logging.getLogger(__name__)

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_THEME_PATH = "ppt/theme/theme1.xml"
_THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
_CONTENT_TYPES_XML = "[Content_Types].xml"
_SLIDES_DIR = "ppt/slides/"

# Font-size tiers for Level-2 scavenge (in centipoints; OOXML sz attribute).
# ≥28pt = title tier; <28pt = body tier. 1pt = 100 centipoints.
_TITLE_SIZE_THRESHOLD = 2800


@dataclass
class RepairResult:
    """Outcome of ``repair_if_needed``."""

    level: Literal["none", "L1", "L2", "L3"]
    mutated: bool
    theme_source: str  # "salvaged", "scavenged", "default", "n/a"
    repaired_path: Optional[str] = None  # the derived file path if mutated


def repair_if_needed(
    prs: Any,
    template_path: str,
    default_template_path: str,
) -> RepairResult:
    """Check if ``prs`` needs master repair (Scenario A); if so, repair it.

    **CRIT-1**: Must be called BEFORE ``get_render_contract`` in
    ``generate_ppt_from_data``, so the contract is fetched from the repaired prs.

    **CRIT-3**: Receives ``template_path`` (for zip-level salvage) and
    ``default_template_path`` (for the master skeleton) — **MAJOR-1**: neither
    is hardcoded; both are injected by the caller.

    Returns a :class:`RepairResult`. If ``mutated`` is True, the caller should
    reload ``prs`` from ``repaired_path``.
    """
    try:
        masters = list(prs.slide_masters)
    except Exception:
        masters = []
    if masters:
        return RepairResult(level="none", mutated=False, theme_source="n/a")

    logger.info(
        "Template has no slide master (Scenario A) — starting repair cascade for %s",
        template_path,
    )

    # --- Level 1: salvage theme1.xml from the zip ---
    theme_bytes = _salvage_theme_part(template_path)
    level: str
    theme_source: str

    if theme_bytes is not None:
        level = "L1"
        theme_source = "salvaged"
        logger.info("Level 1: salvaged ppt/theme/theme1.xml from %s", template_path)
    else:
        # --- Level 2: scavenge styles from slide XML ---
        theme_bytes = _scavenge_slide_styles(template_path, default_template_path)
        if theme_bytes is not None:
            level = "L2"
            theme_source = "scavenged"
            logger.info("Level 2: scavenged theme from slide styles in %s", template_path)
        else:
            # --- Level 3: default fallback ---
            level = "L3"
            theme_source = "default"
            theme_bytes = None
            logger.info("Level 3: using default.pptx theme (no user styling recoverable)")

    # Build the derived file: copy default.pptx, optionally replace theme.
    repaired_path = _build_repaired_file(
        template_path, default_template_path, theme_bytes, level
    )

    return RepairResult(
        level=level,
        mutated=True,
        theme_source=theme_source,
        repaired_path=repaired_path,
    )


# ---------------------------------------------------------------------------
# Level 1: Salvage ppt/theme/theme1.xml from the zip
# ---------------------------------------------------------------------------

def _salvage_theme_part(pptx_path: str) -> Optional[bytes]:
    """Level 1: read ``ppt/theme/theme1.xml`` directly from the PPTX zip.

    The theme part often survives when the master part is stripped, because it
    is an independent zip entry (``ppt/theme/theme1.xml``), not embedded inside
    the master.

    Returns the raw theme XML bytes, or ``None`` if the part is absent.
    """
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            if _THEME_PATH in z.namelist():
                return z.read(_THEME_PATH)
    except (zipfile.BadZipFile, OSError) as exc:
        logger.warning("Level 1 salvage failed (%s)", exc)
    return None


# ---------------------------------------------------------------------------
# Level 2: Scavenge explicit styles from slide XML
# ---------------------------------------------------------------------------

@dataclass
class _ScavengedTheme:
    """Aggregated style data from slide XML."""

    major_font: Optional[str] = None  # title-tier typeface
    minor_font: Optional[str] = None  # body-tier typeface
    colors: List[str] = field(default_factory=list)  # all srgbClr hex values


def _scavenge_slide_styles(
    pptx_path: str, default_template_path: str
) -> Optional[bytes]:
    """Level 2: aggregate explicit styles from surviving slide XML.

    Walks ``ppt/slides/slideN.xml`` for ``<a:rPr>`` (font/size/color) and
    ``<p:spPr>`` fill colors. Aggregates into a dominant major/minor font +
    top accent colors, then synthesizes a ``<a:theme>`` by overriding
    ``default.pptx``'s theme clrScheme + fontScheme.

    Returns the synthesized theme XML bytes, or ``None`` if no slide styles
    are recoverable.
    """
    try:
        scavenged = _aggregate_slide_styles(pptx_path)
    except Exception as exc:
        logger.warning("Level 2 scavenge failed (%s)", exc)
        return None

    if not scavenged.major_font and not scavenged.minor_font and not scavenged.colors:
        logger.info("Level 2: no explicit styles found in slides")
        return None

    # Load default.pptx's theme as the structural base.
    try:
        with zipfile.ZipFile(default_template_path, "r") as z:
            base_theme_bytes = z.read(_THEME_PATH)
    except Exception as exc:
        logger.warning("Level 2: could not read default theme (%s)", exc)
        return None

    return _build_synthetic_theme(base_theme_bytes, scavenged)


def _aggregate_slide_styles(pptx_path: str) -> _ScavengedTheme:
    """Parse all slide XML files and aggregate font/color frequencies."""
    typeface_by_tier: Dict[str, Counter] = {"title": Counter(), "body": Counter()}
    all_colors: List[str] = []

    with zipfile.ZipFile(pptx_path, "r") as z:
        slide_files = sorted(
            n for n in z.namelist() if n.startswith(_SLIDES_DIR) and n.endswith(".xml")
        )
        for slide_path in slide_files:
            try:
                root = etree.parse(BytesIO(z.read(slide_path))).getroot()
            except Exception:
                continue
            # Walk all <a:rPr> elements for typeface + size + color.
            for rpr in root.iter(f"{{{_NS_A}}}rPr"):
                typeface = rpr.get("typeface")
                sz = rpr.get("sz")
                # Determine tier by font size.
                tier = "title"
                if sz:
                    try:
                        tier = "title" if int(sz) >= _TITLE_SIZE_THRESHOLD else "body"
                    except ValueError:
                        pass
                if typeface:
                    typeface_by_tier[tier][typeface] += 1
                # Collect srgbClr children (text-run colors).
                for clr in rpr.iter(f"{{{_NS_A}}}srgbClr"):
                    val = clr.get("val", "").upper()
                    if val and len(val) == 6:
                        all_colors.append(val)
            # Walk shape-level fill colors ONLY (avoid double-counting rPr colors).
            # Use p:spPr namespace to target shape properties, not text runs.
            _NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
            for spPr in root.iter(f"{{{_NS_P}}}spPr"):
                for clr in spPr.iter(f"{{{_NS_A}}}srgbClr"):
                    val = clr.get("val", "").upper()
                    if val and len(val) == 6:
                        all_colors.append(val)

    major_font = typeface_by_tier["title"].most_common(1)[0][0] if typeface_by_tier["title"] else None
    minor_font = typeface_by_tier["body"].most_common(1)[0][0] if typeface_by_tier["body"] else None

    return _ScavengedTheme(
        major_font=major_font,
        minor_font=minor_font,
        colors=all_colors,
    )


def _build_synthetic_theme(
    base_theme_bytes: bytes, scavenged: _ScavengedTheme
) -> bytes:
    """Override default theme's clrScheme + fontScheme with scavenged values.

    Keeps the base theme's fmtScheme and other elements intact (structural
    validity), only replacing the color palette and font names.

    **Fidelity ceiling (m-7):** only accent1–accent6 are overridden from
    scavenged colors. The background/text roles (dk1/lt1/dk2/lt2) keep
    default.pptx's values — a distinctive user background is NOT recovered.
    This is the accepted best-effort limit of Level 2.
    """
    root = etree.parse(BytesIO(base_theme_bytes)).getroot()
    elements = root.find(f"{{{_NS_A}}}themeElements")
    if elements is None:
        return base_theme_bytes  # can't override; return default as-is

    # --- Override fontScheme ---
    if scavenged.major_font or scavenged.minor_font:
        font_scheme = elements.find(f"{{{_NS_A}}}fontScheme")
        if font_scheme is not None:
            if scavenged.major_font:
                major = font_scheme.find(f"{{{_NS_A}}}majorFont")
                if major is not None:
                    latin = major.find(f"{{{_NS_A}}}latin")
                    if latin is not None:
                        latin.set("typeface", scavenged.major_font)
            if scavenged.minor_font:
                minor = font_scheme.find(f"{{{_NS_A}}}minorFont")
                if minor is not None:
                    latin = minor.find(f"{{{_NS_A}}}latin")
                    if latin is not None:
                        latin.set("typeface", scavenged.minor_font)

    # --- Override clrScheme (map top colors to accent roles) ---
    if scavenged.colors:
        color_counts = Counter(scavenged.colors)
        # Exclude pure black/white (they're dk1/lt1 defaults).
        non_mono = [c for c, _ in color_counts.most_common(12) if c not in ("000000", "FFFFFF")]
        clr_scheme = elements.find(f"{{{_NS_A}}}clrScheme")
        if clr_scheme is not None and non_mono:
            accent_roles = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
            for i, role in enumerate(accent_roles):
                if i < len(non_mono):
                    elem = clr_scheme.find(f"{{{_NS_A}}}{role}")
                    if elem is not None:
                        srgb = elem.find(f"{{{_NS_A}}}srgbClr")
                        if srgb is not None:
                            srgb.set("val", non_mono[i])

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ---------------------------------------------------------------------------
# Derived file builder
# ---------------------------------------------------------------------------

def _build_repaired_file(
    template_path: str,
    default_template_path: str,
    theme_bytes: Optional[bytes],
    level: str = "L3",
) -> str:
    """Copy ``default.pptx`` to a derived file; optionally replace the theme.

    The derived file is named ``<stem>_repaired.pptx`` beside the original.
    If ``theme_bytes`` is provided (Level 1/2), the derived file's
    ``ppt/theme/theme1.xml`` is replaced with the salvaged/scavenged content.
    Level 3 (``theme_bytes is None``) uses default.pptx verbatim.
    """
    src = Path(template_path)
    out_dir = src.parent
    repaired_name = f"{src.stem}_repaired{src.suffix}"
    repaired_path = out_dir / repaired_name

    fd, tmp_name = tempfile.mkstemp(suffix=".pptx", dir=str(out_dir))
    os.close(fd)
    try:
        if theme_bytes is not None:
            # Copy default.pptx, replacing the theme entry.
            with zipfile.ZipFile(default_template_path, "r") as zin, \
                    zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    data = zin.read(info.filename)
                    if info.filename == _THEME_PATH:
                        data = theme_bytes  # replace theme
                    out_info = zipfile.ZipInfo(filename=info.filename, date_time=info.date_time)
                    out_info.compress_type = info.compress_type
                    out_info.external_attr = info.external_attr
                    zout.writestr(out_info, data)
        else:
            # Level 3: just copy default.pptx verbatim.
            shutil.copy2(default_template_path, tmp_name)
        os.replace(tmp_name, str(repaired_path))
    except Exception:
        if os.path.exists(tmp_name):
            os.remove(tmp_name)
        raise

    logger.info("Repaired template: %s (level=%s)", repaired_path, level)
    return str(repaired_path)
