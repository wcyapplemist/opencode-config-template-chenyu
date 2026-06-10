"""
ppt_builder.py
==============
PPT engine using template.pptx Slide Master layouts with proper placeholders.

Loads the template, adds new slides from named layouts, fills placeholders
by type (TITLE, SUBTITLE, OBJECT), and saves the result.

Usage:
    from ppt_builder import generate_ppt_from_data, DEFAULT_OUTPUT_DIR

    result = generate_ppt_from_data(
        slide_data_list,
        output_path=str(DEFAULT_OUTPUT_DIR / "report.pptx"),
    )
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Pt

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

_SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = _SCRIPT_DIR / "templates"
DEFAULT_OUTPUT_DIR = _SCRIPT_DIR.parent / "output"

_TEMPLATE_FILE = TEMPLATES_DIR / "template.pptx"

_TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
_SUBTITLE_TYPE = PP_PLACEHOLDER.SUBTITLE
_BODY_TYPE = PP_PLACEHOLDER.BODY
_OBJECT_TYPE = PP_PLACEHOLDER.OBJECT

_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_EXTENDED_LAYOUT_MAP: Dict[str, int] = {
    "title_slide": 0,
    "agenda_slide": 1,
    "section_header_slide": 2,
    "section_header_sub_image_slide": 3,
    "content_image_slide": 4,
    "section_header_sub_slide": 5,
    "two_content_dark_slide": 6,
    "two_content_white_slide": 7,
    "content_slide": 8,
    "content_table_slide": 9,
    "content_2_slide": 10,
    "content_dark_slide": 11,
    "closing_slide": 12,
}

_LAYOUTS_WITH_SUBTITLE = {
    "title_slide", "agenda_slide", "closing_slide",
    "section_header_sub_slide",
}
_LAYOUTS_WITH_BODY = {
    "content_slide", "content_image_slide",
    "content_2_slide", "content_table_slide",
}
_LAYOUTS_WITH_BODY_AS_SUBTITLE = {
    "section_header_sub_image_slide",
}
_LAYOUTS_WITH_TWO_BODIES = {
    "two_content_dark_slide", "two_content_white_slide",
    "content_dark_slide",
}


def _load_config() -> Dict[str, Any]:
    config_path = TEMPLATES_DIR / "template.config.json"
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _remove_all_slides(prs: Presentation) -> int:
    count = len(prs.slides)
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].attrib.get(f"{{{_REL_NS}}}id")
        if rId:
            prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
    return count


def _find_placeholder(slide: Any, ph_type: Any) -> Optional[Any]:
    for ph in slide.placeholders:
        if ph.placeholder_format.type == ph_type:
            return ph
    return None


def _find_placeholders(slide: Any, ph_type: Any) -> List[Any]:
    return [ph for ph in slide.placeholders if ph.placeholder_format.type == ph_type]


def _find_title_placeholder(slide: Any) -> Optional[Any]:
    for ph_type in _TITLE_TYPES:
        ph = _find_placeholder(slide, ph_type)
        if ph:
            return ph
    return None


def _find_body_placeholder(slide: Any) -> Optional[Any]:
    ph = _find_placeholder(slide, _BODY_TYPE)
    if ph:
        return ph
    objects = _find_placeholders(slide, _OBJECT_TYPE)
    return objects[0] if objects else None


def _set_text(shape: Any, text: str) -> bool:
    if not shape or not shape.has_text_frame:
        return False
    try:
        tf = shape.text_frame
        tf.clear()
        tf.paragraphs[0].text = text
        return True
    except Exception as exc:
        logger.warning("Failed to set text: %s", exc)
        return False


def _parse_line(line: str) -> Tuple[str, str]:
    clean = re.sub(r"\*\*", "", line.strip())
    if not clean:
        return ("", "")
    for sep in [" \u2014 ", " - ", ": "]:
        if sep in clean:
            parts = clean.split(sep, 1)
            return (parts[0].strip(), parts[1].strip())
    return (clean, "")


def _set_body_text(shape: Any, text: str) -> bool:
    if not shape or not shape.has_text_frame:
        return False
    try:
        lines = [l for l in text.split("\n") if l.strip()]
        tf = shape.text_frame
        tf.clear()

        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            title_part, desc_part = _parse_line(line)

            if title_part:
                run = p.add_run()
                run.text = title_part
                run.font.bold = True
                run.font.size = Pt(14)
            if desc_part:
                run = p.add_run()
                run.text = f" \u2014 {desc_part}" if title_part else desc_part
                run.font.size = Pt(12)

        return True
    except Exception as exc:
        logger.warning("Failed to set body text: %s", exc)
        return False


def generate_ppt_from_data(
    slide_data_list: List[Dict[str, Any]],
    template_path: Optional[str] = None,
    output_path: str = "output.pptx",
    prompt_text: str = "",
) -> str:
    template = Path(template_path) if template_path and template_path != "auto" else _TEMPLATE_FILE
    output = Path(output_path)

    if not template.exists():
        raise FileNotFoundError(f"Template not found: {template}")

    if not output.is_absolute():
        output = DEFAULT_OUTPUT_DIR / output
    output.parent.mkdir(parents=True, exist_ok=True)

    config = _load_config()
    title_layout_idx = config.get("title_slide_layout", 0)
    content_layout_idx = config.get("content_slide_layout", 8)

    logger.info("Loading template: %s", template.name)
    prs = Presentation(str(template))
    logger.info("Template: %d slides, %d layouts", len(prs.slides), len(prs.slide_layouts))

    removed = _remove_all_slides(prs)
    logger.info("Cleared %d example slides", removed)

    for page_num, slide_data in enumerate(slide_data_list, start=1):
        slide_type = slide_data.get("slide_type", "")

        # Determine layout index
        if slide_type in _EXTENDED_LAYOUT_MAP:
            layout_idx = _EXTENDED_LAYOUT_MAP[slide_type]
        elif slide_type == "title_slide":
            layout_idx = min(title_layout_idx, len(prs.slide_layouts) - 1)
        elif slide_type == "content_slide":
            layout_idx = min(content_layout_idx, len(prs.slide_layouts) - 1)
        else:
            logger.warning("Page %d: unknown slide_type '%s', skipped", page_num, slide_type)
            continue

        try:
            layout = prs.slide_layouts[layout_idx]
            slide = prs.slides.add_slide(layout)
            logger.info("Page %d: added slide from layout[%d] '%s'", page_num, layout_idx, layout.name)

            title_text = slide_data.get("title", "")

            # Always try to fill title placeholder
            if title_text:
                title_ph = _find_title_placeholder(slide)
                if title_ph:
                    _set_text(title_ph, title_text)
                    logger.info("  Title: \"%s\"", title_text)

            # Fill subtitle (for title, agenda, closing, section-header-sub)
            if slide_type in _LAYOUTS_WITH_SUBTITLE:
                subtitle_text = slide_data.get("subtitle", "")
                if subtitle_text:
                    sub_ph = _find_placeholder(slide, _SUBTITLE_TYPE)
                    if sub_ph:
                        _set_text(sub_ph, subtitle_text)
                        logger.info("  Subtitle: \"%s\"", subtitle_text[:50])

            # Fill body text (for content slides)
            if slide_type in _LAYOUTS_WITH_BODY:
                body_text = slide_data.get("body", "")
                if body_text:
                    body_ph = _find_body_placeholder(slide)
                    if body_ph:
                        _set_body_text(body_ph, body_text)
                        logger.info("  Body: %d lines", len([l for l in body_text.split("\n") if l.strip()]))

            # Fill body as subtitle (section_header_sub_image uses BODY placeholder for description)
            if slide_type in _LAYOUTS_WITH_BODY_AS_SUBTITLE:
                subtitle_text = slide_data.get("subtitle", "")
                if subtitle_text:
                    body_ph = _find_body_placeholder(slide)
                    if body_ph:
                        _set_text(body_ph, subtitle_text)
                        logger.info("  Body-as-subtitle: \"%s\"", subtitle_text[:50])

            # Fill two body areas (for two-content slides)
            if slide_type in _LAYOUTS_WITH_TWO_BODIES:
                body_left = slide_data.get("body_left", "")
                body_right = slide_data.get("body_right", "")
                objects = _find_placeholders(slide, _OBJECT_TYPE)
                if len(objects) >= 2:
                    if body_left:
                        _set_body_text(objects[0], body_left)
                        logger.info("  Body-left: %d lines", len([l for l in body_left.split("\n") if l.strip()]))
                    if body_right:
                        _set_body_text(objects[1], body_right)
                        logger.info("  Body-right: %d lines", len([l for l in body_right.split("\n") if l.strip()]))
                elif len(objects) == 1 and (body_left or body_right):
                    _set_body_text(objects[0], body_left or body_right)

        except Exception as exc:
            logger.error("Page %d failed: %s", page_num, exc)

    prs.save(str(output))
    logger.info("Saved: %s (%d slides)", output.resolve(), len(prs.slides))
    return str(output.resolve())


def main() -> None:
    mock: List[Dict[str, Any]] = [
        {"slide_type": "title_slide", "title": "AI Empowering Finance", "subtitle": "2026 Q1"},
        {
            "slide_type": "content_slide",
            "title": "Core AI Scenarios",
            "body": (
                "**Automated Reporting** \u2014 RPA tools auto-generate monthly reports, cutting manual effort by 80%\n"
                "**Smart Reconciliation** \u2014 AI matches bank transactions at 99.5% accuracy\n"
                "**Fraud Detection** \u2014 Real-time anomaly detection with automated alerts\n"
                "**Tax Optimization** \u2014 ML identifies savings opportunities across tax structures"
            ),
        },
        {
            "slide_type": "content_slide",
            "title": "Roadmap",
            "body": (
                "**Phase 1: Pilot** \u2014 Deploy in 2 business units by Q2\n"
                "**Phase 2: Scale** \u2014 Expand to all departments by Q4\n"
                "**Phase 3: Full Deployment** \u2014 Organization-wide adoption by 2027"
            ),
        },
    ]

    print("Test: template.pptx (placeholder-based)")
    result = generate_ppt_from_data(
        mock,
        output_path=str(DEFAULT_OUTPUT_DIR / "test_template.pptx"),
    )
    print(f"Output: {result}")


if __name__ == "__main__":
    main()
