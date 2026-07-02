"""Tests for US-2.1 — header/footer detection + prompt-decision + inject helpers.

Covers:
* Detection on the bundled template (has_header=False, has_footer=True).
* ``needs_header_footer_prompt`` pure helper (both false → True; footer
  present → False).
* ``inject_default_header_zone`` — 4-point polygon + English note + explicit
  polygon assertions (len==4, {x,y}, [0,1] range) + schema validates.
"""
import pytest

from schema_extractor import (
    extract_schema,
    inject_default_header_zone,
    needs_header_footer_prompt,
    validate_template_schema,
)


class TestDetect:
    def test_bundled_has_footer_no_header(self, template_path):
        schema = extract_schema(template_path)
        hf = schema["template_metadata"]["header_footer"]
        assert hf["has_header"] is False
        assert hf["has_footer"] is True

    def test_bundled_does_not_need_prompt(self, template_path):
        schema = extract_schema(template_path)
        # has_footer=True → prompt NOT needed
        assert needs_header_footer_prompt(schema) is False


class TestNeedsPrompt:
    def test_both_false_prompts(self):
        schema = {"template_metadata": {"header_footer": {"has_header": False, "has_footer": False}}}
        assert needs_header_footer_prompt(schema) is True

    def test_footer_present_no_prompt(self):
        schema = {"template_metadata": {"header_footer": {"has_header": False, "has_footer": True}}}
        assert needs_header_footer_prompt(schema) is False

    def test_header_present_no_prompt(self):
        schema = {"template_metadata": {"header_footer": {"has_header": True, "has_footer": False}}}
        assert needs_header_footer_prompt(schema) is False

    def test_empty_header_footer_no_prompt_when_missing(self):
        # no header_footer key at all → treat as absent → prompt
        schema = {"template_metadata": {}}
        assert needs_header_footer_prompt(schema) is True


class TestInject:
    def test_inject_creates_default_zone(self):
        schema = {"template_metadata": {"header_footer": {"has_header": False, "has_footer": False}}}
        inject_default_header_zone(schema)
        header = schema["template_metadata"]["header_footer"]["header"]

        # shape
        assert header["source"] == "user_default"
        assert isinstance(header["note"], str) and len(header["note"]) > 0

        # polygon: exactly 4 points, each {x,y}, all in [0,1]
        poly = header["polygon"]
        assert isinstance(poly, list)
        assert len(poly) == 4
        for pt in poly:
            assert set(pt.keys()) == {"x", "y"}
            assert 0 <= pt["x"] <= 1
            assert 0 <= pt["y"] <= 1

        # winding: TL→TR→BR→BL (top strip)
        assert poly[0] == {"x": 0, "y": 0}
        assert poly[1] == {"x": 1, "y": 0}
        assert poly[2]["y"] > poly[0]["y"]  # bottom edge below top

    def test_inject_idempotent(self):
        schema = {"template_metadata": {"header_footer": {}}}
        inject_default_header_zone(schema)
        first = schema["template_metadata"]["header_footer"]["header"]
        inject_default_header_zone(schema)  # second call
        second = schema["template_metadata"]["header_footer"]["header"]
        assert first is second  # setdefault — no overwrite

    def test_inject_schema_still_validates(self, template_path):
        schema = extract_schema(template_path)
        inject_default_header_zone(schema)
        result = validate_template_schema(schema)
        assert result.is_valid, result.error_messages()
