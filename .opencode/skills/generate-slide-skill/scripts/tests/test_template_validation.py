"""US-4.7 — Template selection & pre-render validation tests."""
import shutil
from pathlib import Path

import pytest
from pptx import Presentation

from ppt_builder import (
    TemplateError,
    _TEMPLATE_FILE,
    _validate_template,
    generate_ppt_from_data,
)


class _FakePrs:
    """Minimal stand-in for a python-pptx Presentation for _validate_template.

    Only the two surfaces _validate_template touches are needed: an iterable
    ``slide_masters`` and a sized ``slide_layouts``.
    """

    def __init__(self, masters, layouts):
        self.slide_masters = masters
        self.slide_layouts = layouts


_MIN_DECK = [
    {"slide_type": "title_slide", "title": "T", "subtitle": "S", "notes": "n"},
]


# ---------------------------------------------------------------------------
# Default-constant sanity (AC1 setup)
# ---------------------------------------------------------------------------
class TestDefaultConstant:
    def test_default_points_to_repo_root_default(self):
        assert _TEMPLATE_FILE.parent.name == "template"
        assert _TEMPLATE_FILE.name == "default.pptx"
        assert _TEMPLATE_FILE.exists()


# ---------------------------------------------------------------------------
# _validate_template unit tests (AC4 / AC5 / AC6)
# ---------------------------------------------------------------------------
class TestValidateTemplateUnit:
    def test_no_master_is_fatal(self):
        prs = _FakePrs(masters=[], layouts=["x"])  # layouts present, no master
        with pytest.raises(TemplateError, match="no slide master"):
            _validate_template(prs, "fake.pptx", None)

    def test_zero_layouts_is_fatal(self):
        prs = _FakePrs(masters=["m"], layouts=[])
        with pytest.raises(TemplateError, match="no slide layouts"):
            _validate_template(prs, "fake.pptx", None)

    def test_serves_no_slide_types_is_fatal(self, monkeypatch):
        prs = _FakePrs(masters=["m"], layouts=["x"])
        monkeypatch.setattr(
            "ppt_builder.servable_slide_types",
            lambda contract: {"title_slide": {"available": False, "reason": "none"}},
        )
        with pytest.raises(TemplateError, match="none of the 8 slide types"):
            _validate_template(prs, "fake.pptx", {"layouts": []})

    def test_contract_none_skips_servability(self):
        # master + layouts present, contract unavailable -> no fatal (AC6 skipped).
        prs = _FakePrs(masters=["m"], layouts=["x"])
        _validate_template(prs, "fake.pptx", None)  # must not raise

    def test_valid_template_passes(self, monkeypatch):
        prs = _FakePrs(masters=["m"], layouts=["x"])
        monkeypatch.setattr(
            "ppt_builder.servable_slide_types",
            lambda contract: {"title_slide": {"available": True}},
        )
        _validate_template(prs, "fake.pptx", {"layouts": [{}]})  # must not raise


# ---------------------------------------------------------------------------
# Integration via generate_ppt_from_data (AC1 / AC2 / AC3 / AC7)
# ---------------------------------------------------------------------------
class TestGenerateTemplateValidation:
    def test_default_used_when_path_omitted(self, tmp_path):
        # AC1: no template_path -> renders against template/default.pptx.
        out = str(tmp_path / "out.pptx")
        result = generate_ppt_from_data(_MIN_DECK, output_path=out)
        assert Path(result).exists()

    def test_default_used_when_path_is_auto(self, tmp_path):
        # AC1 variant: the "auto" sentinel also falls back to the default.
        out = str(tmp_path / "out.pptx")
        result = generate_ppt_from_data(_MIN_DECK, template_path="auto", output_path=out)
        assert Path(result).exists()

    def test_user_path_renders_and_default_untouched(self, template_path, tmp_path):
        # AC2: a user template path is used; the default file is never modified.
        default = Path(template_path)
        user_copy = tmp_path / "user.pptx"
        shutil.copy2(default, user_copy)
        before = default.read_bytes()

        out = str(tmp_path / "out.pptx")
        generate_ppt_from_data(_MIN_DECK, template_path=str(user_copy), output_path=out)

        assert Path(out).exists()
        assert default.read_bytes() == before  # default untouched

    def test_corrupt_file_raises_template_error(self, tmp_path):
        # AC3: a corrupt / non-PPTX file -> TemplateError, no raw traceback.
        bad = tmp_path / "broken.pptx"
        bad.write_bytes(b"this is definitely not a pptx zip")
        out = str(tmp_path / "out.pptx")
        with pytest.raises(TemplateError, match="Could not open template as PPTX"):
            generate_ppt_from_data(
                _MIN_DECK, template_path=str(bad), output_path=out
            )

    def test_missing_template_file_raises(self, tmp_path):
        # Non-existent path -> FileNotFoundError (pre-existing) before validation.
        out = str(tmp_path / "out.pptx")
        with pytest.raises(FileNotFoundError):
            generate_ppt_from_data(
                _MIN_DECK, template_path=str(tmp_path / "nope.pptx"), output_path=out
            )

    def test_non_templated_template_still_renders(self, template_path, tmp_path):
        # AC7: a template lacking an embedded schema still renders (non-fatal).
        plain = tmp_path / "plain.pptx"
        Presentation(template_path).save(str(plain))  # drops embedded json on re-save
        out = str(tmp_path / "out.pptx")
        generate_ppt_from_data(_MIN_DECK, template_path=str(plain), output_path=out)
        assert Path(out).exists()
