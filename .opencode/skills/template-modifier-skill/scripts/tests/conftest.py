"""Shared fixtures for template-modifier-skill tests."""
import sys
from pathlib import Path

import pytest

_HERE = Path(__file__).resolve().parent            # .../template-modifier-skill/scripts/tests
_MODIFIER_SCRIPTS = _HERE.parent                   # .../template-modifier-skill/scripts
_SKILLS = _MODIFIER_SCRIPTS.parent.parent          # .../skills
_FILLER_SCRIPTS = _SKILLS / "generate-slide-skill" / "scripts"
# PLAN-GIT-72: shared contract/extraction infra now in _common.
_COMMON_SCRIPTS = _SKILLS / "_common" / "scripts"

# _FILLER_SCRIPTS stays: tests/test_layout_creator.py imports generate_ppt_from_data
# (the fill entry) for end-to-end clone renders. _COMMON_SCRIPTS added for the
# shared contract/schema modules.
for _p in (str(_MODIFIER_SCRIPTS), str(_COMMON_SCRIPTS), str(_FILLER_SCRIPTS)):
    if _p not in sys.path:
        sys.path.insert(0, _p)


@pytest.fixture
def template_path():
    return str(_FILLER_SCRIPTS / "templates" / "template.pptx")
