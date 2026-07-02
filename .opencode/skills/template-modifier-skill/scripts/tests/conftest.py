"""Shared fixtures for template-modifier-skill tests."""
import sys
from pathlib import Path

import pytest

_HERE = Path(__file__).resolve().parent            # .../template-modifier-skill/scripts/tests
_MODIFIER_SCRIPTS = _HERE.parent                   # .../template-modifier-skill/scripts
_SKILLS = _MODIFIER_SCRIPTS.parent.parent          # .../skills
_FILLER_SCRIPTS = _SKILLS / "generate-slide-skill" / "scripts"

for _p in (str(_MODIFIER_SCRIPTS), str(_FILLER_SCRIPTS)):
    if _p not in sys.path:
        sys.path.insert(0, _p)


@pytest.fixture
def template_path():
    return str(_FILLER_SCRIPTS / "templates" / "template.pptx")
