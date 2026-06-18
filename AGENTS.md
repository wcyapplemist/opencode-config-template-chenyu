# Project-Specific Agent Instructions

## Project Overview

PPTX subagent development — iterating and testing the `pptx-subagent` and `ppt-template-filler` skill.

## Project Structure

```
pptx-subagent-development/
├── .opencode/
│   ├── agents/
│   │   └── pptx-subagent.md       # Project-level PPT subagent (multi-stage workflow)
│   └── skills/
│       └── ppt-template-filler/   # Template filling engine + SKILL.md
│           ├── scripts/
│           │   ├── ppt_builder.py          # Engine: layouts, charts, images
│           │   ├── schema_validator.py      # JSON schema validation + retry (#20)
│           │   ├── schemas/                 # Per-slide-type schema definitions
│           │   ├── resolvers/               # Resource resolution pipeline (#23)
│           │   ├── outline_store.py         # Multi-stage outline artifact (#21/#24)
│           │   └── tests/                   # pytest suite (120 tests)
│           └── docs/                        # DESIGN-*.md architecture docs
├── output/                        # Generated .pptx files
└── AGENTS.md                      # This file
```

## Project-Level Resources

| Resource              | Type  | Scope             |
| --------------------- | ----- | ----------------- |
| `pptx-subagent`       | Agent | This project only |
| `ppt-template-filler` | Skill | This project only |

Global subagents and skills are managed at `~/.config/opencode/` and are available in all projects.

## Development Notes

- The `pptx-subagent` uses `ppt_builder.py` from the `ppt-template-filler` skill to populate `template.pptx` layouts
- Generated files are saved to `output/`
- The subagent is STRICTLY FORBIDDEN from building PPTX files from scratch

## Phase 1: Content Intelligence & Resource Resolution (issues #17–#25)

The engine layers content-intelligence on top of the python-pptx renderer (output stays 100% native/editable):

- **Schema validation (#20, P0)** — `schema_validator.py` validates all 8 slide types + `chart_options`; structured errors; two-layer retry (`parse_and_validate`). The engine raises a clear `ValidationError` on unrecoverable structure and degrades gracefully otherwise; `strict=True` blocks on any schema violation (agent pre-flight gate).
- **Resource pipeline (#19/#18/#23)** — placeholders (`image_prompt`, `icon_query`, `data_query`) → `resolvers/` (image/icon/chart-data) → concrete assets before render. All resolution is non-fatal.
- **Multi-stage generation (#21/#24)** — outline → critique → detail, schema-gated per stage; autonomous by default, optional interactive checkpoint in primary-agent mode.

Run the suite from `.opencode/skills/ppt-template-filler/scripts`:

```bash
python -m pytest tests/ -q
```
