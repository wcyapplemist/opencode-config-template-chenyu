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
│           │   ├── template_introspector.py # Fingerprint-contract extraction (renderer-side)
│           │   ├── schema_extractor.py      # US-1.1: normalized proposed-schema extraction (parallel, non-invasive)
│           │   ├── schema_validator.py      # JSON schema validation + retry (#20)
│           │   ├── schemas/                 # Per-slide-type schema definitions + template_schema.json (US-1.1 extraction spec)
│           │   ├── resolvers/               # Resource resolution pipeline (#23)
│           │   ├── outline_store.py         # Multi-stage outline artifact (#21/#24)
│           │   └── tests/                   # pytest suite (156 tests)
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
- **Resource pipeline (#19/#18/#23)** — placeholders (`data_query`) → `resolvers/` (chart-data) → concrete assets before render. All resolution is non-fatal.
- **Multi-stage generation (#21/#24)** — outline → critique → detail, schema-gated per stage; autonomous by default for headless subagents.
- **Density modes (text-overflow prevention)** — `density_mode.py` fixes a per-slide visible-text word budget per mode (`concise` 0–10 / `standard` 30–50 / `text-heavy` 75–150). The validator emits non-fatal warnings on out-of-budget slides (`validate_slide_data_list(..., density_mode=...)` / `parse_and_validate(..., density_mode=...)`); warnings never block, even in strict mode. This is the content-side defense against text overflowing placeholder boundaries.

  **MANDATORY outline + density-mode checkpoint (Stage 1 → confirm → Stage 3+):** When you (the primary conversation agent) handle a PPT task **directly** — i.e. you did not delegate it to a headless subagent via the Task tool — you **MUST** pause after producing the Stage 1 outline. In a **single `question` call**, ask the user BOTH (a) the density mode (`standard` recommended default) AND (b) outline approval/edits, then **wait for both answers before proceeding** to Stage 3 (JSON) or rendering. Never run outline → detail in one shot when a live user turn-loop is available. Only subagents (which cannot pause) run fully autonomously — they default to `standard` and self-apply the budget.

Run the suite from `.opencode/skills/ppt-template-filler/scripts`:

```bash
python -m pytest tests/ -q
```
