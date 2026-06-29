# PPTX Subagent Development

Turn a one-line request into a fully editable `.pptx` deck — native text, native
charts, and embedded pictures, never screenshots — by filling a Slide Master
template instead of building slides by hand.

## What this project does

A system for generating and templating PowerPoint presentations, built on three project-level skills and one agent:

- A **content strategist** (the `pptx-subagent` agent) turns a plain-language request into
  structured slide content (a JSON array).
- A **rendering engine** (`ppt-template-filler` / `ppt_builder.py`) fills a `template.pptx` Slide Master
  with that content and writes a `.pptx` to `output/`.
- A **template extractor** (`generate-template-skill` / `schema_extractor.py`) reads any `.pptx`,
  emits a normalized JSON schema, and returns a self-describing "templated" PPTX with that JSON embedded.

Everything in the output is a native, editable PowerPoint object (text runs,
charts, pictures) — double-click a chart and it opens live in PowerPoint.

## Architecture

```
                       ┌──────────────────────────────────────┐
   "Make a 5-page      │  pptx-subagent  (content layer)      │   LLM agent
    deck about X" ───▶ │  Stage 1 outline                      │   Produces JSON,
                       │  Stage 2 density mode + critique      │   never rendering code
                       │  Stage 3 detail + schema validation   │
                       └──────────────────┬───────────────────┘
                                          │  slide_data_list  (JSON array)
                                          ▼
                       ┌──────────────────────────────────────┐
                       │  ppt-template-filler (render layer)   │   python-pptx engine
                       │  resolve_slide_data_list()            │   The ONLY entry
                        │    └─ chart-data resolver              │   point that writes .pptx
                       │  generate_ppt_from_data()             │
                       └──────────────────┬───────────────────┘
                                          │
                                          ▼
                               output/<deck>.pptx   (100% native, editable)
```

> **Template extraction is a separate path:** `generate-template-skill` reads any `.pptx` →
> `schema_extractor` JSON → embeds it back as `ppt/template_schema.json` (a "templated" PPTX).
> It does not render slides; it only extracts and packages a template definition. A third skill,
> `template-modifier-skill`, clones extended layouts when content exceeds the base template.

### Separation of concerns (why the strategist never calls python-pptx directly)

The content layer only emits structured JSON; it does **not** write `python-pptx`
code such as `Presentation()` or `prs.slides.add_slide()`. Rendering is routed
exclusively through `ppt_builder.py`. This keeps a mandatory quality pipeline —
schema validation, resource resolution, and per-slide density control — that the
content cannot bypass. Skipping the engine would mean losing every safety check
and producing slides detached from the Slide Master's design.

## Repository layout

```
pptx-subagent-development/
├── .opencode/
│   ├── agents/
│   │   └── pptx-subagent.md              # Content-strategist agent (Stage 0–5 workflow)
│   └── skills/
│       ├── ppt-template-filler/          # Template filling engine + SKILL.md
│       │   ├── SKILL.md                   # Engine usage contract
│       │   ├── docs/                      # DESIGN-*.md architecture deep-dives
│       │   └── scripts/
│       │       ├── ppt_builder.py            # ← THE renderer (only .pptx writer)
│       │       ├── template_introspector.py   # Fingerprint-contract extraction (renderer-side)
│       │       ├── schema_extractor.py        # Epic 1: extraction + font detection + zip embed (US-1.1–1.5)
│       │       ├── schema_validator.py        # JSON schema validation + retry
│       │       ├── density_mode.py            # Per-slide word-budget enforcement
│       │       ├── outline_store.py           # Outline checkpoint artifact
│       │       ├── templates/
│       │       │   ├── template.pptx          # Slide Master with named layouts
│       │       │   └── template.config.json   # Layout-name overrides
│       │       ├── resolvers/                  # chart-data resolver
│       │       ├── schemas/                    # Per-slide-type JSON schemas + template_schema.json (Epic 1 spec)
    │       │       └── tests/                      # pytest suite (112 tests for schema_extractor alone)
    │       ├── generate-template-skill/    # Template extraction + embed (US-3.1; wraps schema_extractor)
    │       └── template-modifier-skill/      # Template extension (Capability B)
├── docs/                                 # Activity diagrams, models, use-cases, workflows
│   └── user-stories/                     # chenyu-user-stories.md + GAP-ANALYSIS.md (+ .zh.md)
├── PLANS/                                # Phased execution plans (PLAN-GIT-48/50/52/54/55/56.md)
├── output/                               # Generated .pptx files (gitignored)
├── chenyu-user requirement.html          # Original requirements source (HTML)
├── requirements.txt                      # Python dependencies
└── AGENTS.md                             # Agent operating rules
```

## Prerequisites

- Python 3.9+
- Install dependencies:

  ```bash
  pip install -r requirements.txt
  ```

## Quick start

Run the engine directly on a JSON array — the fastest way to verify the install:

```bash
python -c "
import sys; sys.path.insert(0, '.opencode/skills/ppt-template-filler/scripts')
from ppt_builder import generate_ppt_from_data, DEFAULT_OUTPUT_DIR

slide_data = [
  {
    'slide_type': 'title_slide',
    'title': 'Hello Deck',
    'subtitle': 'Quick-start demo',
    'notes': 'KEY MESSAGE: A working deck in one command.\nTRANSITION: That is all.\nCOACHING: Confident open.'
  },
  {
    'slide_type': 'content_slide',
    'title': 'What you get',
    'body': '**Native text** - every word editable\n**Live charts** - double-click to edit\n**No screenshots**',
    'notes': 'KEY MESSAGE: Output is editable, not flattened.\nTRANSITION: Done.\nCOACHING: Keep it brief.'
  },
  {
    'slide_type': 'closing_slide',
    'title': 'Thank You',
    'notes': 'KEY MESSAGE: Clean sign-off.\nTRANSITION: Open for questions.\nCOACHING: Smile.'
  }
]

result = generate_ppt_from_data(slide_data, output_path=str(DEFAULT_OUTPUT_DIR / 'demo.pptx'))
print(result)
"
```

Open the printed path in PowerPoint. Text and charts are fully editable.

**Alternative — natural language:** if you use OpenCode, just ask
`"Create a 5-page PPT about ..."` and the `pptx-subagent` agent runs the full
multi-stage pipeline (outline → critique → detail → render) for you. To turn an
existing deck into a reusable template, ask `"Extract the template from this PPTX"`
and the `generate-template-skill` produces a self-describing templated `.pptx`.

## How it works (pipeline)

```
Stage 1  Outline        plain-text plan, one line per slide
Stage 2  Density + gate pick word budget (standard 30–50 / concise 0–10 / text-heavy 75–150);
                        approve or self-critique the outline
Stage 3  Detail + JSON  full slide_data_list, schema-validated, density-aware
Stage 4  Resolve+Render resolvers fill placeholders, then generate_ppt_from_data()
Stage 5  Return         absolute path to the .pptx
```

See `.opencode/agents/pptx-subagent.md` for the full stage contract.

## Template-schema extraction (Epic 1) + Template Generator skill (Epic 3)

`schema_extractor.py` reads any `.pptx` and emits a normalized JSON schema that
mirrors the slide master + layouts. **Epic 1 (5 stories) is fully implemented:**

| Story | Capability |
|---|---|
| US-1.1 | Slide master + all layouts → structured JSON (`extract_schema`) |
| US-1.2 | Normalized polygon positioning with winding check |
| US-1.3 | Component type enum + `type_confidence` + audio/video subtype |
| US-1.4 | Per-textbox font detection + `missing_fonts` + availability check |
| US-1.5 | Embed the schema into the PPTX zip at `ppt/template_schema.json` |

**Epic 3 (4 stories) is also complete** — the `generate-template-skill` wraps the engine
into a full `extract → validate → (title confirm) → embed → return templated PPTX + summary`
pipeline (US-3.1). It surfaces `title_source` provenance (`core_xml`/`slide1`/`filename`/`user`,
US-3.2) and a human-readable extraction summary (US-3.3). Ask "extract the template from this
PPTX" in natural language to invoke it.

**CLI:**
```bash
python schema_extractor.py --input template.pptx --output schema.json        # extract only
python schema_extractor.py --input template.pptx --output schema.json --embed # extract + embed into a .pptx copy
python schema_extractor.py --input template.pptx --output schema.json --embed --summary # + print a human-readable summary
```

The schema is validated by `validate_template_schema()` (hand-rolled, no
`jsonschema` dependency); `title_source` is additionally runtime-enforced via an enum check
keyed off the shared `TITLE_SOURCES` constant. The renderer's fingerprint contract
(`template_introspector.py`) is **untouched** — the two paths coexist (GAP-ANALYSIS
§5 Decision 1).

## Extending the engine

**Add a new slide type** — three places, all under `scripts/`:

1. `schemas/slide_schemas.py` — define the JSON schema for the new type.
2. `ppt_builder.py` — add the type to `_LAYOUT_NAME_MAP` and a render branch.
3. `schemas/__init__.py` — export it from `VALID_SLIDE_TYPES`.

**Add a resource resolver** — implement a function with the
`(slide_data, config) -> slide_data` signature in `scripts/resolvers/`, then
register it in `resolvers/pipeline.py`. The existing chart-data resolver
are the templates to copy. Resolvers must degrade gracefully: a failed fetch logs
a warning and never aborts the build.

## Supported slide types

| `slide_type` | Purpose | Key fields |
|---|---|---|
| `title_slide` | Cover | `title`, `subtitle`, `notes` |
| `content_slide` | Content | `title`, `body`, `notes` |
| `section_header_slide` | Divider | `title`, `notes` |
| `two_content_slide` | Two-column | `title`, `body_left`, `body_right`, `notes` |
| `comparison_slide` | Comparison | `title`, `body_left`, `body_right`, `notes` |
| `content_image_slide` | Image + caption | `title`, `body`, `image_path`, `notes` |
| `chart_slide` | Native chart | `title`, `chart_type`, `categories`, `series`, `notes` |
| `closing_slide` | Closing | `title` (defaults to `Thank You`), `notes` |

**Chart types:** `bar`, `bar_stacked`, `bar_horizontal`, `bar_horizontal_stacked`,
`pie`, `pie_exploded`, `doughnut`, `line`, `line_markers`.

**Resource placeholder:** `data_query` (chart data) is resolved by the agent's
`webfetch` pre-flight, not the resolver (it never networks); fabricating chart
numbers to pass validation is forbidden. Manual images use `image_path` directly.

## Further reading

| Topic | Document |
|---|---|
| Content-strategist workflow (all stages) | `.opencode/agents/pptx-subagent.md` |
| Engine usage contract & field reference | `.opencode/skills/ppt-template-filler/SKILL.md` |
| Template extraction skill (US-3.1) | `.opencode/skills/generate-template-skill/SKILL.md` |
| Template extension skill (Capability B) | `.opencode/skills/template-modifier-skill/SKILL.md` |
| Multi-stage generation design | `.opencode/skills/ppt-template-filler/docs/DESIGN-multi-stage-generation.md` |
| Resource resolver design | `.opencode/skills/ppt-template-filler/docs/DESIGN-resource-resolver.md` |
| Requirements (user stories) | `docs/user-stories/chenyu-user-stories.md` |
| Gap analysis (implementation status) | `docs/user-stories/GAP-ANALYSIS.md` |
| Phased execution plans | `PLANS/PLAN-GIT-*.md` |

## Scope (project-level resources)

All resources are scoped to this repository only — they are **not** installed
globally:

| Resource | Location |
|---|---|
| `pptx-subagent` | `.opencode/agents/` |
| `ppt-template-filler` | `.opencode/skills/` |
| `generate-template-skill` | `.opencode/skills/` |
| `template-modifier-skill` | `.opencode/skills/` |
