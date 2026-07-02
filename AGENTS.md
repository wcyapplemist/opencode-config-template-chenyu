# Project-Specific Agent Instructions

## Project Overview

PPTX subagent development — iterating and testing the `pptx-subagent` agent plus three skills: `ppt-template-filler` (fill), `template-modifier-skill` (extend), and `generate-template-skill` (extract).

## Project Structure

```
pptx-subagent-development/
├── .opencode/
│   ├── agents/
│   │   └── pptx-subagent.md       # Project-level PPT subagent (multi-stage workflow)
│   └── skills/
│       ├── ppt-template-filler/   # Template filling engine + SKILL.md
│       │   ├── scripts/
│       │   │   ├── ppt_builder.py          # Engine: layouts, charts, images; US-4.1: get_render_contract (embedded-preferred)
│       │   │   ├── contract_adapter.py     # US-4.1: bridge — embedded JSON -> sidecar-shape render contract
│       │   │   ├── template_introspector.py # Fingerprint-contract extraction (sidecar fallback)
    │       │   │   ├── schema_extractor.py      # Epic 1: extraction + font detection + zip embed (US-1.1–1.5); US-3.1: title_source + build_extraction_summary
│       │   │   ├── schema_validator.py      # JSON schema validation + retry (#20)
│       │   │   ├── density_mode.py          # Per-slide word-budget enforcement
│       │   │   ├── text_fit.py              # US-4.2: reactive font auto-shrink estimator (pure)
│       │   │   ├── schemas/                 # Per-slide-type schemas + template_schema.json (Epic 1 spec)
│       │   │   ├── resolvers/               # Resource resolution pipeline (#23)
│       │   │   ├── outline_store.py         # Multi-stage outline artifact (#21/#24)
    │       │   │   └── tests/                   # pytest suite (389 tests; 112 for schema_extractor)
│       │   └── docs/                        # DESIGN-*.md architecture docs
│       ├── generate-template-skill/         # Template extraction + embed (US-3.1; wraps schema_extractor)
│       └── template-modifier-skill/         # Template extension (Capability B)
├── docs/user-stories/              # chenyu-user-stories.md + GAP-ANALYSIS.md (+ .zh.md translations)
├── PLANS/                          # Phased execution plans (PLAN-GIT-48/50/52/54/55/56/58/60/63.md)
├── output/                         # Generated .pptx files
└── AGENTS.md                       # This file
```

## Project-Level Resources

| Resource                  | Type  | Scope             |
| ------------------------- | ----- | ----------------- |
| `pptx-subagent`           | Agent | This project only |
| `ppt-template-filler`     | Skill | This project only |
| `generate-template-skill` | Skill | This project only |

Global subagents and skills are managed at `~/.config/opencode/` and are available in all projects.

## Development Notes

- The `pptx-subagent` uses `ppt_builder.py` from the `ppt-template-filler` skill to populate `template.pptx` layouts
- The `generate-template-skill` extracts a template into JSON and embeds it back (`schema_extractor`); it is a peer of the fill and extend skills, invoked directly by the primary agent for "extract/generate template" requests
- Generated files are saved to `output/`
- The subagent is STRICTLY FORBIDDEN from building PPTX files from scratch

## Epic 1: Template Extraction & JSON Schema (US-1.1–1.5 — COMPLETE)

`schema_extractor.py` extracts a normalized template schema from any `.pptx` and can embed it back into the zip. All 5 Must-Have stories are Met (112 tests in `test_schema_extractor.py`):

- **US-1.1** — `extract_schema()` reads slide master + all layouts → structured JSON conforming to `schemas/template_schema.json`.
- **US-1.2** — `normalize_polygon()` emits 4 normalized `{x,y}` points; `_signed_area()` + winding check (algebraic CCW).
- **US-1.3** — `_classify_shape()` applies the full 10-value type enum + always-emitted `type_confidence`; `"audio"` reachable via OOXML `<a:audioFile>`/`<a:videoFile>`; `shape/low` surfaces a non-fatal WARNING.
- **US-1.4** — `_extract_text_fonts()` populates per-textbox `font` (explicit-only) + nested `runs[]`; deduped `missing_fonts[]` against `_BUILTIN_FONTS` with theme-aware `fallback` (AC4 → ERROR); non-fatal WARNING per missing font (AC3).
- **US-1.5** — `embed_schema()` writes `ppt/template_schema.json` into the PPTX zip via an order-preserving rewrite (`[Content_Types].xml` first + injected `json` Default; idempotent; atomic); `read_embedded_schema()` retrieves it. CLI: `--embed` + `--output-pptx`.

Since US-4.1 the renderer **prefers the embedded JSON** via `get_render_contract` (→ `contract_adapter`), falling back to the sidecar introspection contract (`template_introspector.py`) for legacy/non-templated templates — the two paths coexist (GAP-ANALYSIS §5 Decision 1).

## Epic 3: Template Generator (US-3.1–3.4 — COMPLETE)

A standalone `generate-template-skill` (`.opencode/skills/generate-template-skill/SKILL.md`) extracts any `.pptx` into a normalized schema and returns a self-describing "templated" PPTX with the JSON embedded at `ppt/template_schema.json`. All 4 stories are Met (112 tests in `test_schema_extractor.py`):

- **US-3.1** — `generate-template-skill` orchestrates the full pipeline end-to-end: extract → validate → (title confirm) → embed → return templated PPTX + summary. NL intent routing is via the SKILL.md `description` (extraction verbs) + a one-line "What NOT to Handle" deferral in `pptx-subagent.md` (architecture review MAJOR-1).
- **US-3.2** — `_infer_title` returns a `TitleInference(title, source)` NamedTuple; `_build_metadata` emits `title_source`; the skill prompts the user when `source == "filename"` and always displays the title for confirmation.
- **US-3.3** — the skill returns a downloadable templated PPTX (`embed_schema`) + a human-readable summary (`build_extraction_summary` + CLI `--summary`); the round-trip test (`test_round_trip_deep_equal`) already exists.
- **US-3.4** — `_build_theme()` maps semantic color roles + `font_palette`; sensible defaults on a missing/malformed theme.

Since US-4.1 the renderer **prefers the embedded JSON** via `get_render_contract` (→ `contract_adapter`), falling back to the sidecar introspection contract (`template_introspector.py`) for legacy/non-templated templates (GAP-ANALYSIS §5 Decision 1, Coexist). `title_source` is runtime-enforced by `validate_template_schema` keyed off the shared `TITLE_SOURCES` constant (architecture review MAJOR-2).

## Phase 1: Content Intelligence & Resource Resolution (issues #17–#25)

The engine layers content-intelligence on top of the python-pptx renderer (output stays 100% native/editable):

- **Schema validation (#20, P0)** — `schema_validator.py` validates all 8 slide types + `chart_options`; structured errors; two-layer retry (`parse_and_validate`). The engine raises a clear `ValidationError` on unrecoverable structure and degrades gracefully otherwise; `strict=True` blocks on any schema violation (agent pre-flight gate).
- **Resource pipeline (#19/#18/#23)** — placeholders (`data_query`) → `resolvers/` (chart-data) → concrete assets before render. All resolution is non-fatal.
- **Multi-stage generation (#21/#24)** — outline → critique → detail, schema-gated per stage; autonomous by default for headless subagents.
- **Density modes (text-overflow prevention)** — `density_mode.py` fixes a per-slide visible-text word budget per mode (`concise` 0–10 / `standard` 30–50 / `text-heavy` 75–150). The validator emits non-fatal warnings on out-of-budget slides (`validate_slide_data_list(..., density_mode=...)` / `parse_and_validate(..., density_mode=...)`); warnings never block, even in strict mode. This is the content-side defense against text overflowing placeholder boundaries.
- **Reactive text-fitting (US-4.2, #60)** — `text_fit.py` is a pure heuristic estimator that, at render time, shrinks a placeholder's font in −2pt steps (8pt floor) when text would overflow its box. Base size is template-derived (schema `size_pt` → layout sample-run → conservative role ceiling body 14 / title 28 / subtitle 18); an explicit `run.font.size` is written **only on actual shrink** (else inheritance is preserved); an auto-grow guard skips shrinking on short-base-height placeholders. The body `Pt(14)`/`Pt(12)` hardcode is retired. Per-slide per-placeholder fit decisions (incl. the `font_size_adjusted` flag, AC3) are written to a `<output>.render.json` sidecar (the engine return type is unchanged). **AC1 is best-effort / deferred** — python-pptx has no layout engine, so a hard overflow guarantee needs an external render oracle (see GAP-ANALYSIS §US-4.2 Rev 10).
- **Auto-chain / templated output (US-4.3, #63)** — every generated `.pptx` is **self-describing**: after `prs.save` (which strips the unmodeled part), `generate_ppt_from_data(auto_template=True)` re-embeds `ppt/template_schema.json` into the **output**, sourced from the **input template** (so the schema describes the template, never the rendered deck's cover) and skipping a stale embedded input schema. The agent detects a non-templated input at Stage 0 (`read_embedded_schema`, exception-safe) and emits *"No template found — extracting first, then generating slides..."* (AC3). The output's `<output>.render.json` gains an additive `templating` field. One user prompt → a templated, reusable deck.

  **MANDATORY outline + density-mode checkpoint (Stage 1 → confirm → Stage 3+):** When you (the primary conversation agent) handle a PPT task **directly** — i.e. you did not delegate it to a headless subagent via the Task tool — you **MUST** pause after producing the Stage 1 outline. In a **single `question` call**, ask the user BOTH (a) the density mode (`standard` recommended default) AND (b) outline approval/edits, then **wait for both answers before proceeding** to Stage 3 (JSON) or rendering. Never run outline → detail in one shot when a live user turn-loop is available. Only subagents (which cannot pause) run fully autonomously — they default to `standard` and self-apply the budget.

Run the suite from `.opencode/skills/ppt-template-filler/scripts`:

```bash
python -m pytest tests/ -q
```
