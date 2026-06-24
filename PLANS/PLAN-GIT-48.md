# PLAN-GIT-48 — US-1.1: Extract Slide Master to Structured JSON

**Issue**: #48
**Branch**: GIT-48 (base: dev)
**Priority**: Must Have (P0)
**Status**: Planning

## Goal

Read any .pptx file, parse its slide master and all slide layouts, and emit a
structured JSON representation conforming to the Reference JSON Schema in
`docs/user-stories/chenyu-user-stories.md`. Scripted, deterministic, no LLM
guessing.

## Architecture Decisions (locked)

1. **New parallel module** `schema_extractor.py` — do NOT modify
   `template_introspector.py` / `ppt_builder.py`. The existing introspector
   emits a "fingerprint contract" that the renderer consumes via `get_contract()`;
   its coupling is narrow (`placeholders[]`, `fingerprint`, `slide_size`) so a
   parallel schema-emitting module is zero-risk.
2. **Basic field fill** — polygon (normalized rectangular) + type
   (MSO_SHAPE_TYPE mapping) populated; font left as an empty stub `{}`.
3. **Lightweight validation** — define `template_schema.json` (JSON Schema draft
   2020-12) as the spec, but validate with a self-contained
   `validate_template_schema()`. The project does NOT use the `jsonschema`
   library; follow the existing no-external-deps validator pattern.

## Deliverables (all new files — zero modifications to existing code)

- `scripts/schemas/template_schema.json`
- `scripts/schema_extractor.py`
- `scripts/tests/test_schema_extractor.py`

## Acceptance Criteria (US-1.1)

- [ ] Accepts a .pptx input and does not crash on any valid PPTX.
- [ ] Slide master XML is parsed; every layout under the master is extracted.
- [ ] Output is a valid JSON object conforming to the proposed schema.
- [ ] Extraction is performed by a Python script — not by LLM guessing.

## Implementation Phases

### Phase 1: Schema Definition

- [ ] Task 1: Author `template_schema.json` (draft 2020-12) capturing
      `template_metadata` / `theme` / `slide_layouts[].components[]` /
      `component_type_enum` / `placeholder_type_enum` per the Reference schema.
      Each component object defines `id`, `type` (enum), `name`,
      `placeholder_type` (enum|null), `polygon` (array of `{x,y}`), `z_order`
      (int), `font` (object), `content_template` (string|null).

### Phase 2: Extraction Engine (`schema_extractor.py`)

- [ ] Task 2: Element type mapper — `MSO_SHAPE_TYPE` -> proposed 10-value enum
      (`AUTO_SHAPE`/`FREEFORM` -> `shape`, `PICTURE` -> `image`, `CHART` ->
      `chart`, `TABLE` -> `table`, `GROUP` -> `group`, `TEXT_BOX` -> `textbox`,
      `PLACEHOLDER` -> `placeholder`, `MEDIA` -> `video`/`audio`, else ->
      `shape`). Map `placeholder_format.type` -> `placeholder_type` enum.
- [ ] Task 3: Polygon normalizer (basic) — `left/top/width/height` (EMU) divided
      by slide dimensions -> 4 normalized `{x,y}` points (TL->TR->BR->BL, 0-1).
      Rectangular only; winding validation deferred to US-1.2.
- [ ] Task 4: `extract_schema(pptx_path) -> dict` main function — parse
      `prs.slide_masters[0].shapes` (master components) + each
      `prs.slide_layouts[i].shapes` (all elements, GROUP recurses into
      children). Emit `template_metadata` (title from `docProps/core.xml` ->
      first slide title -> filename; `schema_version`; `generated_at`;
      `slide_dimensions` with EMU/inches/ratio). Emit `theme` (raw theme colors
      -> semantic roles primary/secondary/accent/background). Emit
      self-documenting `component_type_enum` + `placeholder_type_enum`.

### Phase 3: Validation

- [ ] Task 5: `validate_template_schema(schema_dict) -> ValidationResult` —
      structural conformance check (top-level keys, component field types,
      polygon value ranges [0,1], enum legality) without the `jsonschema`
      dependency.

### Phase 4: Testing

- [ ] Task 6: `test_schema_extractor.py` — bundled `template.pptx` extract ->
      `validate_template_schema` passes; non-placeholder elements captured
      (layout 2: >=5 components vs 3 placeholders); master components non-empty
      (3 shapes); polygon values in [0,1], exactly 4 points; group nesting
      recurses; empty/blank layout does not crash.

### Phase 5: CLI & Docs

- [ ] Task 7 (optional): CLI entry `python schema_extractor.py --input X.pptx
      --output schema.json` for standalone invocation (bridges to US-5.1).

## Out of Scope (deferred to sibling stories)

- **US-1.2**: cross-product winding validation, non-rectangular polygon vertices.
- **US-1.3**: full 10-value enum refinement + `type_confidence: "low"` fallback.
- **US-1.4**: font/runs detection (US-1.1 components carry an empty `font: {}`).
- **US-1.5**: zip embedding (US-1.1 only outputs JSON, does not write into the zip).
- **US-3.x**: title inference refinement / downloadable templated PPTX.

## Verified Context (audited)

- python-pptx 1.0.2 available; `MSO_SHAPE_TYPE` exposes AUTO_SHAPE, CHART,
  GROUP, etc.
- Bundled `template.pptx`: 63 layouts, slide master has 3 shapes, layout 2 has 5
  total shapes but only 3 placeholders (2 non-placeholder elements dropped today).
- `jsonschema` is NOT installed; existing validators (`schema_validator.py`) are
  custom/no-external-deps.
- Slide dimensions are already recorded by the existing introspector
  (`_build_slide_size`) — reusable as reference.

## Risks

- python-pptx may not cleanly expose `grpSp`/`graphicFrame` subtypes -> may need
  lxml-level parsing for disambiguation.
- Bundled template is a Google Slides export — element naming is non-standard
  ("Google Shape;13;p2"); `name` preserved as-is, mapping must not depend on it.
- SmartArt (`IGX_GRAPHIC`) degrades to `shape`.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` (Epic 1, US-1.1 + Reference JSON Schema).
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` (section 2, US-1.1).
- Issue: #48.
