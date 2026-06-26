# PLAN-GIT-48 — US-1.1: Extract Slide Master to Structured JSON

**Issue**: #48
**Branch**: GIT-48 (base: dev)
**Priority**: Must Have (P0)
**Status**: Implementation complete (v2 — architecture review findings incorporated)

## Strategic Context

This plan implements **GAP-ANALYSIS §5 Decision 1 — option (a) Coexist**: the new
proposed-schema extractor (`schema_extractor.py`) and the existing fingerprint
contract (`template_introspector.py`) will coexist. No bridge or migration is in
scope for US-1.1; the render path (`ppt_builder.py`) continues to consume the
existing contract unchanged.

**GAP-ANALYSIS §5 Decision 2** (does the renderer eventually denormalize polygons
back to EMU and place OOXML manually, or do polygons stay metadata-only?)
remains open. US-1.1 is designed to be valid under either resolution: the
extracted schema is self-contained JSON usable by any future consumer. If
Decision 2 resolves to "polygons drive rendering," a follow-up epic will add the
denormalization consumer in `ppt_builder.py` and plan the introspector
deprecation.

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
2. **Basic field fill — type and polygon only.** polygon (normalized
   rectangular) + type (MSO_SHAPE_TYPE mapping) populated. **Font handling (per
   architecture review C1):** in `template_schema.json` (Task 1) all `font`
   sub-fields are optional (`additionalProperties: false`, no `required`).
   Non-text components (`image`, `shape`, `chart`, `table`, `video`, `audio`,
   `group`) **omit** `font` entirely (matching the Reference schema).
   Text-bearing components (`textbox`, `placeholder`) carry `font` but it may be
   `{}` in US-1.1; US-1.4 populates the sub-fields.
3. **Lightweight validation** — define `template_schema.json` (JSON Schema draft
   2020-12) as the spec, but validate with a self-contained
   `validate_template_schema()`. The project does NOT use the `jsonschema`
   library; follow the existing no-external-deps validator pattern.
4. **Accepted duplication during coexistence (per architecture review M1).**
   `schema_extractor.py` re-walks the same .pptx as `template_introspector.py`,
   duplicating slide-size and theme-parse logic. This is accepted accidental
   complexity for the coexistence period (GAP-ANALYSIS §5 Decision 1).
   Consolidation into shared primitives is deferred to the migration epic.
   (Rejected alternative: extract `_pptx_primitives.py` shared helpers now —
   would relax the zero-modification constraint.)

## Deliverables (all new files — zero modifications to existing code)

- `scripts/schemas/template_schema.json`
- `scripts/schema_extractor.py`
- `scripts/tests/test_schema_extractor.py`

## Acceptance Criteria (US-1.1)

- [x] Accepts a .pptx input and does not crash on any valid PPTX.
- [x] Slide master XML is parsed; every layout under the master is extracted.
- [x] Output is a valid JSON object conforming to `template_schema.json` (the
      schema authored in Task 1).
- [x] Extraction is performed by a Python script — not by LLM guessing.

## Implementation Phases

### Phase 1: Schema Definition

- [x] Task 1: Author `template_schema.json` (draft 2020-12) capturing
      `template_metadata` / `theme` / `slide_layouts[].components[]` /
      `component_type_enum` / `placeholder_type_enum` per the Reference schema.
      Each component object defines `id`, `type` (enum), `name`,
      `placeholder_type` (enum|null), `polygon` (array of `{x,y}`), `z_order`
      (int), `content_template` (string|null). **`font` is an object with
      all-optional properties (no `required`, `additionalProperties: false`);
      it is present only on text-bearing components (`textbox`, `placeholder`)
      and absent on `image` / `shape` / `chart` / `table` / `video` / `audio` /
      `group` (matching the Reference schema).**

### Phase 2: Extraction Engine (`schema_extractor.py`)

- [x] Task 2: Element type mapper — `MSO_SHAPE_TYPE` -> proposed 10-value enum
      (`AUTO_SHAPE`/`FREEFORM` -> `shape`, `PICTURE` -> `image`, `CHART` ->
      `chart`, `TABLE` -> `table`, `GROUP` -> `group`, `TEXT_BOX` -> `textbox`,
      `PLACEHOLDER` -> `placeholder`, `MEDIA` -> `video`/`audio`, else ->
      `shape`). Map `placeholder_format.type` -> `placeholder_type` enum.
- [x] Task 3: Polygon normalizer (basic) — `left/top/width/height` (EMU) divided
      by slide dimensions -> 4 normalized `{x,y}` points (TL->TR->BR->BL, 0-1).
      Rectangular only; winding validation deferred to US-1.2.
- [x] Task 4: `extract_schema(pptx_path) -> dict` main function — parse
      `prs.slide_masters[0].shapes` (master components) + each
      `prs.slide_layouts[i].shapes` (all elements, GROUP recurses into
      children). Component `id` is **globally unique** `comp_NNN` (zero-padded 3
      digits), assigned master-first then layouts in index order, shapes in
      z-order. Emit `template_metadata` (title from `docProps/core.xml` -> first
      slide title -> filename; `schema_version`; `generated_at`;
      `generated_by`: `"opencode-pptx-subagent/schema_extractor"`;
      `slide_dimensions` with EMU/inches/ratio). Emit `theme` (raw theme colors
      -> semantic roles; initial heuristic mapping `primary`<-`dk2`,
      `secondary`<-`lt2`, `accent`<-`accent1`, `background`<-`lt1`,
      `text_color`<-`dk1`; refined in US-3.4). Emit self-documenting
      `component_type_enum` + `placeholder_type_enum`.

### Phase 3: Validation

- [x] Task 5: `validate_template_schema(schema_dict) -> ValidationResult` —
      structural conformance check (top-level keys, component field types,
      polygon value ranges [0,1], enum legality, font-cardinality rule: non-text
      components must NOT carry `font`) without the `jsonschema` dependency.

### Phase 4: Testing

- [x] Task 6: `test_schema_extractor.py`:
      - bundled `template.pptx` extract -> `validate_template_schema` passes;
      - **count invariant**: `len(result['slide_layouts']) == 63` and every
        layout has a `components` array;
      - non-placeholder elements captured (layout 2: >=5 components vs 3
        placeholders);
      - master components non-empty (3 shapes);
      - polygon values in [0,1], exactly 4 points;
      - **font-cardinality assertion**: `assert "font" not in image_component`
        and `assert "font" not in shape_component`;
      - group nesting recurses; empty/blank layout does not crash;
      - **synthetic edge cases** (per architecture review M2): build PPTX in-test
        via python-pptx containing (a) a group with nested shapes, (b) a table
        graphicFrame, (c) a master with zero shapes — assert no crash +
        components captured;
      - **negative test**: non-PPTX / truncated-zip input raises a domain error
        (`TemplateExtractionError`), not a raw `lxml`/`zipfile` traceback.

### Phase 5: CLI & Docs

- [x] Task 7 (optional): CLI entry `python schema_extractor.py --input X.pptx
      --output schema.json` for standalone invocation (bridges to US-5.1).

## Out of Scope (deferred to sibling stories)

- **US-1.2**: cross-product winding validation, non-rectangular polygon vertices.
- **US-1.3**: `type_confidence: "low"` fallback for shapes that do not map
  cleanly via MSO_SHAPE_TYPE (US-1.1 Task 2 covers the deterministic happy-path
  mapping).
- **US-1.4**: font/runs detection (US-1.1 text components carry an empty
  `font: {}`; non-text components omit `font`).
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
  (`_build_slide_size`) — reusable as reference (duplicated per Decision #4).

## Risks

- python-pptx may not cleanly expose `grpSp`/`graphicFrame` subtypes -> may need
  lxml-level parsing for disambiguation.
- Bundled template is a Google Slides export — element naming is non-standard
  ("Google Shape;13;p2"); `name` preserved as-is, mapping must not depend on it.
- SmartArt (`IGX_GRAPHIC`) degrades to `shape`.
- **Polygon winding (per architecture review n1)**: US-1.2 requires
  "anti-clockwise" but lists TL->TR->BR->BL, which is clockwise in both common
  coordinate conventions. US-1.1 emits TL->TR->BR->BL as written; US-1.2 must
  reconcile the label vs the order (possibly reversing output or redefining
  "anti-clockwise" as screen-space CCW).

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` (Epic 1, US-1.1 + Reference JSON Schema).
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` (section 2, US-1.1; section 5 open decisions).
- Architecture review: findings C1/M1/M2/M3/m1/m2/m3/m4/n1/n2 incorporated above.
- Issue: #48.
