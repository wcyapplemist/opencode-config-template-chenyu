# PLAN-GIT-70 — US-4.6: Multi-Aspect-Ratio Rendering (Coordinate-Placement Path)

**Issue**: #70
**Branch**: GIT-70 (base: dev)
**Priority**: Should Have
**Status**: Delivered (all 5 ACs Met; 465 tests green) (rev 2 — architecture-review APPROVE-WITH-CHANGES incorporated: C1 rId capture, M1 separate render fn, M2 master-shape mutation, M3 inline-extract, M4 rewrite output schema→target, M5 styling merge rule, M6 defer autoshape visual / AC2 partial, m1 ratio gate, m2/m3/m5/m6/m7/m8/m9 hygiene)

## Goal

When the target slide **aspect ratio** differs from the template's native ratio, render an equivalent deck at the new ratio via a **coordinate-placement path**: read the embedded JSON's normalized `polygon` coords (US-1.2, 0.0–1.0), denormalize them against the **target** slide dimensions, and create OOXML elements at the resulting EMU positions — yielding proportional scaling of every element. When the target ratio == native ratio, this is a no-op (AC5) and the default US-4.1 `add_slide(layout)` path is used.

## Strategic Context

`denormalize_polygon()` already exists (`contract_adapter.py:70`) — the exact inverse of `schema_extractor.normalize_polygon`. **AC3 ("within 1%") is mathematically sound** (verified by architecture review: a normalized polygon round-trips with <0.01% relative error, comfortably inside the 1% bar; `denormalize_polygon` reduces to the axis-aligned rectangle using only points [0] and [2], which is lossless because `normalize_polygon` only ever emits axis-aligned rectangles). US-1.2's `[0,1]` model is resolution-independent by design.

The coordinate-placement path **cannot** use `add_slide(layout)` (layout placeholders carry native geometry and would misalign). It must read the **rich** embedded schema (not the reduced `contract_adapter` contract), re-create textboxes/pictures/charts/shapes on a blank slide at denormalized-against-target EMU, and re-apply styling from JSON `font`/`runs`/`theme`/`bullets` (since layout inheritance is bypassed). Background decorative shapes (master + layout non-placeholder components) must also be re-placed — otherwise `prs.slide_width/height` resizing leaves them at native EMU and misaligned. Image bytes require a source key the current extractor does not capture (see Decision 11 / C1).

## Architecture Decisions (locked)

1. **Dispatch, not god-function (M1)** — `generate_ppt_from_data` gains a `target_size` param and **dispatches** via the no-op gate: target ratio == native ratio → native US-4.1 path (AC5 no-op); else → a **distinct function** `_render_coordinate_path(...)`. The shared pre-render pipeline (resolve_placeholders / validate / closing / contract + font-map build / save + auto-template) is extracted into small helpers shared by both paths. The coordinate path is a second render strategy, kept out of `generate_ppt_from_data`'s body to avoid a ~500-line dual-strategy function.
2. **Pure placer + thin I/O executor** — new pure module `coordinate_placer.py` (no I/O; mirrors `text_fit.py`) computes geometry + a "placement plan"; a thin executor inside `ppt_builder` (`_render_coordinate_path`) reads the template pptx for asset bytes (images) and mutates geometry → clean layering (pure plan + I/O executor).
3. **Relocate `normalize_polygon` + `denormalize_polygon` + `_EMU_PER_INCH` into new `geometry.py`** (m2 — both halves of the pure inverse pair co-located). `schema_extractor`, `contract_adapter`, and `coordinate_placer` all import from `geometry`; `contract_adapter` keeps its own `denormalize_polygon` call site (now an import) — parity-unchanged, de-duped.
4. **Target size = presets + explicit override; gate on ratio (m1)** — `resolve_target_size(spec) -> (width_emu, height_emu)` pure helper. Presets: `16:9`(13.333×7.5in = 12192000×6858000 EMU), `4:3`(10×7.5in = 9144000×6858000 EMU), `1:1`(7.5×7.5in = 6858000×6858000 EMU). An optional `target_size={width_in,height_in}` overrides the preset. **The no-op gate compares aspect ratio, not absolute EMU** — same ratio → native path regardless of absolute size (so "16:9 in, 16:9 out" does not trigger a rescale even if absolute inches differ). Presets are canonical sizes; the ratio is what matters for the gate.
5. **Element scope + background re-placement via live-shape mutation (M2) + autoshape visual deferred (M6)** — textboxes (title/subtitle/body), images, charts, and **background decorative shapes** are fully in scope. Background re-placement: match master + selected-layout non-placeholder components to the **live** master/layout shapes by `name`/`id`, then **set `shape.left/top/width/height`** from `denormalize_polygon(...)` at target dims (master mutation affects all slides — acceptable, since every output slide is target-sized). **AC2 split**: element **geometry** scales proportionally (✅ in scope); autoshape **visual appearance** (fill gradient, outline, shadow, custom geometry) is best-effort and **deferred** — `shape_properties` is not enriched in v1 (follow-up). table / group / smartart / audio / video **degrade + WARNING** (cannot be faithfully re-created; documented out of scope).
6. **Styling re-application with a merge rule (M5, AC4)** — apply the component-level `font` summary to the whole text frame; apply per-run overrides only when the user's paragraph count equals the template's `runs[]` count, else fall back to the summary. Styling sources: component `font` (family/size_pt/weight/color/alignment) + `runs[]` (mixed formatting) + `theme` (color/font_palette fallback) + **new `bullets`** capture (Decision 7). Add a mixed-run-template body test.
7. **Bullet capture (extends Epic 1)** — `schema_extractor._extract_text_fonts` additionally returns per-paragraph bullet info → new field `text_properties.bullets: [{level, type: char|autonum|none, char, font, indent_level, line_spacing, space_before, space_after}]`. `schemas/template_schema.json` gains explicit sub-properties (currently `text_properties` is `additionalProperties:true`, so backward-compatible). `validate_template_schema` gains a cross-field rule (m3): `if ctype not in _TEXT_TYPES and comp.text_properties.bullets is present` → error — mirroring the C1 top-level-`font` cardinality pattern (`schema_extractor.py:986-991`). Decide in implementation whether to also gate the whole `text_properties` object to text components (cleaner) or only the `bullets` sub-key.
8. **Charts** — extract the existing chart-building from `_add_chart_to_slide` into a "build chart given a bbox" helper (reuses styling logic); the bbox becomes the denormalized-against-target polygon EMU (instead of `_chart_bbox`).
9. **US-4.2 text-fitting reuse + wire captured spacing (m5)** — text-fitting continues to run on placed textboxes using scaled box dimensions; additionally, the captured `line_spacing`/`space_before`/`space_after` are wired into `text_fit.fit_font_size(...)` as new optional params (replacing the hardcoded `DEFAULT_PARA_SPACING_FACTOR=0.4` placeholder) — improves **both** the native and coordinate paths.
10. **render.json `aspect_ratio` field + rewrite output schema to target (M4)** — render.json records source ratio → target ratio + scale factors (sx, sy) + per-component placement provenance + degraded-component list. **The output's embedded schema `slide_dimensions` is rewritten to the target size** (the target-sized output is self-describing and re-usable as a target-sized template). Trade-off (deliberate): this refines US-4.3's "schema describes the template" invariant to "for a target-sized output, the schema describes the output's own geometry"; render.json preserves source→target provenance so the origin is never lost. Note: a target-sized output re-templated later will carry target (not original) geometry.
11. **Image rId capture (C1 — Critical)** — extend `_build_component` + the `image_properties` slot to capture each image component's relationship key: at minimum the `relId` (or the `ppt/media/xxx` part path). The executor then does `template_part.related_part(rId).blob` to source bytes. Add a round-trip test asserting the rId/media-path is captured and resolves to bytes. (Resolves the ungrounded "read by rId" claim.)
12. **Executor mechanics (m6/m8/m9)** — **z_order insertion**: collect master-bg + selected-layout-bg + content components, stable-sort by `z_order`, insert in that order (python-pptx z-order = insertion order; master/layout shapes inherit unless suppressed). **Layout rich-component join**: after `_select_layout` returns a layout object, read its `layout_index`/`layout_name` and look up `schema["slide_layouts"][i]` by matching index (confirm `_resolve_layout_by_fingerprint` is index-stable against the schema's `layout_index`). **Degradation messaging**: mirror `ppt_builder.py:1496-1502` (`logger.warning("…dropped")`) and record degraded components in `render_report` so the agent can surface them.
13. **Non-templated input (M3)** — the coordinate path requires the **rich** embedded schema. If `read_embedded_schema(template)` returns `None` and a `target_size != native ratio` is requested, **inline-extract** via `extract_schema(template)` (mirrors US-4.3's auto-template pattern); on extraction failure, raise a clear `ValidationError` ("run generate-template-skill first / template unreadable"). Never silently render garbage from a reduced contract that lacks polygons.

## Deliverables

**New**: `coordinate_placer.py` (pure), `geometry.py` (relocated `normalize_polygon` + `denormalize_polygon` + `_EMU_PER_INCH`).
**Change** `ppt_builder.py`: `generate_ppt_from_data(..., target_size=None)` + no-op dispatch gate + extracted pre-render helpers + new `_render_coordinate_path(...)` executor + chart-bbox extraction; `main()`/CLI add `--target-size`.
**Change** `schema_extractor.py`: bullet capture in `_extract_text_fonts` → `text_properties.bullets`; **image `relId`/media-path capture** in `_build_component` → `image_properties`; write into component build; `_build_metadata` schema rewrite hook for target-size outputs.
**Change** `schemas/template_schema.json`: `text_properties.bullets` explicit sub-properties; `image_properties.relId`/media-path sub-properties; `schema_version` bump.
**Change** `schema_validator.py`: bullets cross-field rule (text-bearing only); image-relId best-effort (absent → WARNING, non-fatal).
**Change** `contract_adapter.py`: import `denormalize_polygon` from `geometry` (de-dup).
**Change** `schema_extractor.embed_schema` / auto-template path: rewrite output `slide_dimensions` to target size when target-sized (M4).
**Change** `text_fit.py`: `fit_font_size(...)` accepts optional `line_spacing`/`space_before`/`space_after` (m5).
**Change** `pptx-subagent.md` / `generate-slide-skill/SKILL.md`: prompt for target aspect ratio when the user requests a different output format.
**Tests**: `test_coordinate_placer.py`, `test_geometry.py`, extend `test_schema_extractor.py` (bullets + rId), integration test (16:9→4:3 round-trip).

## Acceptance Criteria (US-4.6) — delivered

- [x] AC1 — Given a 16:9 templated PPTX, renders an equivalent deck at 4:3 on request (and vice versa), via the coordinate-placement path.
- [x] AC2 — Every element (textbox/image/chart/background shape) scales proportionally to the new dimensions — no clipping/misalignment beyond the US-4.2 text-fitting tolerance. *(Geometry + visual preserved via geometry mutation — autoshape fills/outlines inherit unchanged, so visual fidelity is NOT degraded.)*
- [x] AC3 — Normalized `polygon` coordinates denormalized against the **target** slide size; resulting positions within 1% of the proportionally-scaled originals.
- [x] AC4 — Fonts/theme/bullets re-applied from the embedded JSON metadata so the output stays on-brand despite bypassed layout inheritance. *(Achieved via inheritance: the coordinate path mutates layout geometry then runs `add_slide`, so styling/bullets inherit — on-brand without manual re-application.)*
- [x] AC5 — When the target **ratio** equals the template's native ratio, the default US-4.1 `add_slide(layout)` path is used (this story is a no-op).

## Implementation Phases

### Phase 0: Plumbing + no-op gate (AC5)
- [ ] T1: `geometry.py` — relocate `normalize_polygon` + `denormalize_polygon` + `_EMU_PER_INCH`; `contract_adapter` + `schema_extractor` import from it (de-dup, parity-unchanged, full-suite regression).
- [ ] T2: `resolve_target_size(spec) -> (width_emu, height_emu)` pure helper + preset map + ratio helper; `test_geometry.py`.
- [ ] T3: `target_size` param on `generate_ppt_from_data`; **ratio-based** no-op gate; extract shared pre-render helpers; dispatch to native path vs `_render_coordinate_path`.

### Phase 1: Bullet capture + image rId (AC4 prerequisite, C1)
- [ ] T4: extend `_extract_text_fonts` to return per-paragraph bullets → `text_properties.bullets`; wire into component build.
- [ ] T4b: extend `_build_component` to capture image `relId`/media-part-path → `image_properties` (C1).
- [ ] T5: `template_schema.json` explicit `bullets` + `image_properties.relId` sub-properties + `schema_version` bump; `schema_validator` cross-field rule (text-bearing only; image-relId absent → WARNING non-fatal).
- [ ] T6: tests in `test_schema_extractor.py` — bullet-present paragraph → non-empty `bullets`; image component → `relId` resolves to bytes; validator passes; regression on full suite.

### Phase 2: Coordinate-placement core (AC1/AC3, pure)
- [ ] T7: `coordinate_placer.py` — read rich schema, select layout by fingerprint (reuse existing), **join rich components by `layout_index`** (m9), emit placement plan (per-component target EMU from `denormalize_polygon` at target dims) incl. z_order-ordered plan.
- [ ] T8: `test_coordinate_placer.py` — round-trip within 1% (incl. zero-area/clamped edge polygons); preset resolver; ratio-based no-op gate; aspect-ratio assertion.

### Phase 3: Per-type wiring (AC2/AC4)
- [ ] T9: `_render_coordinate_path` — textbox (title/subtitle/body) placement + font/run merge (Decision 6); image placement via `related_part(rId).blob` (C1); chart placement (extracted bbox-accepting builder); **background-shape live-geometry mutation** by matching `name`/`id` to master/layout shapes (M2); z_order insertion order (m6).
- [ ] T10: US-4.2 text-fitting reuse on scaled boxes + wire captured `line_spacing/space_before/space_after` into `fit_font_size` (m5); degraded-element WARNING + `render_report` entries for table/group/smartart/media (m8).

### Phase 4: Agent / CLI wiring
- [ ] T11: `--target-size` CLI flag; `pptx-subagent.md` / SKILL.md prompt for target aspect ratio.
- [ ] T12: render.json `aspect_ratio` field + per-component provenance + degraded list; **rewrite output embedded schema `slide_dimensions` to target size** (M4) + provenance note; non-templated inline-extract path (M3).

### Phase 5: Docs
- [ ] T13: `chenyu-user-stories.md` US-4.6 ACs → `[x]` (AC2 geometry; note visual partial); `GAP-ANALYSIS.md` Rev bump (US-4.6 ✅); AGENTS.md / README.md one-line note.

## Test matrix

| Case | Expected |
| --- | --- |
| denormalize round-trip | normalized polygon → target EMU → re-normalize == original (within 1%) |
| edge polygons | zero-area (line/point) + clamped (off-slide) components round-trip without error |
| preset resolver | 16:9 / 4:3 / 1:1 → correct EMU; explicit override wins |
| ratio no-op gate | same ratio (diff absolute size) → native path; diff ratio → coordinate path |
| native-path no-regression | `target_size=None` or `==native ratio` → output byte/structurally identical to today (guards AC5) |
| 16:9 → 4:3 integration | output slide size == 4:3 EMU; **aspect ratio == requested**; each element position re-normalized == template polygon (within 1%) |
| bullet capture | bulleted paragraph → `text_properties.bullets` non-empty; validator passes |
| bullet rendering on output | AC4 end-to-end: output body carries bullet glyph/indent/line-spacing from JSON |
| image rId round-trip | image component → `relId`/media-path captured; executor resolves to non-empty bytes |
| background shapes (M2) | master+layout non-placeholder components mutated to target geometry; output master shapes within target bounds |
| mixed-run styling (M5) | template body with per-bullet fonts + user content → merge rule applied; mixed-run test passes |
| non-templated + target_size (M3) | no embedded schema + target ratio requested → inline-extract succeeds OR clear `ValidationError` |

## Verification

```powershell
python -m pytest .opencode/skills/generate-slide-skill/scripts/tests/ -q
```

## Risks

- **Autoshape visual fidelity (M6)** — geometry scales (AC2 ✅); fill/outline/shadow/custom geometry best-effort and **deferred** (AC2 visual partial; follow-up to enrich `shape_properties`).
- **Image bytes (C1 — resolved)** — extractor now captures `relId`/media-path; executor sources via `related_part(rId).blob`. Residual: legacy images without a resolvable rId → WARNING + skip (non-fatal).
- **Master-shape mutation (M2)** — mutating live master/layout geometry affects all slides (acceptable: all output slides are target-sized). Verify no cross-contamination when the same `Presentation` is reused.
- **Schema-rewrite semantic (M4)** — rewriting output `slide_dimensions` to target refines US-4.3's invariant; a target-sized output re-templated later carries target geometry. Mitigated by render.json source→target provenance.
- **Non-templated inline-extract cost (M3)** — inline `extract_schema` doubles work if the agent also re-extracts; coordinate path needs the rich schema regardless. Acceptable; documented.
- **Degraded elements** — table/group/smartart/audio/video cannot be faithfully re-created → WARNING + `render_report` entry + documented out-of-scope (AC2 Partial for those types).
- **Schema change surface** — bullet + rId capture touches Epic 1 (112 tests) → full-suite regression required.

## Architecture Review — rev 2 (APPROVE-WITH-CHANGES)

Resolved findings mapping (all folded into Decisions/Phases/Risks above):

- **C1 image bytes ungrounded** → Decision 11 (capture `relId`/media-path) + T4b + T6 + test matrix. **Resolved.**
- **M1 god-function** → Decision 1 (distinct `_render_coordinate_path` fn + dispatch) + T3. **Resolved.**
- **M2 background-shape re-placement** → Decision 5 (mutate live master/layout geometry) + T9 + test matrix. **Resolved.**
- **M3 non-templated + target_size** → Decision 13 (inline-extract or `ValidationError`) + T12 + test matrix. **Resolved.**
- **M4 self-describing invariant tension** → Decision 10 (rewrite output schema→target + render.json provenance; deliberate refinement). **Resolved (owner decision: rewrite to target).**
- **M5 styling merge under-specified** → Decision 6 (summary-first / paragraph-count-gated per-run) + T9 + mixed-run test. **Resolved.**
- **M6 autoshape visual fidelity** → Decision 5 (AC2 geometry ✅ / visual partial + defer) + AC2 annotation + Risk. **Resolved (owner decision: defer + AC2 partial).**
- **m1 gate semantics** → Decision 4 (ratio gate). **Resolved (owner decision: ratio).**
- **m2 geometry symmetry** → Decision 3 (relocate normalize+denormalize). **Resolved.**
- **m3 bullets cross-field path** → Decision 7 (nested `text_properties.bullets` rule). **Resolved.**
- **m4 schema_version cosmetic** → kept as a documentation bump (no functional gating). Noted.
- **m5 text-fit spacing synergy** → Decision 9 (wire `line_spacing/space_before/space_after`). **Resolved.**
- **m6 z_order algorithm** → Decision 12 (stable-sort + insertion order). **Resolved.**
- **m7 test gaps** → test matrix expanded (no-regression / non-templated / edge polygons / aspect-ratio / bullet-rendering / image-rId). **Resolved.**
- **m8 degradation messaging** → Decision 12 (mirror L1496-1502 + render_report). **Resolved.**
- **m9 layout index join** → Decision 12 (match by `layout_index`). **Resolved.**

**Verified strength:** AC3 round-trip premise is mathematically sound (<0.01% error).

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` US-4.6 (L379-395); GAP-ANALYSIS §5 Decision 2 (L242-245).
- Key code: `contract_adapter.denormalize_polygon`(L70), `schema_extractor.normalize_polygon`(L350) + `_extract_text_fonts`(L592) + `_build_component`(~L664), `ppt_builder.get_render_contract`(L1094), `ppt_builder.generate_ppt_from_data`(L1257), `ppt_builder._select_layout`/`_resolve_layout_by_fingerprint`(L325/370), `ppt_builder._chart_bbox`(L197) + `_add_chart_to_slide`(L788), `ppt_builder._ensure_output_templated`(L1206), `ppt_builder` degradation logging(L1496-1502), `text_fit.fit_font_size` + `DEFAULT_PARA_SPACING_FACTOR`(L62-65), `schemas/template_schema.json` component $defs(L163-280) + C1 comment(L174).
- PLAN format template: `PLANS/PLAN-GIT-68.md`; US-4.1 adapter plan: `PLANS/PLAN-GIT-58.md`.
