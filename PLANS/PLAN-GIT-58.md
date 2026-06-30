# PLAN-GIT-58 — US-4.1: Renderer Reads Embedded JSON (Source-Swap Migration)

**Issue**: #58
**Branch**: GIT-58 (base: dev)
**Priority**: Must Have (P1)
**Status**: Planned (rev 2 — architecture-review C1/M3/M4/M5/M6 + m1-m3 incorporated)

## Goal

Migrate the slide-generation render path to read the embedded `ppt/template_schema.json` (produced by Epic 1/3) instead of the sidecar fingerprint contract (`template_introspector.get_contract`), via a **source-swap adapter**. The renderer keeps `prs.slides.add_slide(layout)` + fingerprint/name matching (low risk, output quality unchanged) and satisfies US-4.1's three ACs. Per the US-4.1 clarification (chenyu's #4: generation "uses the slide master's slide template"), generation uses `add_slide(layout)` and the embedded JSON drives layout selection — **coordinate placement was never a requirement** (see GAP-ANALYSIS §5 Decision 2, clarified). An optional polygon-fidelity consistency check may run, but is not an acceptance gate.

This closes the "write-only" gap: after US-4.1 the embedded JSON is consumed by generation (no longer write-only), and the sidecar degrades to a backward-compat fallback (GAP-ANALYSIS architecture review MAJOR-3 / §5 Decision 1 Coexist; Decision 2 clarified — coordinate placement was never required, so "source swap, polygons as faithful metadata" is the faithful path, not a compromise).

## Strategic Context

Epic 1 (`extract_schema` / `embed_schema`) and Epic 3 (`generate-template-skill`) produce a self-describing templated PPTX with the schema embedded at `ppt/template_schema.json`. But the renderer (`ppt_builder.py:859`) still reads the sidecar `template.pptx.contract.json` via `template_introspector.get_contract` — the embedded JSON has **no production consumer** (referenced only in tests). US-4.1 is the story that makes the renderer consume it.

Research (render-path + data-model map) confirmed: the embedded JSON carries **all raw inputs** the renderer needs; only two *derived* rollups (`fingerprint`, `content_area_in2`) and a few renamed fields are absent. A faithful adapter bridges the gap with **no renderer rewrite forced**. Coordinate placement (abandoning `add_slide`) is **not a requirement** — chenyu's #4 specifies "using the slide master's slide template"; it would also be high-risk and lossy (breaks placeholder inheritance, bullets/theme defaults, the image fast-path, and the entire `template-modifier-skill` clone contract), so it is out of scope (see US-4.1 clarification + GAP-ANALYSIS §5 Decision 2).

## Architecture Decisions (locked)

1. **Source-swap adapter (a bridge, not the end-state intake — arch-review M6)** — new `embedded_schema_to_contract(schema) -> dict` (in `ppt-template-filler/scripts/contract_adapter.py`) produces a dict structurally identical to the sidecar contract. `_resolve_layout_by_fingerprint`, `servable_slide_types`, and `constraint_checker` consume it unchanged (they already take a `contract`). It intentionally *reduces* the embedded schema to the sidecar shape (dropping polygon/font/runs); the v2 intake that exposes those rich fields is **US-4.6** (coordinate placement consumes them).
2. **Prefer embedded, fallback sidecar + provenance (arch-review M4)** — new `get_render_contract(template)` in `ppt_builder.py`: try `read_embedded_schema` → adapter; on `None` (absent) → silent sidecar fallback; on parse error/`TemplateExtractionError` (corrupt) → sidecar fallback **with a warning** naming the template + error (absent ≠ corrupt). The returned contract is tagged `_source ∈ {"embedded","sidecar"}` and logged. Replaces the single call site at `ppt_builder.py:859`. Backward compatible.
3. **Adapter derivation rules (match the sidecar exactly — arch-review M3)** — `fingerprint`: filter `components` to `type=="placeholder"`, **top-level only (do NOT recurse into groups — mirrors sidecar's `layout.placeholders` view)**, drop chrome (`placeholder_type ∈ {date, slide_number, footer, header}`), apply the explicit canonical map `{body→OBJECT, title→TITLE, subtitle→SUBTITLE, picture→PICTURE, chart→OBJECT, table→TABLE, media→MEDIA}` and uppercase. **`chart` (ORG_CHART) → `OBJECT`** (Decision M3-A: match sidecar `_TYPE_CANONICAL`, so it contributes to both fingerprint and `content_area`), NOT the embedded `chart` enum. `content_area_in2`: sum area over placeholders whose canonical type ∈ `{OBJECT}`; compute each by denormalizing the polygon against `slide_dimensions` (the embedded JSON carries no raw per-component EMU). **Note:** clamp+round in `normalize_polygon` (US-1.2) can introduce sub-1% drift vs the sidecar's raw-EMU area, so the parity test uses a tolerance, not strict byte-equality. Renames: `layout_name`/`layout_index` ← embedded; `slide_size` ← `template_metadata.slide_dimensions`.
4. **Generate via layouts (chenyu's intent); polygon fidelity is an optional consistency check, not placement** — re-confirmed against chenyu's #4 ("using the slide master's slide template"), generation uses `add_slide(layout)` (inheriting exact positioning/styling); the embedded JSON drives layout selection. The polygon model (US-1.2) is a faithful description, NOT a placement source. A denormalization consistency check (denormalized polygon vs live placeholder geometry, ≤1%) may run post-`add_slide` as a non-fatal conformance signal, but is **not** an acceptance gate. (US-4.1's original AC3 "creates OOXML at exact positions" was an over-elaboration — corrected; see the chenyu-user-stories US-4.1 historical note + GAP-ANALYSIS §5 Decision 2.)
5. **Templated default template + staleness guard (arch-review M5)** — `embed_schema` the bundled `templates/template.pptx` and commit the templated version, so the default render path uses embedded JSON. Because the embedded JSON lacks the sidecar's mtime-invalidation, `get_render_contract` warns if the embedded schema's layout-name/count signature diverges from the live template (catches edit-without-re-embed staleness).
6. **Single source of truth incl. the clone path (arch-review C1)** — all `get_contract` consumers switch to `get_render_contract`. **Clone path:** `clone_for_over_limit` must, after `prs.save`, **re-run `extract_schema` + `embed_schema`** on the derived `template_new.pptx` so it carries an embedded JSON describing the *cloned* (resized) layout — otherwise python-pptx's save strips the unmodeled `ppt/template_schema.json` part and the extend workflow silently falls back to the sidecar (two-track). This closes two-track and unblocks sidecar deprecation.
7. **`read_embedded_schema` provenance** — it reads from the *same* PPTX being rendered (embed put it there), so schema/template mismatch is not a normal-flow risk. The adapter still guards on `slide_dimensions` sanity.
8. **Consumer call-site coverage (arch-review m1/m2/m3)** — the `get_contract` call sites to migrate are: `ppt_builder.py:859`, `pptx-subagent.md` Stage 0 (incl. the `from template_introspector import get_contract` import line), `ppt-template-filler/SKILL.md` runnable snippets, and **both** `state_machine.py:107` and `:171`. `constraint_checker.py` is **excluded** — it takes `contract` as a parameter (already source-agnostic; no `get_contract` call). A grep audit in Phase 3 confirms no site is missed.

## Deliverables

**New** `.opencode/skills/ppt-template-filler/scripts/contract_adapter.py`
- `embedded_schema_to_contract(schema) -> dict` — sidecar-shaped contract (`source_file`, `slide_size`, `theme` best-effort, `layouts[]` with `name`/`index`/`fingerprint`/`content_area_in2`/`placeholders`).
- `_CANONICAL_MAP` constant (incl. `chart→OBJECT` per Decision 3); `_derive_fingerprint(components)` (top-level placeholders only, chrome dropped, canonical-mapped, uppercased); `_derive_content_area(components, dims)` (sum over `OBJECT`-canonical placeholders, denormalized-polygon area); `denormalize_polygon(polygon, dims) -> tuple[int,int,int,int]` (EMU; shared with the optional AC3 check).

**Change** `.opencode/skills/ppt-template-filler/scripts/ppt_builder.py`
- New `get_render_contract(template_path) -> contract` (prefer embedded → adapter; `None`→silent sidecar; parse error→sidecar **+ warning**; tag `_source`; log). Replaces `get_contract(str(template))` at line 859.
- New `_assert_polygon_fidelity(slide, layout_components, dims)` (optional post-`add_slide` consistency check; warn on >1% drift; non-fatal).

**Change** downstream consumers → `get_render_contract` (arch-review m1/m2/m3)
- `.opencode/agents/pptx-subagent.md` — Stage 0 `servable_slide_types(...)` call site **and** the `from template_introspector import get_contract` import line.
- `.opencode/skills/ppt-template-filler/SKILL.md` — runnable `servable_slide_types(get_contract(...))` snippets.
- `.opencode/skills/template-modifier-skill/scripts/state_machine.py` — **both** call sites (line 107 `plan_resolution`, line 171 `resolve_and_clone`).
- *(Excluded: `constraint_checker.py` — takes `contract` as a param, no `get_contract` call.)*

**Change** clone path (arch-review C1) — `.opencode/skills/template-modifier-skill/scripts/layout_creator.py` (or `state_machine.resolve_and_clone`): after `prs.save(template_new.pptx)`, re-run `extract_schema` + `embed_schema` on the derived file so it carries the cloned layout's embedded JSON.

**Templated default template + staleness guard (arch-review M5)** — `embed_schema` the bundled `templates/template.pptx`; commit the templated version. Add a divergence check in `get_render_contract` (embedded layout-name/count signature vs live template → warn on drift).

**Tests** (`.opencode/skills/ppt-template-filler/scripts/tests/`)
- Adapter **parity** (broadened — arch-review M3): adapter output vs sidecar per-layout on the bundled template **plus synthetic fixtures** containing an ORG_CHART placeholder, a grouped placeholder, and an off-slide-overflowing placeholder. `fingerprint` exact; `content_area_in2` within tolerance (sub-1%, due to polygon clamp/round).
- **`_source` provenance + malformed handling (M4)**: templated PPTX → `_source=="embedded"`; non-templated → `_source=="sidecar"` (silent); corrupted `ppt/template_schema.json` → `_source=="sidecar"` **+ warning**.
- **Clone re-embed (C1)**: after `resolve_and_clone`, `read_embedded_schema(template_new.pptx)` is non-`None` and lists the extended layout.
- **AC3 (uses-layouts)**: generation uses `add_slide(layout)`; embedded JSON drives selection.
- **End-to-end**: render from a templated PPTX (embedded JSON consumed; sidecar not triggered).
- **Fallback**: non-templated PPTX → sidecar path still renders.
- **Regression**: existing 112 schema_extractor + render tests stay green.

## Acceptance Criteria (US-4.1) — to deliver

- [ ] AC1 — Engine reads JSON from the zip (`read_embedded_schema`); does not re-extract or re-parse XML.
- [ ] AC2 — Layout selection based on `layout_name` matching (fingerprint/name matching preserved via the adapter).
- [ ] AC3 — Generated slides use the template's own layouts (`add_slide`); the embedded JSON drives layout selection, not element placement at polygon coordinates. (A polygon-fidelity consistency check is optional and non-fatal.)

## Implementation Phases

### Phase 1: Adapter + parity tests (contract_adapter.py)
- [ ] Task 1: `embedded_schema_to_contract(schema)` + `_CANONICAL_MAP` (incl. `chart→OBJECT`) + `_derive_fingerprint` (top-level only, chrome-dropped) + `_derive_content_area` + `denormalize_polygon`.
- [ ] Task 2: Parity test — broadened (arch-review M3): adapter vs sidecar on the bundled template **plus synthetic fixtures** (ORG_CHART placeholder, grouped placeholder, off-slide-overflowing placeholder); `fingerprint` exact, `content_area_in2` within sub-1% tolerance.

### Phase 2: Wire ppt_builder + provenance + AC3 check
- [ ] Task 3: `get_render_contract(template)` — prefer embedded → adapter; `None`→silent sidecar; parse error→sidecar **+ warning**; tag `_source`; log.
- [ ] Task 4: Replace `ppt_builder.py:859` call site; add end-to-end + fallback + `_source`/malformed tests.
- [ ] Task 5: AC3 assertion — verify generation uses `add_slide(layout)` and the embedded JSON drives selection (not coordinate placement). Optional `_assert_polygon_fidelity` consistency check (post-`add_slide`, warn on drift); non-fatal, not an AC gate.

### Phase 3: Migrate downstream consumers + clone re-embed (arch-review C1/m1/m2/m3)
- [ ] Task 6: `pptx-subagent.md` Stage 0 call site **and** import line → `get_render_contract`; `ppt-template-filler/SKILL.md` runnable snippets updated.
- [ ] Task 7a: `template-modifier-skill/scripts/state_machine.py` — **both** call sites (107, 171) → `get_render_contract`. (`constraint_checker.py` excluded — source-agnostic.)
- [ ] Task 7b: **Clone re-embed (C1)** — in `layout_creator`/`state_machine.resolve_and_clone`, after `prs.save(template_new.pptx)` re-run `extract_schema` + `embed_schema`; test `read_embedded_schema(template_new.pptx)` non-`None` + lists extended layout.
- [ ] Task 7c: grep audit — confirm no `get_contract(` call site remains unmigrated.

### Phase 4: Templated default template + staleness guard + docs
- [ ] Task 8: `embed_schema` the bundled `template.pptx`; commit the templated version; add the staleness-guard divergence check in `get_render_contract` (arch-review M5).
- [ ] Task 9: `chenyu-user-stories.md` US-4.1 ACs → `[x]`; `GAP-ANALYSIS.md` Revision 9 (US-4.1 ✅); `AGENTS.md` / `README.md` migration note.

## Test matrix

| Case | Expected |
| --- | --- |
| adapter parity (bundled + synthetic fixtures: ORG_CHART, grouped, off-slide) | per-layout `fingerprint` exact == sidecar; `content_area_in2` within sub-1% tolerance |
| `_source` provenance + malformed (M4) | templated→`embedded`; non-templated→`sidecar` (silent); corrupted→`sidecar` + warning |
| clone re-embed (C1) | after `resolve_and_clone`, `read_embedded_schema(template_new.pptx)` non-`None` + lists extended layout |
| AC3 (uses-layouts) | generation uses `add_slide(layout)`; embedded JSON drives selection, not coordinate placement |
| polygon-fidelity consistency check (optional) | denormalized polygons ≤1% from live placeholder geometry (non-fatal, warn-only) |
| render from templated PPTX | consumes embedded JSON; sidecar not triggered |
| non-templated PPTX | falls back to sidecar; renders normally |
| existing schema_extractor + render suites | green (no regression) |

## Verification

```powershell
python -m pytest .opencode/skills/ppt-template-filler/scripts/tests/ -q
# End-to-end: render from a templated PPTX via the engine
python -c "import sys; sys.path.insert(0,'.opencode/skills/ppt-template-filler/scripts'); from ppt_builder import generate_ppt_from_data, get_render_contract; c=get_render_contract('.opencode/skills/ppt-template-filler/scripts/templates/template.pptx'); print('contract source:', 'embedded' if c.get('_source')=='embedded' else 'sidecar')"
```

## Out of Scope / Open Questions

- **Coordinate placement** — not a requirement (chenyu's #4 specifies "using the slide master's slide template" / `add_slide`); out of scope. The earlier "Option A — rejected" framing assumed it was a live alternative, which the US-4.1 clarification corrects (see GAP-ANALYSIS §5 Decision 2). It would also be high-risk and lossy (breaks placeholder inheritance + the template-modifier-skill clone chain).
- **Full sidecar removal** — this issue degrades it to a fallback; the clone-path re-embed (Decision 6 / C1) removes the last two-track holdout, so a separate issue can fully delete `template_introspector` once all templates are templated.
- **Chart-theme semantic-role → clrScheme fine mapping** — adapter does best-effort; fidelity verified in Phase 1 parity tests.
- **Non-rectangular polygon vertices** — a known polygon-model limitation (rectangular bounding boxes only); this affects metadata fidelity, not placement (placement is layout-inherited via `add_slide`).

## Risks

- **Adapter derivation fidelity (arch-review M3)** — `fingerprint` must match the sidecar exactly (incl. `ORG_CHART→OBJECT`, top-level-only grouping, chrome drop); mitigated by the broadened Phase 1 parity test. `content_area_in2` carries sub-1% drift (polygon clamp/round) → parity test uses a tolerance.
- **Clone path strips embedded JSON (arch-review C1)** — mitigated by Task 7b (re-embed after `prs.save`) + the clone-re-embed test; without it the extend workflow silently two-tracks to the sidecar.
- **Silent fallback masks drift (arch-review M4)** — mitigated by the `_source` tag + log + the absent/corrupt distinction (corrupt → warning).
- **Templated default template staleness (arch-review M5)** — embedded JSON lacks mtime-invalidation; mitigated by the divergence-check warning in `get_render_contract`.
- **Theme shape divergence** — sidecar `theme.colors` (clrScheme roles) vs embedded semantic roles (`primary_color`/…); adapter maps best-effort, chart-color fidelity verified in Phase 1.
- **Consumer call-site coverage (arch-review m1/m2/m3)** — any missed `get_contract` call leaves a two-track read; mitigated by Task 7c grep audit. `constraint_checker` correctly excluded (source-agnostic).
- **Templated default template bloat** — embedding adds ~tens of KB to the committed `template.pptx`; acceptable, size logged.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 4, US-4.1.
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 US-4.1 (🟡 Partial), §5 Decision 1/2.
- Predecessors: US-1.5 (issue #55), US-3.1 (issue #56).
- PLAN format template: `PLANS/PLAN-GIT-56.md`.
- Render-path / data-model gap map: cited inline (`ppt_builder.py:859`, `template_introspector.py:182-207`, `schema_extractor.py:724-833`).
