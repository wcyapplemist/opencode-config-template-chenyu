# PLAN-GIT-58 — US-4.1: Renderer Reads Embedded JSON (Source-Swap Migration)

**Issue**: #58
**Branch**: GIT-58 (base: dev)
**Priority**: Must Have (P1)
**Status**: Planned

## Goal

Migrate the slide-generation render path to read the embedded `ppt/template_schema.json` (produced by Epic 1/3) instead of the sidecar fingerprint contract (`template_introspector.get_contract`), via a **source-swap adapter**. The renderer keeps `prs.slides.add_slide(layout)` + fingerprint/name matching (low risk, output quality unchanged) and satisfies US-4.1's three ACs. **AC3 is satisfied as a conformance assertion** — denormalized polygons are verified ≤1% from the live placeholder geometry (and `add_slide` inherits the exact original geometry at 0% drift, so this is more accurate than manual coordinate placement).

This closes the "write-only" gap: after US-4.1 the embedded JSON is consumed by generation (no longer write-only), and the sidecar degrades to a backward-compat fallback (GAP-ANALYSIS architecture review MAJOR-3 / §5 Decision 1 Coexist → Decision 2 resolved as "source swap, polygons metadata-only").

## Strategic Context

Epic 1 (`extract_schema` / `embed_schema`) and Epic 3 (`generate-template-skill`) produce a self-describing templated PPTX with the schema embedded at `ppt/template_schema.json`. But the renderer (`ppt_builder.py:859`) still reads the sidecar `template.pptx.contract.json` via `template_introspector.get_contract` — the embedded JSON has **no production consumer** (referenced only in tests). US-4.1 is the story that makes the renderer consume it.

Research (render-path + data-model map) confirmed: the embedded JSON carries **all raw inputs** the renderer needs; only two *derived* rollups (`fingerprint`, `content_area_in2`) and a few renamed fields are absent. A faithful adapter bridges the gap with **no renderer rewrite forced**. The literal alternative (coordinate placement, abandoning `add_slide`) was rejected as high-risk and lossy (breaks placeholder inheritance, bullets/theme defaults, the image fast-path, and the entire `template-modifier-skill` clone contract).

## Architecture Decisions (locked)

1. **Source-swap adapter** — new `embedded_schema_to_contract(schema) -> dict` (in `ppt-template-filler/scripts/contract_adapter.py`) produces a dict structurally identical to the sidecar contract. `_resolve_layout_by_fingerprint`, `servable_slide_types`, and `constraint_checker` consume it unchanged (they already take a `contract`).
2. **Prefer embedded, fallback sidecar** — new `get_render_contract(template)` in `ppt_builder.py`: try `read_embedded_schema` → adapter; on absence/`None`/error, fall back to `get_contract` (sidecar). Replaces the single call site at `ppt_builder.py:859`. Backward compatible.
3. **Adapter derivation rules** — `fingerprint`: filter `components` to `type=="placeholder"`, drop chrome (`placeholder_type ∈ {date, slide_number, footer, header}`), map `body`→`OBJECT`, uppercase. `content_area_in2`: sum denormalized polygon area of `placeholder_type=="body"` components (EMU → inches² via `/914400`). Renames: `layout_name`/`layout_index` ← embedded; `slide_size` ← `template_metadata.slide_dimensions`.
4. **AC3 as conformance assertion (not placement)** — after `add_slide`, denormalize each placeholder component's polygon to EMU and assert ≤1% from the live placeholder's `left/top/width/height` (`ppt_builder.py:435-454`). Warn (non-fatal) on drift. Satisfies AC3 more accurately than manual placement (0% drift via inheritance).
5. **Templated default template** — `embed_schema` the bundled `templates/template.pptx` and commit the templated version, so the default render path uses embedded JSON (not the sidecar fallback).
6. **Consumers migrated together (single source of truth)** — agent Stage 0 (`servable_slide_types(get_contract(...))` in `pptx-subagent.md`) and `template-modifier-skill` (`state_machine.py` / `template_reader.py` / `constraint_checker.py` `get_contract` call sites) all switch to `get_render_contract`. No two-track.
7. **`read_embedded_schema` provenance** — it reads from the *same* PPTX being rendered (embed put it there), so schema/template mismatch is not a normal-flow risk. The adapter still guards on `slide_dimensions` sanity.

## Deliverables

**New** `.opencode/skills/ppt-template-filler/scripts/contract_adapter.py`
- `embedded_schema_to_contract(schema) -> dict` — sidecar-shaped contract (`source_file`, `slide_size`, `theme` best-effort, `layouts[]` with `name`/`index`/`fingerprint`/`content_area_in2`/`placeholders`).
- `_derive_fingerprint(components) -> list[str]`, `_derive_content_area(components, dims) -> float`, `denormalize_polygon(polygon, dims) -> tuple[int,int,int,int]` (EMU; shared with the AC3 check).

**Change** `.opencode/skills/ppt-template-filler/scripts/ppt_builder.py`
- New `get_render_contract(template_path) -> contract` (prefer embedded → adapter; fallback `get_contract`).
- Replace `get_contract(str(template))` at line 859 → `get_render_contract`.
- New `_assert_polygon_fidelity(slide, layout_components, dims)` (post-`add_slide` AC3 check; warn on >1% drift).

**Change** downstream consumers → `get_render_contract`
- `.opencode/agents/pptx-subagent.md` Stage 0 `servable_slide_types(...)` call site.
- `.opencode/skills/template-modifier-skill/scripts/` `state_machine.py`, `template_reader.py`, `constraint_checker.py`.

**Templated default template** — `templates/template.pptx` gets `ppt/template_schema.json` embedded and committed.

**Tests** (`.opencode/skills/ppt-template-filler/scripts/tests/`)
- Adapter **parity**: `embedded_schema_to_contract(read_embedded_schema(bundled))` matches `get_contract(bundled)` per-layout (`fingerprint`, `content_area_in2`, `name`).
- **AC3**: denormalized polygons ≤1% from live placeholder geometry.
- **End-to-end**: render from a templated PPTX (embedded JSON consumed; sidecar not triggered).
- **Fallback**: non-templated PPTX → sidecar path still renders.
- **Regression**: existing 112 schema_extractor + render tests stay green.

## Acceptance Criteria (US-4.1) — to deliver

- [ ] AC1 — Engine reads JSON from the zip (`read_embedded_schema`); does not re-extract or re-parse XML.
- [ ] AC2 — Layout selection based on `layout_name` matching (fingerprint/name matching preserved via the adapter).
- [ ] AC3 — Denormalized EMU coordinates within 1% of original positions (conformance assertion; warns on deviation).

## Implementation Phases

### Phase 1: Adapter + parity tests (contract_adapter.py)
- [ ] Task 1: `embedded_schema_to_contract(schema)` + `_derive_fingerprint` + `_derive_content_area` + `denormalize_polygon`.
- [ ] Task 2: Parity test — adapter output == sidecar contract per-layout on the bundled template (fingerprint, content_area, name).

### Phase 2: Wire ppt_builder + AC3 assertion
- [ ] Task 3: `get_render_contract(template)` (prefer embedded → adapter; fallback sidecar).
- [ ] Task 4: Replace `ppt_builder.py:859` call site; add end-to-end + fallback tests.
- [ ] Task 5: `_assert_polygon_fidelity` (AC3 conformance check, post-add_slide, warn on drift) + AC3 test.

### Phase 3: Migrate downstream consumers
- [ ] Task 6: `pptx-subagent.md` Stage 0 call site → `get_render_contract`.
- [ ] Task 7: `template-modifier-skill` `state_machine` / `template_reader` / `constraint_checker` call sites → `get_render_contract`.

### Phase 4: Templated default template + docs
- [ ] Task 8: `embed_schema` the bundled `template.pptx`; commit the templated version.
- [ ] Task 9: `chenyu-user-stories.md` US-4.1 ACs → `[x]`; `GAP-ANALYSIS.md` Revision 9 (US-4.1 ✅); `AGENTS.md` / `README.md` migration note.

## Test matrix

| Case | Expected |
| --- | --- |
| adapter parity (bundled template) | per-layout `fingerprint`/`content_area_in2`/`name` == sidecar |
| AC3 denormalization | denormalized polygons ≤1% from live placeholder geometry |
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

- **Literal coordinate placement (Option A)** — rejected (high risk, lossy, breaks template-modifier-skill clone chain).
- **Full sidecar removal** — this issue degrades it to a fallback; a separate issue removes `template_introspector` once all templates are templated.
- **Chart-theme semantic-role → clrScheme fine mapping** — adapter does best-effort; fidelity verified in Phase 1 parity tests.
- **Non-rectangular polygon vertices** — out of scope (polygons are rectangular bounding boxes); coordinate placement of rotated shapes stays deferred.

## Risks

- **Adapter derivation fidelity** — `fingerprint`/`content_area_in2` must match the sidecar exactly; mitigated by the Phase 1 parity test (byte-equivalent on the bundled template).
- **Theme shape divergence** — sidecar `theme.colors` (clrScheme roles) vs embedded semantic roles (`primary_color`/…); adapter maps best-effort, chart-color fidelity verified in Phase 1.
- **Consumer call-site coverage** — any missed `get_contract` call leaves a two-track read; mitigated by a grep audit in Phase 3.
- **Templated default template bloat** — embedding adds ~tens of KB to the committed `template.pptx`; acceptable, size logged.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 4, US-4.1.
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 US-4.1 (🟡 Partial), §5 Decision 1/2.
- Predecessors: US-1.5 (issue #55), US-3.1 (issue #56).
- PLAN format template: `PLANS/PLAN-GIT-56.md`.
- Render-path / data-model gap map: cited inline (`ppt_builder.py:859`, `template_introspector.py:182-207`, `schema_extractor.py:724-833`).
