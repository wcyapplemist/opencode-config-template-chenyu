# PLAN-GIT-72 — Epic 5: Shared `common/` Package (US-5.2 Foundation + Parasitism Fix)

**Issue**: #72
**Branch**: GIT-72 (base: dev)
**Priority**: Should Have (Epic 5 — Skill Architecture)
**Status**: Planned (rev 2 — architecture-review APPROVE-WITH-CHANGES incorporated: C1 extraction closure 9→13, M1 schemas split, M2 template-modifier test sites, M3 zero-coupling reframe, m1/m2/m3/m4/m6; m5 shared-bootstrap deferred to Phase B)

## Goal

Extract the template-extraction / contract / schema infrastructure that is shared by multiple skills into a single `.opencode/skills/_common/scripts/` package, so that `template-modifier-skill`'s **production code** has **zero** coupling to `generate-slide-skill`, `generate-template-skill` resolves `schema_extractor` via `_common`, and `generate-slide-skill` stops carrying extraction infra that is not part of "filling a template". This delivers the shared-schema home prescribed by US-5.2 and advances US-5.1. **No skills are split** — shared code is consolidated, not fragmented.

## Strategic Context

`generate-slide-skill/scripts` is 10 modules / ~4.4k lines (schema_extractor 1362 + ppt_builder 1600 dominate). But a dependency audit shows much of it is cross-skill infra, parasitically nested under one skill:

- `generate-template-skill` has **no scripts** — borrows `schema_extractor`.
- `template-modifier-skill`'s 3 production scripts each `sys.path.insert(...generate-slide-skill/scripts)` and borrow `ppt_builder.{servable_slide_types, _resolve_layout_by_fingerprint, _SLIDE_TYPE_FINGERPRINT, get_render_contract}` + `schema_extractor.{extract_schema, embed_schema, read_embedded_schema}` + `template_introspector.introspect`. (Note: `generate_ppt_from_data` is referenced only in `state_machine.py`'s **docstring** + imported only in `tests/test_layout_creator.py` — **not** a production coupling. M3.)

Internal layering of `generate-slide-skill/scripts`:
```
Layer 0 pure leaves : geometry, density_mode, text_fit, template_introspector, outline_store, coordinate_placer
Layer 1 extract/val : schema_extractor(→geometry), contract_adapter(→geometry), schema_validator(→density_mode)
Layer 2 render hub  : ppt_builder(→几乎所有)  ← 1600-line god-module
```
Layers 0–1 (extraction/contract) are cross-skill; Layer 2 (render/fill) is generate-slide-skill's own job. Verified: `density_mode`/`text_fit`/`outline_store`/`coordinate_placer` are referenced **only** within generate-slide-skill → correctly stay.

## Architecture Decisions (locked)

1. **Consolidate, don't split** — extract a shared `_common/` package; do NOT create new skills (aligns with US-5.1 "expose exactly 2 skills" + chenyu's 2-skill intent). The pain points (maintainer cognitive load + shared-infra parasitism) are code-level, not routing-level.
2. **Location**: `.opencode/skills/_common/scripts/`. **Trade-off note (m4):** kept under `skills/` for relative-path convenience (`../_common/scripts`); a sibling `.opencode/common/` would be semantically cleaner but gains little. **Discovery risk (m3) is low**: OpenCode discovers skills by `SKILL.md` presence (not dir name) — a `_common/` without `SKILL.md` is invisible to discovery; the `_` prefix is defense-in-depth.
3. **What moves to `_common/scripts/`**: `schema_extractor.py`, `geometry.py`, `contract_adapter.py`, `template_introspector.py`, and **only `schemas/template_schema.json`** (M1 — NOT the whole `schemas/` dir), plus a **new `layout_contract.py`** extracting the pure contract logic out of `ppt_builder`. `validate_template_schema` (shared, extraction-side) lives in `schema_extractor` and moves with it → both skills can validate the template schema (US-5.2 AC1/AC2). `validate_slide_data_list` (fill-side) stays in `schema_validator` (review-focus #4: `schema_validator` does NOT belong in `_common`).
4. **`layout_contract.py` public surface — 13 symbols (C1 closure fix)**, moved from ppt_builder. All pure (operate on the contract dict / read embedded JSON), grouped:
   - Constants: `_SLIDE_TYPE_FINGERPRINT`, `_SERVES_LAYOUT`, `_LAYOUT_NAME_MAP`, `_LAYOUT_TYPES_NEEDING_SIDEBYSIDE`.
   - Pure helpers: `_composition_diff`, `_name_affinity`, `_content_placeholders_stacked`, `_normalize_layout_name`.
   - Contract entrypoints: `_resolve_layout_by_fingerprint`, `servable_slide_types`, `get_render_contract`.
   - Embedded-schema probe helpers (C1): `_warn_if_embedded_stale`, `_live_layout_count` (Q1: lives here, single contract-probe file).
   `get_render_contract`'s whole call chain (`read_embedded_schema`, `_warn_if_embedded_stale`, `_live_layout_count`, `embedded_schema_to_contract`, `get_contract`) now resolves **inside `_common`** → closed loop. The 4 closure additions (C1) are required: without them `layout_contract.py` back-imports from ppt_builder → circular import.
   **Back-imports (correct direction, fill → `_common`):** ppt_builder's *staying* code re-imports `_LAYOUT_NAME_MAP` (used by `_select_layout`) and `_live_layout_count` (used by `_embedded_schema_stale` → `_ensure_output_templated`). `_embedded_schema_stale` itself stays in ppt_builder (it's on the fill/render path).
   **Invariant (test-matrix):** `grep -rE 'from ppt_builder|import ppt_builder' .opencode/skills/_common/` → **0 matches** (no `_common` → ppt_builder dependency).
5. **What stays in `generate-slide-skill`** (fill-specific, binds to live `prs`): `_build_layout_index`, `_resolve_layout`, `_select_layout`, `_embedded_schema_stale`, `_slide_dims_emu`, `_chart_bbox`, `_image_bbox`, `generate_ppt_from_data`, `schemas/slide_schemas.py` + `schemas/__init__.py`, and filling/charts/images/text-fit/density/outline/resolvers/coordinate_placer.
6. **Import-path convergence**: each skill's `conftest.py`/scripts add ONE bootstrap line putting `_common/scripts` on `sys.path`; module names stay unchanged ⇒ tests keep the same imports (resolved from a different path).
7. **template-modifier migration — 3 distinct sites (M2):**
   - **3 production scripts** (`constraint_checker.py`, `layout_creator.py`, `state_machine.py`): DELETE the `sys.path.insert(...generate-slide-skill...)` hack; import `layout_contract` + `schema_extractor` from `_common`. → **production coupling to generate-slide-skill becomes ZERO** (M3 reframe).
   - **`tests/conftest.py`**: **KEEP** `_FILLER_SCRIPTS` (legitimately backs `tests/test_layout_creator.py`'s `generate_ppt_from_data` import); **ADD** `_COMMON_SCRIPTS`.
   - **`tests/test_layout_creator.py`**: DELETE the inline redundant hack (conftest already injects the path); the import resolves via conftest.
   - The grep AC (M2) is scoped to the **3 production files only** (conftest/test legitimately retain `_FILLER_SCRIPTS`).
8. **Pure relocation + extraction** — no behavior change. The full suite (count confirmed at execution; was 465 at PLAN-GIT-70 delivery) is the safety net; green at every step (m2 — not a hardcoded count).

## Deliverables

**New**: `.opencode/skills/_common/scripts/{schema_extractor,geometry,contract_adapter,template_introspector,layout_contract}.py` + `schemas/template_schema.json`.
**Move** (git mv, history-preserving): the 4 modules + `template_schema.json` out of `generate-slide-skill/scripts/`. **`schemas/slide_schemas.py` + `schemas/__init__.py` stay** (M1).
**Change** `generate-slide-skill/scripts/ppt_builder.py`: remove the 13 moved contract symbols (import them from `layout_contract`); keep the fill-specific code.
**Change** `template-modifier-skill/scripts/{constraint_checker,layout_creator,state_machine}.py`: delete `sys.path` hack; import from `_common`. **Change** its `tests/conftest.py` (add `_COMMON_SCRIPTS`, keep `_FILLER_SCRIPTS`) + `tests/test_layout_creator.py` (delete inline hack).
**Change** `generate-template-skill/SKILL.md`: point all schema_extractor references at `_common` (m1 — 5 touch points: L39 prose + L69/L108/L127/L144 bash blocks).
**Change** each skill's `conftest.py`/bootstrap: add `_common/scripts` to `sys.path`.

## Acceptance Criteria

- [ ] 5 modules + `layout_contract.py` live under `.opencode/skills/_common/scripts/`; `generate-slide-skill` no longer carries extraction/contract infra.
- [ ] `grep -rE 'from ppt_builder|import ppt_builder' .opencode/skills/_common/` → 0 matches (no `_common` → ppt_builder back-import; C1 invariant).
- [ ] `template-modifier-skill`'s **3 production scripts** contain no `sys.path.insert(...generate-slide-skill...)` (conftest/test may retain `_FILLER_SCRIPTS`).
- [ ] `generate-template-skill` resolves `schema_extractor` via `_common` (all 5 SKILL.md touch points updated).
- [ ] Full suite green at every step; no behavior change (pure relocation + extraction).
- [ ] US-5.2 advances toward Met (shared schema + `validate_template_schema` in `common/`); US-5.1 partially advanced.

## Implementation Phases

### Phase 1: Create `_common/` + relocate modules (no behavior change)
- [ ] T1: create `.opencode/skills/_common/scripts/` + `schemas/`; `git mv` `schema_extractor.py`, `geometry.py`, `contract_adapter.py`, `template_introspector.py`, and `schemas/template_schema.json` (**only** the json — M1; `slide_schemas.py`/`__init__.py` stay).
- [ ] T2: add the bootstrap (each skill's `conftest.py` puts `_common/scripts` on `sys.path`); verify the moved modules import cleanly.
- [ ] T3: update `generate-slide-skill/scripts/{ppt_builder.py,tests/conftest.py}` to import the 4 moved modules from `_common`; run generate-slide-skill tests → green.

### Phase 2: Extract `layout_contract.py` (13 symbols) from `ppt_builder`
- [ ] T4: move the 13 contract symbols (Decision 4) into `_common/scripts/layout_contract.py`; ppt_builder imports them back (`_LAYOUT_NAME_MAP`, `_live_layout_count`, etc.); verify the C1 invariant grep = 0 and `get_render_contract`'s loop closes inside `_common`.
- [ ] T5: run generate-slide-skill tests → green (pure extraction, same behavior).

### Phase 3: Re-point the other two skills
- [ ] T6: `template-modifier-skill` — 3 prod scripts delete the hack + import from `_common`; `tests/conftest.py` keeps `_FILLER_SCRIPTS` + adds `_COMMON_SCRIPTS`; `tests/test_layout_creator.py` deletes its inline hack; run its tests → green.
- [ ] T7: `generate-template-skill/SKILL.md` — update all 5 schema_extractor touch points (L39 prose + L69/108/127/144 bash) to `_common`.

### Phase 4: Docs + full gate
- [ ] T8: AGENTS.md / README.md note the `_common/` package; note US-5.2 advancing (location/home delivered — not full AC1–AC4; GAP-ANALYSIS runtime-validation gap still open, m-Q4).
- [ ] T9: full suite — run **each skill's tests from its own dir independently (as CI would, m6)**, not only from repo root → green.

## Test matrix

| Case | Expected |
| --- | --- |
| relocate parity | moved modules import from `_common`; generate-slide-skill tests unchanged-green |
| C1 no back-import | `grep -rE 'from ppt_builder\|import ppt_builder' .opencode/skills/_common/` → 0 matches |
| layout_contract extraction | `get_render_contract` / fingerprint results byte-identical before/after extraction |
| template-modifier no prod hack | grep scoped to the 3 **production** scripts → 0 matches; conftest/test may retain `_FILLER_SCRIPTS` |
| template-modifier zero prod coupling | its 3 prod scripts import only from `_common` (+ stdlib) |
| generate-template resolution | all 5 SKILL.md touch points reference `_common/schema_extractor` |
| per-skill CI independence (m6) | each skill's suite passes when run from its own dir |
| full gate | all tests across the 3 skills pass (count confirmed at execution) |

## Verification

```powershell
# Per-skill (run from each skill's own dir as CI would — m6)
python -m pytest .opencode/skills/generate-slide-skill/scripts/tests/ -q
python -m pytest .opencode/skills/template-modifier-skill/scripts/tests/ -q
# C1 invariant
# (PowerShell) select-string -Path .opencode/skills/_common/scripts/*.py -Pattern 'from ppt_builder|import ppt_builder' -> none
```

## Risks

- **Import path churn** across 3 skills + tests → mitigated by keeping module names stable (only `sys.path` bootstrap changes) and the per-step test gate.
- **C1 circular import** → the 13-symbol closure (incl. `_warn_if_embedded_stale`/`_live_layout_count`) is the fix; enforced by the no-back-import grep.
- **`schemas/` over-move (M1)** → only `template_schema.json` moves; `slide_schemas.py` stays (fill-only).
- **template-modifier test breakage (M2)** → conftest KEEPS `_FILLER_SCRIPTS`; only the redundant inline test hack is deleted; grep AC scoped to production files.
- **`_common` discovery** → low risk (SKILL.md-keyed discovery; m3).
- **US-5.2 scope** → this delivers the shared-schema **home** + shared `validate_template_schema`; runtime `jsonschema` validation (GAP-ANALYSIS L236) remains open — "advances toward Met," not full AC1–AC4 (m-Q4).

## Out of scope (Phase B/C — later)

Splitting `ppt_builder.py` (1600 lines) internally; sub-packaging `generate-slide-skill`; new skills; **shared bootstrap helper** (`_common/scripts/_bootstrap.py` DRYing the 8+ `sys.path.insert` sites — m5).

## Architecture Review — rev 2 (APPROVE-WITH-CHANGES)

- **C1** (extraction closure 9→13) → Decision 4 + T4 + C1 invariant grep. **Resolved.**
- **M1** (schemas split) → Decision 3/Deliverables/T1: only `template_schema.json` moves. **Resolved.**
- **M2** (template-modifier test sites) → Decision 7 + T6 + scoped grep AC. **Resolved.**
- **M3** (zero production coupling reframe) → Decision 7 + Strategic Context. **Resolved.**
- **m1** (5 generate-template touch points) → T7. **Resolved.**
- **m2** (no hardcoded count) → Decision 8 + AC. **Resolved.**
- **m3** (`_common` discovery low-risk) → Decision 2. **Resolved.**
- **m4** (location trade-off) → Decision 2. **Resolved.**
- **m6** (CI per-skill verification) → T9 + test matrix. **Resolved.**
- **m5** (shared bootstrap) → deferred to Phase B.
- **Q1** (`_live_layout_count` home) → layout_contract.py (owner decision). **Resolved.**

**Confirmed strengths:** parasitism diagnosis accurate; "consolidate don't split" aligns with US-5.1; fill-only modules correctly stay; two-validator split correct for US-5.2; `get_render_contract` purity holds (modulo the C1 closure).

## References

- Borrowing today: `template-modifier-skill/scripts/{constraint_checker,layout_creator,state_machine}.py` (sys.path hack); `generate-template-skill` (no scripts).
- Key code: `generate-slide-skill/scripts/{schema_extractor,geometry,contract_adapter,template_introspector}.py` + contract symbols in `ppt_builder.py` (`get_render_contract` L1094, `_resolve_layout_by_fingerprint` L335, `servable_slide_types` L430, `_SLIDE_TYPE_FINGERPRINT` L127, `_warn_if_embedded_stale` L1084, `_live_layout_count` L1065, `_LAYOUT_NAME_MAP` L98, `_LAYOUT_TYPES_NEEDING_SIDEBYSIDE` L317).
- PLAN format template: `PLANS/PLAN-GIT-68.md` / `PLANS/PLAN-GIT-70.md`.
