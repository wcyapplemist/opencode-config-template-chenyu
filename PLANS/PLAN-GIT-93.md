# PLAN-GIT-93 — slide_type decoupling: make any template layout targetable via layout_name (hybrid A+C)

**Issue**: #93 (parent) — sub-issues #94 (Phase 1), #95 (Phase 2), #98 (Phase 3), #96 (Phase 4), #100 (Phase 5), #99 (Phase 6), #97 (Phase 7)
**Branch**: feat/slide-type-decoupling-93
**Priority**: High (P0) — unblocks all templates with >8 layouts
**Status**: Planning — Rev 2 PASSED architecture re-review (GO). Incorporates GO-WITH-CONDITIONS amendments (3 MAJOR closed) + 4 re-review NF clarifications (LOW/NIT). Ready to implement.

## Goal

Make the engine's `slide_type` field a **semantic label** (backward compatible) instead of a **hard gate** for layout selection. Introduce a first-class `layout_name` field that points directly to any template layout, so a template with 35 layouts can target all 35 — not just the 8 with matching fingerprints.

The chosen solution is **HYBRID A+C**: keep the 8-type enum as a recommended set, add `layout_name` as an optional overlay, and switch fill dispatch from type-membership gates (`_LAYOUTS_WITH_*`) to **field-presence + placeholder-availability** checks. Purely additive — no contract break, no 649-test rewrite.

## Motivation

The engine hard-limits `slide_type` to exactly 8 values. This contradicts **US-4.1**'s original design intent: *"通过 layout_name 匹配确定使用哪个幻灯片版式"*. A template with 35 layouts can only use ~8; the other 27 are unreachable.

### Decoupling Table

| Concern | Before | After |
|---|---|---|
| `slide_type` | Hard gate — unknown types rejected/skipped | Semantic label only — recommended 8 types, free-form when `layout_name` present |
| Layout selection | Fingerprint matching on 8 ideal profiles | `layout_name` → config-pin → fingerprint → name fallback → degradation |
| Fill dispatch | `if slide_type in _LAYOUTS_WITH_*` type gates | Field-presence + placeholder-availability checks |
| Validation | Per-type schema — unknown type validates nothing | Per-type schema for known types; generic per-field schema for `layout_name` path |

## Dependency & Consumer Map

```
Phase 1 (contract) ──→ Phase 2 (engine: layout selection)
                        └──→ Phase 3 (engine: fill dispatch) ──[HARD]──→ Phase 5 (multipass simplify)
                                  ‖ (same PR)
                             Phase 4 (validator)
                                  └──→ Phase 6 (estimators)
Phase 7 (docs) ───────→ depends on all above
```
(NF-4 fix: tree now reflects P3→P5 hard edge and P3∥P4 same-PR, matching the textual constraints below.)

**Consumers affected (architecture-review amended — adds 2 missed consumers):**
- `pptx-subagent.md` (agent routing, Stage 0/−1)
- `generate-slide-skill/SKILL.md` (slide_data schema docs)
- `AGENTS.md` (project notes)
- `multipass_render.py` (batch-to-engine pipeline)
- `overflow_check.py`, `density_mode.py` (estimators)
- **`ppt_builder._validate_template` (AC6 fatal gate)** — architecture-review MAJOR-2: blocks `layout_name`-only templates; relaxed in T2.5
- **`overflow_check._available_height_for_field`** — architecture-review MINOR-1: resolves placeholder geometry by `slide_type`; must switch to `layout_name` in T6.1

## Implementation Phases

> Dependency graph: `Phase 1 → Phase 2 → Phase 3 → Phase 4 → Phase 6 → Phase 5 → Phase 7`
>
> **Hard ordering constraints (architecture-review NIT-2):**
> - `Phase 3 → Phase 5` is a HARD edge: after Phase 2 relaxes the gate, pseudo-typed slides reach the fill loop but are not in `_LAYOUTS_WITH_*` — if Phase 5 (stop pseudo-typing) ships before Phase 3 (field-driven fill), those slides render EMPTY. Phase 3 MUST land before (or with) Phase 5.
> - `Phase 3 ∥ Phase 4` ship in the SAME PR: between Phase-3 merge and Phase-4 merge, unknown-type slides are filled but unvalidated (a window of malformed-but-rendered slides). Land them together.

### Phase 1: Contract layer — `available_layouts()` + layout classifier — #94
- [ ] T1.1: NEW `available_layouts(contract) -> List[dict]` in `_common/scripts/layout_contract.py` — returns ALL template layouts `{name, index, fingerprint, content_area_in2}`. Iterates `contract["layouts"]` (not limited to 8 types).
  — **Why:** Without this the agent cannot discover which layouts to target in a user's template.
  — **Done when:** `available_layouts` on a 35-layout contract returns 35 entries.
  — **Consumers affected:** `pptx-subagent.md` Stage 0, `generate-slide-skill/SKILL.md`.
- [ ] T1.2: NEW (optional helper) `classify_layout_fingerprint(fp) -> str` — maps an arbitrary fingerprint to the nearest standard slide_type by reusing `_composition_diff` against the 8 ideal fingerprints (min missing).
  — **Why:** Gives the agent a semantic hint for non-standard layouts.
  — **Done when:** Maps a non-standard layout to the closest of the 8 standard types.
  — **Consumers affected:** Agent prompt (informational).
- [ ] T1.3: Verify `servable_slide_types` unchanged (still returns 8).
  — **Why:** Backward compat gate.
  — **Done when:** Existing 649+ tests pass.
  — **Consumers affected:** None (verification only).

### Phase 2: Engine — `layout_name` first-class + gate relaxation (+ latent bug fix) — #95
- [ ] T2.1: Modify `_select_layout` in `generate-slide-skill/scripts/ppt_builder.py` (~lines 310-357): add `layout_name` as the **highest-precedence** selector (step 0, before config-pin). New order: `layout_name` → config-pin → fingerprint → name fallback → degradation.
  — **Why:** This is the core mechanism that makes any layout targetable.
  — **Done when:** `{"slide_type":"agenda","layout_name":"Agenda","title":"X"}` renders against the template's Agenda layout.
  — **Consumers affected:** `ppt_builder.py` render loop, `multipass_render.py`.
  — **Precedence policy (architecture-review NIT-1, DEFAULT pending user confirm):** `layout_name` intentionally outranks the operator `.config.json` config-pin (`<slide_type>_layout`) — **per-slide authoring wins over deck-wide operator config**. This is a policy inversion vs the prior config-pin-wins behavior; must be documented in SKILL.md (Phase 7). NOTE: `template-modifier-skill`'s `resolve_and_clone` routes cloned layouts via config-pin — confirm it does not rely on config-pin beating an explicit `layout_name` (open question #1).
- [ ] T2.2: The caller render loop (~line 1454) passes `slide_data.get("layout_name")` into `_select_layout`.
  — **Why:** The field must flow from slide_data to the layout selector.
  — **Done when:** `layout_name` is threaded through the call chain.
  — **Consumers affected:** `ppt_builder.py` render loop.
- [ ] T2.3: Relax the gate at ~line 325 (`if slide_type not in _SLIDE_TYPE_FINGERPRINT ...: return None`) so an unknown slide_type is NOT rejected when a `layout_name` or config-pin rescues it.
  — **Why:** Unknown types + `layout_name` must reach the config-pin/layout_name branches.
  — **Done when:** Unknown slide_type + layout_name does NOT return None.
  — **Consumers affected:** `ppt_builder.py` layout selection.
- [ ] T2.4: **BUG FIX** — this also fixes the structurally-broken pseudo-type escape hatch. `multipass_render._batch_to_engine_slides` creates pseudo-types `_custom_N` + config pins, but `_select_layout`'s gate rejects them BEFORE the config-pin branch runs, so those slides were silently dropped. This phase makes that path functional.
  — **Why:** Latent bug: multipass-rendered slides with pseudo-types were never rendered. End-to-end tests deliberately avoided pseudo-types.
  — **Done when:** `test_pseudo_type_bug_fixed` passes — pseudo-type + config-pin now renders.
  — **Consumers affected:** `multipass_render.py`, `test_multipass_render_merge.py`.
- [ ] T2.5: **(architecture-review MAJOR-2)** Relax `_validate_template` AC6 fatal gate in `ppt_builder.py` (~lines 420-430): currently raises `TemplateError` when `servable_slide_types(contract)` returns ZERO of the 8 standard types. Since `servable_slide_types` iterates `_SLIDE_TYPE_FINGERPRINT` only (no notion of `layout_name`), a template whose layouts fingerprint-match none of the 8 ideals (exotic/custom layouts — precisely the use case this feature serves) **aborts before any slide renders**, even when every slide carries a valid `layout_name`.
  — **Relaxation (DEFAULT pending user confirm on open question #2):** when `len(prs.slide_layouts) > 0` AND at least one slide in `slide_data_list` carries a `layout_name`, demote "serves 0 of 8" from fatal `TemplateError` to a warning. Keep it fatal when NO slide carries `layout_name` (preserves the guard for pure-8-type decks).
  — **Why:** Without this, the headline feature (AC1) fails for the exact exotic-template inputs it targets.
  — **Done when:** A template serving 0 of 8 standard types renders successfully when ≥1 slide carries `layout_name`; `test_zero_servable_template_renders_via_layout_name` passes.
  — **Consumers affected:** `ppt_builder._validate_template`, AC6/AC8 tests.

### Phase 3: Engine — field-driven fill dispatch (deprecate `_LAYOUTS_WITH_*` type gates) — #98
- [ ] T3.1: **(architecture-review MAJOR-1)** Replace the 4 `if slide_type in _LAYOUTS_WITH_*` gates in the fill loop (~lines 1491-1593) with field-presence + placeholder-availability checks. **CRITICAL: field-driven fill must distinguish FILL from SWEEP** — the current subtitle block (`ppt_builder.py:1485-1516`) calls `_remove_placeholder(sub_ph)` (L1515) when subtitle text is absent, to prevent the master's inherited sample text ("Click to edit Master subtitle style") from bleeding in (an empty `<a:p/>` still inherits — empirically verified). A naive `if slide_data.get("subtitle") and placeholder` drops this guard and regresses the standard `title_slide`/`closing_slide`/`section_header_slide` happy path.
  - subtitle: `if slide_data.get("subtitle") and _find_placeholder(slide, _SUBTITLE_TYPE)` → fill; **elif a SUBTITLE placeholder exists and is master-inherited-sample → `_remove_placeholder`** (preserve L1515). **(NF-3)** For `slide_type == "closing_slide"`, the subtitle text MUST be resolved via `_compose_signoff(slide_data)` first (ppt_builder.py:1007-1023, 1494-1497) — i.e. `subtitle_text = _compose_signoff(slide_data) or slide_data.get("subtitle", "")` — so a closing slide with no explicit subtitle but presenter fields still renders the sign-off (don't let the sweep drop it).
  - body: `if slide_data.get("body") and _find_body_placeholder(slide)` → fill
  - two-body: `if (body_left or body_right) and len(_find_placeholders(slide, _OBJECT_TYPE)) >= 2` → fill
  - chart: `if slide_data.get("chart_type")` → fill
  — **Why:** Type-gated fill prevents unknown slide_types from filling any content. Field-driven fill is template-aware. The sweep preserves the bleed-protection that was a side-effect of the type gate.
  — **Done when:** (a) unknown slide_type "agenda" with a `body` field fills the layout's OBJECT placeholder correctly; (b) a `title_slide` with NO `subtitle` still has its subtitle placeholder REMOVED (no sample bleed) — asserted by `test_subtitle_removed_when_empty`.
  — **Consumers affected:** `ppt_builder.py` fill loop (all slide types).
  — **Future refactor (open question #4, OPTIONAL):** a small dispatch table `(placeholder-composition) → fill-actions` would let the subtitle sweep + two-body fallback (L1553-1590) share one path and lower the regression surface. Not blocking; defer if time-boxed.
- [ ] T3.2: title/image/notes are already field-driven — leave them unchanged.
  — **Why:** No change needed; already correct.
  — **Done when:** Existing title/image/notes behavior verified.
  — **Consumers affected:** None.
- [ ] T3.3: Deprecate (comment out or remove) `_LAYOUTS_WITH_SUBTITLE/_BODY/_TWO_BODIES/_CHART` as gates.
  — **Why:** Dead code after T3.1. Safety rationale: a `title_slide` layout has no OBJECT placeholder → body fill finds nothing → natural no-op (backward-compatible).
  — **Done when:** No `_LAYOUTS_WITH_*` gate references remain in the fill loop.
  — **Consumers affected:** `ppt_builder.py` constants.
- [ ] T3.4: Multi-slot layouts (>=3 slots) reuse the EXISTING `placeholder_backfill.backfill_deck` (`body_slots`/`image_paths`/`title_slots`) — do NOT rewrite it.
  — **Why:** The backfill mechanism already handles multi-slot fills correctly.
  — **Done when:** Documented that >=3-slot layouts should use `body_slots`.
  — **Consumers affected:** Documentation only.

### Phase 4: Validator — generic per-field schema for the `layout_name` path — #96
- [ ] T4.1: In `generate-slide-skill/scripts/schemas/slide_schemas.py`: add `ALL_FIELD_SPECS` = union of all required+optional field specs across the 8 schemas (dedup by field name) + extended fields (`layout_name`:string, `body_slots`:array, `image_paths`:array, `title_slots`:array).
  — **Why:** Unknown slide_types need validation too — currently they validate nothing.
  — **Done when:** `ALL_FIELD_SPECS` covers all fields from the 8 schemas + extended fields.
  — **Consumers affected:** `schema_validator.py`.
- [ ] T4.2: In `generate-slide-skill/scripts/schema_validator.py` (~line 236): in the unknown-slide_type branch, when `slide_data.get("layout_name")` is present, validate each present field against `ALL_FIELD_SPECS` via the existing `_validate_field` (instead of the current early-return that validates nothing).
  — **Why:** Closes the validation gap for `layout_name`-targeted slides.
  — **Done when:** `{"slide_type":"team","layout_name":"Team","title":"X"}` validates clean.
  — **Consumers affected:** `schema_validator.py`.
- [ ] T4.2b: **(architecture-review MAJOR-3)** Per-field validation only checks type/shape of PRESENT fields — it enforces NO requiredness. Per-type schemas mark `title` + `notes` as `required` (slide_schemas.py:98-151). Under the generic path, `{"slide_type":"team","layout_name":"Team"}` (no title, no notes) would validate clean → renders an empty-title slide (master-inherited sample bleed) + notes-less slide. After T4.2's per-field validation, enforce a MINIMAL required set in the `layout_name` branch: `title` recommended (warn-if-absent) and `notes` recommended. Reuse the existing `recommended: True` mechanism (non-fatal but visible).
  — **Why:** Required-field enforcement is a load-bearing invariant being dropped; the chart pair (T4.3) is the only cross-field invariant the plan currently preserves.
  — **Done when:** `{"slide_type":"team","layout_name":"Team"}` (no title/notes) emits warnings for missing `title` and `notes`; `test_layout_name_slide_missing_title_warns` passes.
  — **Consumers affected:** `schema_validator.py`.
  — **Invariant downshift (document explicitly):** unknown-type + layout_name loses the per-type REQUIRED guarantee in favor of recommended-warnings. This is an accepted trade-off for layout flexibility.
- [ ] T4.3: Enforce chart pairs: if `chart_type` present, require `categories`+`series`.
  — **Why:** Chart slides need paired data even with unknown types.
  — **Done when:** A chart slide missing `series` reports an error even with an unknown slide_type.
  — **Consumers affected:** `schema_validator.py`.
- [ ] T4.4: Unknown-type WITHOUT `layout_name` keeps current behavior (warn "will be skipped").
  — **Why:** Backward compat — these slides will still be skipped by the engine.
  — **Done when:** Existing warn-on-unknown-type behavior preserved.
  — **Consumers affected:** None.

### Phase 5: Multipass simplification (>8-layout limit disappears as byproduct) — #100
- [ ] T5.1: **(architecture-review MINOR-2)** `layout_name` is now native (Phase 2), so a single pass renders N distinct layouts — the 8-layout/batch ceiling no longer exists. **Deletion decision (DEFAULT pending user confirm on open question #3):** DELETE `_batch_to_engine_slides` (~L220-263) and `partition_slides` — their only purpose (pseudo-typing `_custom_N`) is obsolete post-GIT-93 and keeping them is a maintenance trap (the module docstring L1-26 now actively misleads: it claims to "solve L1 ... >8 distinct layouts"). KEEP `merge_decks` (independently useful for stitching pre-rendered decks). Rewrite the module docstring so it no longer claims to solve the >8-layout ceiling.
  — **Why:** Vestigial pseudo-type machinery whose remaining purpose is obsolete; a misleading docstring is worse than no docstring.
  — **Done when:** `_batch_to_engine_slides`/`partition_slides` removed; `merge_decks` retained; module docstring updated; `multipass_render` reduced to a single-pass call (delegating to `generate_ppt_from_data`) + `merge_decks` for explicit multi-deck cases; 10 slides with distinct `layout_name` values render in a single pass to 10 slides.
  — **Consumers affected:** `multipass_render.py` (+ its tests, simplified).
  — **Soft-deprecated alternative (if user prefers one-release-cycle deprecation):** mark the two functions `# DEPRECATED post-GIT-93, remove in next release` instead of deleting.
- [ ] T5.2: `merge_decks` still works for explicit multi-deck merge scenarios.
  — **Why:** Backward compat — some callers may still want deck merging.
  — **Done when:** Existing multipass tests still pass.
  — **Consumers affected:** `multipass_render.py` tests.
- [ ] T5.3: Stage -1 engine-limits auto-routing should NOT force multipass for >8 layouts anymore (only when a caller explicitly wants to merge pre-rendered decks). Prompt update deferred to Phase 7.
  — **Why:** The >8-layout trigger is now unnecessary; `layout_name` handles it natively.
  — **Done when:** Documented in Phase 7 agent prompt update.
  — **Consumers affected:** `pptx-subagent.md` (Phase 7).

### Phase 6: Downstream estimators — unknown-type safety — #99
- [ ] T6.1: **(architecture-review MINOR-1)** `overflow_check.py` — the real defect is that overflow geometry is keyed by `slide_type`, not `layout_name`. `_available_height_for_field` (~L93-96) looks up `contract["layouts_by_slide_type"][slide_type]` → returns `None` for unknown types → verdict deferred-FIT (rich-body unknown-type slides silently get FIT, no overflow estimate). Rewrite to resolve placeholder geometry by **`layout_name` → `contract["layouts"][i]` → placeholders** (the contract exposes `layouts[i]` with full placeholder geometry, not just `layouts_by_slide_type`). Unknown-type slides then get real overflow estimates against their actual targeted layout.
  — **Why:** The prior plan ("default to content-like path") addressed a symptom; `count_slide_words` already counts unknown types. The root cause is geometry lookup by slide_type. **(NF-1)** Note the contract emits the `"type"` key per placeholder (`contract_adapter.py:111-118`: `"OBJECT"`, `"SUBTITLE"`), but `_available_height_for_field` reads `ph.get("role","")` — the rewrite must read `ph.get("type")` (or map); the existing match tuple already includes `"OBJECT"`. **(NF-2)** Bonus bug-fix: `layouts_by_slide_type` is a DEAD key (read by overflow_check, produced by NO module) — this rewrite incidentally retires it.
  — **Done when:** Unknown-type slide with a `layout_name` gets a real overflow verdict against that layout's geometry (not deferred-FIT); `test_overflow_unknown_type_uses_layout_name_geometry` passes.
  — **Consumers affected:** `overflow_check.py`, `overflow_check._available_height_for_field`.
- [ ] T6.2: **(architecture-review MINOR-1 — clarified)** `density_mode.count_slide_words` (~L104-121) iterates `WORD_COUNT_FIELDS` **unconditionally by field, not by slide_type** — unknown-type slides with a `body` are ALREADY counted; `validate_density` (~L141-153) walks the full list regardless of type. The prior plan's premise ("silently return 0") is false — this is a NO-OP. **No code change needed.** Add only a regression test `test_density_counts_unknown_type_body` to lock the existing correct behavior.
  — **Why:** Confirms density already handles unknown types; prevents future regression.
  — **Done when:** Regression test passes (no behavior change expected).
  — **Consumers affected:** `density_mode.py` (test-only).

### Phase 7: Documentation + agent prompt updates — #97
- [ ] T7.1: `generate-slide-skill/SKILL.md` (~line 107): document `slide_type` = "8 standard semantic types (recommended) OR a free-form label when `layout_name` is provided". Add `layout_name` field reference + `available_layouts` section with examples (agenda page, team page, timeline page).
  — **Why:** Users and agents need to know the new field and its semantics.
  — **Done when:** SKILL.md documents `layout_name` with 3+ concrete examples.
  — **Consumers affected:** Agent users of the skill.
- [ ] T7.2: `.opencode/agents/pptx-subagent.md` Stage -1: drop forced multipass for >8 distinct layouts.
  — **Why:** `layout_name` makes multipass unnecessary for layout-count reasons.
  — **Done when:** Stage -1 no longer forces multipass based on layout count.
  — **Consumers affected:** Agent routing logic.
- [ ] T7.3: `.opencode/agents/pptx-subagent.md` Stage 0: use `available_layouts` to discover layouts; add routing for custom `slide_type` + `layout_name`.
  — **Why:** The agent needs to discover what layouts exist before generating slide_data.
  — **Done when:** Stage 0 references `available_layouts` and shows `layout_name` routing.
  — **Consumers affected:** Agent Stage 0 flow.
- [ ] T7.4: `AGENTS.md`: append a "slide_type decoupling" note under the BT-142 section.
  — **Why:** Project-level documentation must reflect the architectural change.
  — **Done when:** AGENTS.md has a decoupling summary paragraph.
  — **Consumers affected:** Project documentation.

## Acceptance Criteria (per phase + overall)

Overall (parent #93):
- [ ] AC1 — A template with >8 layouts can target ALL of them via `layout_name` (not limited to 8).
- [ ] AC2 — The existing 8-type path is backward compatible: all 649+ existing tests pass unchanged.
- [ ] AC3 — `slide_type` remains a valid semantic label for the 8 standard types.
- [ ] AC4 — Fill dispatch is driven by field-presence + placeholder-availability, not type-membership gates.
- [ ] AC5 — Validation works for unknown slide_types when `layout_name` is present.
- [ ] AC6 — The latent pseudo-type bug in multipass is fixed.
- [ ] AC7 — Downstream estimators (overflow, density) handle unknown types safely.
- [ ] AC8 — **(architecture-review MAJOR-2)** A template serving 0 of 8 standard types still renders when ≥1 slide carries `layout_name` (no fatal `TemplateError`); remains fatal when no slide carries `layout_name`.
- [ ] AC9 — **(architecture-review MAJOR-1)** Subtitle bleed-protection preserved: a standard `title_slide`/`closing_slide` with no `subtitle` still has the placeholder REMOVED (no master-inherited sample text bleed) after the field-driven fill refactor.

Per-phase gates:
- Phase 1 (#94): `available_layouts` returns all layouts; `servable_slide_types` unchanged; classifier maps to nearest type.
- Phase 2 (#95): `layout_name` overrides fingerprint; unknown type + `layout_name` renders; pseudo-type + config-pin renders (bug fix).
- Phase 3 (#98): Field-driven fill works for unknown types; `title_slide` with stray `body` does NOT mis-fill; `_LAYOUTS_WITH_*` gates removed.
- Phase 4 (#96): Generic schema validates `layout_name` path; chart pair enforcement works for unknown types.
- Phase 5 (#100): 10 `layout_name` slides render in single pass; `merge_decks` backward compat preserved.
- Phase 6 (#99): Unknown-type slides get overflow + density estimates.
- Phase 7 (#97): SKILL.md, pptx-subagent.md, AGENTS.md all updated with `layout_name` docs.

## Test Plan

New file `generate-slide-skill/scripts/tests/test_layout_name_targeting.py`:
- `test_layout_name_overrides_fingerprint` — same slide_type, different layout_name → different layouts used
- `test_unknown_slide_type_with_layout_name_renders` — agenda+layout_name not skipped
- `test_unknown_slide_type_without_layout_name_skipped` — current degradation preserved
- `test_available_layouts_returns_all` — 35-layout contract → 35 entries
- `test_generic_schema_validates_layout_name_path` — chart missing series → error
- `test_field_driven_fill_unknown_type` — agenda layout [TITLE,OBJECT] + body → filled
- `test_multipass_passthrough_with_native_layout_name` — 10 layout_name slides, single pass → 10 slides
- `test_pseudo_type_bug_fixed` — pseudo-type + config-pin now renders (regression for the Phase 2 bug fix)
- `test_subtitle_removed_when_empty` — **(architecture-review MAJOR-1)** standard `title_slide` with no subtitle still removes the placeholder (no sample bleed)
- `test_zero_servable_template_renders_via_layout_name` — **(architecture-review MAJOR-2)** template serving 0 of 8 types renders when slides carry `layout_name`
- `test_layout_name_slide_missing_title_warns` — **(architecture-review MAJOR-3)** unknown-type + layout_name with no title/notes emits recommended-warnings
- `test_overflow_unknown_type_uses_layout_name_geometry` — **(architecture-review MINOR-1)** unknown-type overflow resolved by layout_name geometry
- `test_density_counts_unknown_type_body` — **(architecture-review MINOR-1)** locks existing correct density behavior for unknown types

Modify existing tests that assert "unknown type skipped" to distinguish with/without layout_name.

Regression: all 8-type paths' 649+ tests unchanged (additive).

## Risks

- **R1 (HIGH) — Rich-layout placeholder mapping ambiguity (>=3 slots).** Mitigated by reusing `body_slots` backfill; single/double body via field-driven fill; document ">=3 slots → body_slots".
- **R2 (MEDIUM) — Weaker validation guarantees for unknown types.** Mitigated by Phase 4 generic per-field schema + chart pair enforcement.
- **R3 (MEDIUM) — Agent doesn't know available layouts.** Mitigated by Phase 1 `available_layouts()` + fingerprint classifier (usability prerequisite).
- **R4 (MEDIUM) — Pseudo-type path untested end-to-end.** Mitigated by Phase 5 + `test_pseudo_type_bug_fixed`.
- **R5 (LOW-MED) — Downstream estimators on unknown types.** Mitigated by Phase 6 default-to-content-like path.
- **R6 (LOW) — Backward compat regression.** Purely additive; old path untouched; 649+ tests unchanged.
- **R7 (HIGH — architecture-review MAJOR-1) — Subtitle/placeholder bleed-protection regression.** Field-driven fill naively drops the `_remove_placeholder` guard (L1515). Mitigated by T3.1's fill-vs-sweep split + `test_subtitle_removed_when_empty`.
- **R8 (HIGH — architecture-review MAJOR-2) — Exotic-template fatal abort.** `_validate_template` AC6 rejects templates serving 0 of 8 types before render. Mitigated by T2.5 relaxation + `test_zero_servable_template_renders_via_layout_name`.
- **R9 (MEDIUM — architecture-review MAJOR-3) — Validator required-invariant loss.** Generic per-field validation drops title/notes requiredness. Mitigated by T4.2b recommended-warnings + documented invariant downshift.

## Non-Goals

- Does **NOT** abolish the 8-type enum (that's Option C — full removal).
- Does **NOT** template-derive a type set (that's Option B — auto-discovery).
- Does **NOT** rewrite `placeholder_backfill` for multi-slot layouts (reuses existing `body_slots`/`image_paths` mechanism).
- Does **NOT** change the render contract JSON schema (additive fields only).

## Architecture Review Amendments (Rev 2)

Rev 2 incorporates the GO-WITH-CONDITIONS findings from the architecture-review-subagent pass. Changes:

| # | Severity | Amendment | Where |
|---|----------|-----------|-------|
| MAJOR-1 | High | T3.1 rewritten: field-driven fill splits FILL from SWEEP; subtitle removal-on-empty (L1515) preserved | Phase 3 |
| MAJOR-2 | High | T2.5 added: relax `_validate_template` AC6 when ≥1 slide carries `layout_name` | Phase 2 |
| MAJOR-3 | High | T4.2b added: enforce `title`/`notes` recommended-warnings in layout_name validator branch | Phase 4 |
| MINOR-1 | Med | T6.1 rewritten: overflow geometry by `layout_name → layouts[i]`; T6.2 clarified as test-only no-op | Phase 6 |
| MINOR-2 | Med | T5.1 rewritten: DELETE `_batch_to_engine_slides`/`partition_slides`, keep `merge_decks`, fix docstring | Phase 5 |
| NIT-1 | Low | T2.1: documented `layout_name > config-pin` precedence policy inversion | Phase 2 |
| NIT-2 | Low | Dependency graph annotated: P3→P5 hard edge, P3∥P4 same PR | Dependency Map |
| — | — | Consumers list +2: `_validate_template`, `overflow_check._available_height_for_field` | Consumer Map |
| — | — | AC8/AC9 + 4 new tests + R7/R8/R9 added | AC / Tests / Risks |

## Open Questions (defaults applied, pending user confirm)

1. **Config-pin vs `layout_name` precedence (NIT-1).** DEFAULT: per-slide `layout_name` wins over operator `.config.json` config-pin. Needs confirm that `template-modifier-skill`'s `resolve_and_clone` (routes cloned layouts via config-pin) does not rely on config-pin beating an explicit `layout_name`.
2. **AC6 relax boundary (MAJOR-2).** DEFAULT: keep "serves 0 of 8" fatal ONLY when no slide carries `layout_name`; warn when ≥1 slide has `layout_name`.
3. **Multipass deletion (MINOR-2).** DEFAULT: delete `partition_slides`/`_batch_to_engine_slides` now, keep `merge_decks`. Alternative: soft-deprecate one cycle.
4. **Fill-strategy abstraction (optional).** DEFAULT: keep inline `if` branches + sweep for now; a `(placeholder-composition) → fill-actions` dispatch table noted as a future refactor to lower the MAJOR-1 regression surface. Not blocking.

## References

- Requirements source: issue #93 — confirmed hybrid A+C solution.
- Sub-issues: #94 (Phase 1), #95 (Phase 2), #98 (Phase 3), #96 (Phase 4), #100 (Phase 5), #99 (Phase 6), #97 (Phase 7).
- Related: US-4.1 (layout selection by layout_name), BT-142 Phase 3.5 L1 (multipass_render).
- PLAN format template: `PLANS/PLAN-GIT-78.md`.