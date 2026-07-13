# PLAN-GIT-78 — pptx-subagent: handle invalid/incomplete PPTX templates (missing master, missing layouts)

**Issue**: #78 (parent) — sub-issues #79 (Phase 0), #80 (Phase 1), #81 (Phase 2), #82 (Phase 3), #83 (Phase 4), #84 (Phase 5)
**Branch**: GIT-78
**Priority**: Medium-High (unblocks real-world user-supplied templates)
**Status**: Planned — design decisions locked with user; awaiting Phase 0 execution.

## Goal

Make the `pptx-subagent` robust against user-supplied `.pptx` files that are **not well-formed**: files with no slide master (Scenario A), files whose master lacks usable layouts for some slide types (Scenario B), and files with a minimal master carrying no usable placeholders (Scenario C). Today Scenario A is fatally rejected (`schema_extractor.py:882` raises `TemplateExtractionError`; `layout_creator.py:120`/`:171` crash with raw `IndexError`), and Scenarios B/C degrade silently (the over-limit clone is skipped at `layout_creator.py:201-203`).

The customer's **hard constraint** drives every decision: *preserve the user's existing master and styles — never ignore them.* The "throw away the user's file and use `default.pptx`" escape hatch is explicitly forbidden; the user's theme/fonts/branding must survive every repair path except the Level-3 last resort (where nothing of the user's file is left to preserve).

## Strategic Context

The engine has three crash/degradation points that this plan addresses:

1. **`schema_extractor.py:877-882`** — `if not masters: raise TemplateExtractionError("presentation has no slide master")`. A masterless deck never reaches generation.
2. **`layout_creator.py:120`** (`_clone_layout_into`) and **`:171`** (`_verify_layouts`) — bare `prs.slide_masters[0]` accesses crash with `IndexError` on a masterless input.
3. **`layout_creator.py:201-203`** (`clone_for_over_limit`) — when `_resolve_layout_by_fingerprint` returns `None` (no donor for a needed slide type), the clone is **skipped** and the user gets a deck with missing/mis-rendered slides, silently.

The infrastructure to repair/extend already partially exists and is **reused, not rebuilt**:

- `schema_extractor.read_embedded_schema` / `embed_schema` — zip-level part reading + order-preserving zip rewrite.
- `schema_extractor._raw_theme_colors_and_fonts` — theme XML parsing (clrScheme + fontScheme).
- `layout_creator._clone_layout_into` — 7-step layout clone (analogous pattern for a master clone).
- `layout_creator._max_layout_id` — ID allocation pattern.
- `layout_creator.clone_for_over_limit` — reload-verify + rollback safety model.
- `geometry.normalize_polygon` / `denormalize_polygon` / `aspect_ratios_match` — pure EMU scaling primitives (US-4.6).
- `default.pptx` — master skeleton + a theme that covers all 8 slide types (the geometry-borrowing source).

### Three-path dispatch

```
                         user-supplied .pptx
                                  │
                  ┌───────────────┴───────────────┐
                  ▼                               ▼
         prs.slide_masters empty?          master present?
            (Scenario A)               ┌──────────┴──────────┐
                  │                   ▼                     ▼
                  ▼            layouts missing        minimal master,
        _repair_if_needed        for a slide type?     no placeholders?
        (master_repairer)             (Scenario B)         (Scenario C)
                  │                       │                     │
   ┌──────────────┼──────────────┐        └──────────┬──────────┘
   ▼              ▼              ▼                   ▼
 L1 salvage   L2 scavenge    L3 default        master_cloner
 theme1.xml   slide <a:rPr/   fallback          (clone user master +
 from zip     a:spPr>         (default.pptx      borrow default.pptx
              styles          theme)             layout geometry,
              → best-effort   *only when          scale via geometry.py)
              theme           nothing survives*        │
   └──────────────┴──────────────┘                   ▼
            inject default.pptx              new layouts under clone
            master skeleton,                 (shared theme → user
            replace theme content            branding survives)
            with salvaged/scavenged
                  │                                │
                  ▼                                ▼
        _validate_template ───────────► clone_for_over_limit
        (now passes)                  (no-donor → delegate to
                                        master_cloner instead of skip)
```

## Architecture Decisions (locked)

1. **Repair instead of reject (Scenario A).** A masterless deck is repaired via a three-level cascade (`master_repairer._repair_if_needed`), never fatally rejected. `ppt_builder.generate_ppt_from_data` calls `_repair_if_needed` between `Presentation()` load and `_validate_template`.
2. **Three-level cascade, ordered by fidelity to the user's file:**
   - **Level 1 — salvage `ppt/theme/theme1.xml` from the zip.** The theme part often survives when the master part is stripped. Read at the zip level, parse via `_raw_theme_colors_and_fonts`, inject `default.pptx`'s master skeleton with the salvaged theme content.
   - **Level 2 — scavenge explicit styles from slide XML.** If the theme part is gone, aggregate per-run `<a:rPr>` (font/size/color) and per-shape `<a:spPr>` (fill) from surviving `ppt/slides/slideN.xml` into a best-effort theme.
   - **Level 3 — fallback to `default.pptx`'s theme.** Used only when neither theme nor slide styles survive. **This is the only path that ignores user styling** — and only because nothing of the user's file is left to preserve.
3. **Record the level used.** The repair level actually applied is written to the `<output>.render.json` sidecar (additive `templating`/repair field), so the result is auditable.
4. **Extend instead of skip (Scenario B/C).** When the master exists but a needed layout type is missing (or the master has no usable placeholders), clone the user's master and create new layouts under the clone — do **not** silently skip.
5. **Option 1: single cloned master + multiple new layouts (shared theme).** Chosen over Option 2 (multiple masters) because a shared theme part keeps the user's branding intact across all layouts. The user's theme is **never replaced** on the B/C path.
6. **Borrow placeholder geometry from `default.pptx`.** `default.pptx` covers all 8 slide types; its placeholder geometry is borrowed and scaled to the user's slide dimensions via `geometry.py` primitives (`normalize_polygon` → `denormalize_polygon`, with `aspect_ratios_match` as a no-op gate when ratios already match — mirroring US-4.6's resize prep).
7. **No-donor delegation.** `layout_creator.clone_for_over_limit` (line 201-203) currently skips when `donor_idx is None`; it will instead delegate to `master_cloner`. `state_machine` dispatches Level 0 → Level 1 (extension).
8. **Reload-verify + rollback.** The new `master_cloner` mirrors `clone_for_over_limit`'s safety model: save → reload-verify → rollback (delete the derived file) on any failure, so the base file stays authoritative.
9. **Schema tolerance.** `schema_extractor` tolerates a missing master (emits an empty `slide_master` schema entry instead of raising), and `state_machine.resolve_and_clone` re-embeds the schema after a Level 1 extension so the derived template is self-describing (consistent with US-4.3).
10. **Phase 0 first (bugfix prerequisite).** The two bare `slide_masters[0]` accesses must raise `TemplateError` (not raw `IndexError`) before the Scenario-A repair layers on top — otherwise a masterless input crashes opaquely at the layout layer.
11. **`_common` is the home for `master_repairer.py`.** It is shared infra (consumed by `ppt_builder` and peer to `schema_extractor`/`geometry`), so it lives in `.opencode/skills/_common/scripts/` per US-5.2's shared-`common/` design. `master_cloner.py` lives in `template-modifier-skill/scripts/` (it is Capability-B-specific, parallel to `layout_creator`/`state_machine`).

## Deliverables

**New** `.opencode/skills/_common/scripts/master_repairer.py`:
- `_salvage_theme_part(zip_path)` (Level 1) — read `ppt/theme/theme1.xml` at zip level; parse via `_raw_theme_colors_and_fonts`.
- `_scavenge_slide_styles(slides_xml)` (Level 2) — aggregate `<a:rPr>`/`<a:spPr>` from `ppt/slides/slideN.xml` into a best-effort theme.
- Level 3 fallback — `default.pptx` theme verbatim.
- `_repair_if_needed(prs)` — inject `default.pptx` master skeleton, optionally replace theme content; return the level used.

**New** `.opencode/skills/template-modifier-skill/scripts/master_cloner.py`:
- Clone the user's master (shared theme part preserved).
- Create new layouts under the clone with placeholder geometry borrowed from `default.pptx`, scaled via `geometry.py`.
- Reload-verify + rollback (mirror `clone_for_over_limit`).

**Change** `.opencode/skills/template-modifier-skill/scripts/layout_creator.py`:
- Line 120 (`_clone_layout_into`) + line 171 (`_verify_layouts`) — guard bare `slide_masters[0]`; raise `TemplateError` when empty (Phase 0).
- Lines 201-203 (`clone_for_over_limit`) — when `donor_idx is None`, delegate to `master_cloner` instead of skipping (Phase 2).

**Change** `.opencode/skills/template-modifier-skill/scripts/state_machine.py`:
- Dispatch Level 0 → Level 1 (extension) when a needed slide type is missing; trigger schema re-embed after extension (Phase 2/3).

**Change** `.opencode/skills/generate-slide-skill/scripts/ppt_builder.py`:
- `generate_ppt_from_data` calls `_repair_if_needed` between `Presentation()` load and `_validate_template` (Phase 1).

**Change** `.opencode/skills/_common/scripts/schema_extractor.py`:
- Lines 878-882 — tolerate missing master (emit empty `slide_master` schema entry instead of `TemplateExtractionError`) (Phase 1/3).

**Docs**: `.opencode/agents/pptx-subagent.md` (User-Supplied Templates section), `.opencode/skills/template-modifier-skill/SKILL.md`, `AGENTS.md` (US-4.8 epic entry) — Phase 4.

**Tests**: `test_master_repairer.py`, `test_master_cloner.py`, extended `test_template_validation.py`/`test_layout_creator.py`, + fixtures (A-L1, A-L2, A-L3, B, C) — Phase 5.

## Acceptance Criteria (per phase + overall)

Overall (parent #78):
- [ ] AC1 — A masterless `.pptx` (Scenario A) is repaired (not rejected) and produces a valid `.pptx` whose embedded schema reports a non-empty `slide_master`.
- [ ] AC2 — The three-level cascade is tried in order (L1 → L2 → L3); the level used is recorded in the render/templating sidecar.
- [ ] AC3 — A deck missing a layout for a needed slide type (Scenario B/C) gets a cloned-master + borrowed-geometry layout; the user's original theme/fonts survive (no `default.pptx` theme bleed except on L3).
- [ ] AC4 — `layout_creator` raises `TemplateError` (not `IndexError`) on a masterless input, at both `_clone_layout_into` (line 120) and `_verify_layouts` (line 171).
- [ ] AC5 — `schema_extractor` tolerates a missing master (empty `slide_master` entry); `state_machine.resolve_and_clone` re-embeds the schema after a Level 1 extension.
- [ ] AC6 — All new modules have unit-test coverage (fixtures A-L1, A-L2, A-L3, B, C); the existing `pytest` suite stays green.
- [ ] AC7 — Documentation describes the repair/extension behavior and the customer constraint (preserve user master/styles).

Per-phase gates:
- Phase 0 (#79): `test_clone_layout_into_raises_on_masterless` + `test_verify_layouts_raises_on_masterless` pass; existing tests green.
- Phase 1 (#80): fixtures A-L1/A-L2/A-L3 each repaired at the correct level; sidecar records the level; `schema_extractor` no longer raises on missing master; 9 unit tests pass.
- Phase 2 (#81): fixtures B/C get a cloned-master + borrowed-geometry layout that renders; user theme survives; geometry round-trips via `normalize_polygon`/`denormalize_polygon`; 9 unit tests pass.
- Phase 3 (#82): repaired/extended decks' embedded schemas are non-stale and resolve via `get_render_contract`; 112 existing extractor tests stay green.
- Phase 4 (#83): three doc files updated; no stale "masterless files are rejected" claims remain.
- Phase 5 (#84): 5 fixtures loadable; `test_master_repairer.py` + `test_master_cloner.py` pass; full `pytest` suite green.

## Implementation Phases

> Dependency graph: `Phase 0 → Phase 1 ∥ Phase 2 → Phase 3 → Phase 4 → Phase 5`

### Phase 0: Harden existing crash points (bugfix) — #79
- [ ] T0.1: `layout_creator.py:120` (`_clone_layout_into`) — replace bare `prs.slide_masters[0]` with a guard raising `TemplateError` (message naming the missing master) when `prs.slide_masters` is empty.
- [ ] T0.2: `layout_creator.py:171` (`_verify_layouts`) — same guard for `reloaded.slide_masters[0].slide_layouts`.
- [ ] T0.3: Ensure `TemplateError` is imported/defined consistently with the existing error hierarchy (`_common` / `template-modifier` error types).
- [ ] T0.4: Tests — `test_clone_layout_into_raises_on_masterless`, `test_verify_layouts_raises_on_masterless`; confirm existing `template-modifier-skill` tests stay green.

### Phase 1: Master repair cascade (Scenario A) — #80
- [ ] T1.1: Create `_common/scripts/master_repairer.py` — `_salvage_theme_part(zip_path)` (Level 1): zip-level read of `ppt/theme/theme1.xml`, parse via `_raw_theme_colors_and_fonts`.
- [ ] T1.2: `_scavenge_slide_styles(slides_xml)` (Level 2): aggregate `<a:rPr>` (font/size/color) + `<a:spPr>` (fill) from `ppt/slides/slideN.xml` into a best-effort theme.
- [ ] T1.3: Level 3 fallback — `default.pptx` theme verbatim.
- [ ] T1.4: `_repair_if_needed(prs)` — inject `default.pptx` master skeleton, optionally replace theme content with salvaged/scavenged data; return the level used.
- [ ] T1.5: `ppt_builder.generate_ppt_from_data` — call `_repair_if_needed` between `Presentation()` load and `_validate_template`.
- [ ] T1.6: `schema_extractor.py:878-882` — tolerate missing master (emit empty `slide_master` schema entry).
- [ ] T1.7: Record repair level in `<output>.render.json` sidecar (additive `templating`/repair field).
- [ ] T1.8: Tests — 9 unit tests covering L1/L2/L3 + edge cases (fixtures A-L1/A-L2/A-L3).

### Phase 2: Master cloning + default geometry borrowing (Scenario B/C) — #81
- [ ] T2.1: Create `template-modifier-skill/scripts/master_cloner.py` — clone the user's master (shared theme part preserved; never replace the user's theme).
- [ ] T2.2: Create new layouts under the clone with placeholder geometry borrowed from `default.pptx` (all 8 slide types), scaled via `geometry.py` (`normalize_polygon`/`denormalize_polygon`/`aspect_ratios_match`).
- [ ] T2.3: Reuse ID-allocation (`layout_creator._max_layout_id`) and the 7-step clone analog (`layout_creator._clone_layout_into`).
- [ ] T2.4: Reload-verify + rollback, mirroring `layout_creator.clone_for_over_limit`.
- [ ] T2.5: `layout_creator.py:201-203` — when `donor_idx is None`, delegate to `master_cloner` instead of skipping.
- [ ] T2.6: `state_machine.py` — dispatch Level 0 → Level 1 when a needed slide type is missing.
- [ ] T2.7: Tests — 9 unit tests (theme preservation, geometry scaling round-trip, fingerprint matching; fixtures B/C).

### Phase 3: Schema tolerance + contract refresh — #82
- [ ] T3.1: Confirm `schema_extractor` missing-master tolerance (Phase 1 change) is consistent with `layout_contract` / `contract_adapter` / `template_introspector`.
- [ ] T3.2: Verify `state_machine.resolve_and_clone` triggers schema re-embed after a Level 1 extension.
- [ ] T3.3: Verify `_warn_if_embedded_stale` detects layout-count changes (no false positives after extension).
- [ ] T3.4: Confirm `get_render_contract` (→ `contract_adapter`) resolves to the freshly-embedded JSON for repaired/extended decks; 112 existing extractor tests stay green.

### Phase 4: Agent prompt + documentation — #83
- [ ] T4.1: `.opencode/agents/pptx-subagent.md` "User-Supplied Templates" section — describe scenarios A/B/C, repair-vs-reject philosophy, three-level cascade, master-cloning extension, customer constraint.
- [ ] T4.2: `.opencode/skills/template-modifier-skill/SKILL.md` — document the master-cloning extension path (Option 1) and the no-donor delegation.
- [ ] T4.3: `AGENTS.md` — add US-4.8 epic entry (mirror existing US-4.x entry style).
- [ ] T4.4: Consistency sweep — no stale "masterless files are rejected" claims remain in the three files.

### Phase 5: Test suite + fixtures — #84
- [ ] T5.1: Fixtures — A-L1 (no master, theme present), A-L2 (no master, no theme, slide styles present), A-L3 (no master, no theme, no styles), B (master, missing layouts), C (minimal master, no placeholders).
- [ ] T5.2: `test_master_repairer.py` — 9 tests (L1/L2/L3 + edge cases).
- [ ] T5.3: `test_master_cloner.py` — 9 tests (theme preservation, geometry scaling, fingerprint matching).
- [ ] T5.4: Extend `test_template_validation.py`/`test_layout_creator.py` — Phase 0 masterless-guard tests.
- [ ] T5.5: Run `python -m pytest tests/ -q` from `.opencode/skills/generate-slide-skill/scripts`; confirm full suite green (no regression).

## Risks

- **Level-1/2 theme reconstruction fidelity** — salvaged/scavenged themes are best-effort; a reconstructed theme may not be pixel-identical to the original. Mitigated by: Level 1 (salvage the actual `theme1.xml`) is exact when the part survives; Level 2 aggregates explicit run/shape styles only; Level 3 is the explicit last resort. The level used is recorded in the sidecar for auditability (Decision 3).
- **Master-clone theme bleed (Scenario B/C)** — if the clone accidentally references `default.pptx`'s theme instead of the user's, branding is lost. Mitigated by Decision 5 (shared theme part from the user's master; never replace the user's theme on the B/C path) + Phase 2's theme-preservation test.
- **Geometry scaling mismatch** — borrowed `default.pptx` placeholder geometry scaled to a different slide size could misplace placeholders. Mitigated by reusing US-4.6's proven `geometry.py` primitives (`normalize_polygon`/`denormalize_polygon`/`aspect_ratios_match`) + Phase 2's geometry round-trip test.
- **Reload-verify false negatives** — a cloned layout that is present but not findable by name could trigger a spurious rollback. Mitigated by mirroring `clone_for_over_limit`'s proven reload-verify + rollback model (Decision 8).
- **Schema tolerance breaking downstream consumers** — emitting an empty `slide_master` entry could surprise `layout_contract`/`contract_adapter`. Mitigated by Phase 3 (explicit verification across the contract layer) + the 112-test extractor regression gate.
- **Scope creep into freeform rebuild** — the freeform coordinate-placement path (US-4.6 `coordinate_placer.py`) is tempting but out of scope; this plan keeps the `add_slide` inheritance path (styling/bullets stay on-brand) and does not rebuild shapes from coordinates.

## Out of scope (explicit)

- **Freeform shape rebuild** — `coordinate_placer.py` / polygon reconstruction of shapes from normalized coordinates. This plan borrows placeholder *geometry* (position/size) from `default.pptx` and scales it; it does not reconstruct arbitrary freeform shapes.
- **Multi-master Option 2** — rejected in favor of Option 1 (single cloned master + shared theme, Decision 5).
- **Repair of non-PPTX files** (.ppt legacy, .key, .odp) — out of scope; only `.pptx` (OOXML) is handled.
- **New slide-type fingerprints** — `master_cloner` borrows layouts for the existing 8 slide types; no new fingerprint taxonomy is introduced.

## Testing strategy

- **Unit tests** (Phase 5): `test_master_repairer.py` (9 — one per cascade level × fixture + edge cases), `test_master_cloner.py` (9 — theme preservation, geometry round-trip, fingerprint matching), extended `test_template_validation.py`/`test_layout_creator.py` (Phase 0 guards).
- **Fixtures** (Phase 5): 5 minimal synthetic `.pptx`/zip artifacts — A-L1, A-L2, A-L3, B, C — each exercising one scenario/level. Built small and deterministic (no dependence on external decks).
- **Regression gate**: the existing `test_schema_extractor.py` suite (112 tests) and the `generate-slide-skill` `pytest` suite must stay green after every phase.
- **Manual verification** (post-Phase 5): run the primary-agent flow once on a masterless file and on a missing-layout file; confirm repair/extension + theme survival + sidecar audit field.

## References

- Requirements source: issue #78 — confirmed design decisions (11 locked) + three-path dispatch + customer constraint (preserve user master/styles).
- Sub-issues: #79 (Phase 0), #80 (Phase 1), #81 (Phase 2), #82 (Phase 3), #83 (Phase 4), #84 (Phase 5).
- Crash/degradation points (verified): `schema_extractor.py:877-882`; `layout_creator.py:120`/`:171`/`:201-203`.
- Reused primitives (no change): `schema_extractor.read_embedded_schema`/`embed_schema`/`_raw_theme_colors_and_fonts`; `layout_creator._clone_layout_into`/`_max_layout_id`/`clone_for_over_limit`; `geometry.normalize_polygon`/`denormalize_polygon`/`aspect_ratios_match`; `default.pptx`.
- Related epics: US-4.3 (auto-chain/templated output — self-describing embed), US-4.6 (multi-aspect-ratio — `geometry.py` primitives), US-5.2 (shared `_common/` home).
- PLAN format template: `PLANS/PLAN-GIT-76.md`.
