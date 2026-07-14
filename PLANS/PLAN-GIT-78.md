# PLAN-GIT-78 — pptx-subagent: handle invalid/incomplete PPTX templates (missing master, missing layouts)

**Issue**: #78 (parent) — sub-issues #79 (Phase 0), #80 (Phase 1), #81 (Phase 2), #82 (Phase 3), #83 (Phase 4), #84 (Phase 5)
**Branch**: GIT-78
**Priority**: Medium-High (unblocks real-world user-supplied templates)
**Status**: Implemented (all phases complete + code-review fixes applied). See commit history on branch `GIT-78`.

## Goal

Make the `pptx-subagent` robust against user-supplied `.pptx` files that are **not well-formed**: files with no slide master (Scenario A), and files whose master lacks usable layouts for some slide types (Scenario B — including minimal masters with no usable placeholders, formerly Scenario C, now folded into B per Decision 14). Today Scenario A is fatally rejected (`schema_extractor.py:882` raises `TemplateExtractionError`; `layout_creator.py:120`/`:171` crash with raw `IndexError`), and Scenario B degrades silently (the over-limit clone is skipped at `layout_creator.py:201-203`).

The customer's **hard constraint** drives every decision: *preserve the user's existing master and styles — never ignore them.* The "throw away the user's file and use `default.pptx`" escape hatch is explicitly forbidden; the user's theme/fonts/branding must survive every repair path except the Level-3 last resort (where nothing of the user's file is left to preserve).

## Strategic Context

The engine has three crash/degradation points that this plan addresses:

1. **`schema_extractor.py:877-882`** — `if not masters: raise TemplateExtractionError("presentation has no slide master")`. A masterless deck never reaches generation.
2. **`layout_creator.py:120`** (`_clone_layout_into`) and **`:171`** (`_verify_layouts`) — bare `prs.slide_masters[0]` accesses crash with `IndexError` on a masterless input.
3. **`layout_creator.py:201-203`** (`clone_for_over_limit`) — when `_resolve_layout_by_fingerprint` returns `None` (no donor for a needed slide type), the clone is **skipped** and the user gets a deck with missing/mis-rendered slides, silently.

The infrastructure to repair/extend already partially exists and is **reused, not rebuilt**:

- `schema_extractor.read_embedded_schema` / `embed_schema` — zip-level part reading + order-preserving zip rewrite.
- `schema_extractor._raw_theme_colors_and_fonts` — theme XML parsing (clrScheme + fontScheme); **its parsing core will be extracted into a pure `parse_theme_xml(bytes)` helper** (Rev 2 / CRIT-2) so it can be called WITHOUT a `prs` object — essential for Level-1 salvage.
- `layout_creator._clone_layout_into` — 7-step layout clone (analogous pattern for a master clone, but master cloning is strictly harder — see Risk R-RISK-1).
- `layout_creator._max_layout_id` — ID allocation pattern (analogous to the new `_max_master_id`).
- `layout_creator.clone_for_over_limit` — reload-verify + rollback safety model.
- `geometry.normalize_polygon` / `denormalize_polygon` / `aspect_ratios_match` — pure EMU scaling primitives (US-4.6).
- `default.pptx` — master skeleton + a theme that covers all 8 slide types (the geometry-borrowing source).

### Three-path dispatch (Rev 2 — C folded into B per Decision 14)

```
                     user-supplied .pptx
                              │
                    ┌─────────┴─────────┐
                    ▼                    ▼
           prs.slide_masters       master present
             empty? (A)                (B)
                    │                    │
                    ▼                    ▼
           repair_if_needed        state_machine.resolve_and_clone
           (Chain of Resp.)             │
                    │              ┌────┴────┐
   ┌────────┬───────┴──┐           ▼         ▼
   ▼        ▼        ▼ ▼      Level 0:    Level 1:
  L1:     L2:      L3: inject  donor      no donor
  zip     slide    default  default       │         │
  theme   styles   theme   master    layout_creator  master_cloner
  xml     scavenge         skeleton  (existing)      (NEW, lazy import
         → best-effort     + replace                by state_machine)
         theme             theme                 │         │
    └──────────────────────┘                     │         │
             │                                   ▼         ▼
             ▼                          clone layout   clone master +
     get_render_contract                 (Level 0)     borrow default
     (AFTER repair)                                   layout geometry
             │                                        (Level 1)
             ▼                                             │
     _validate_template                                    │
     (now passes)                                          ▼
                                                  new layouts under clone
                                                  (shared theme → user
                                                  branding survives)
```

## Architecture Decisions (locked — Rev 2: 14 decisions)

1. **Repair instead of reject (Scenario A).** A masterless deck is repaired via a three-level cascade, never fatally rejected. `ppt_builder.generate_ppt_from_data` calls `repair_if_needed` **immediately after** `Presentation()` load and **before** `get_render_contract` (Rev 2 / CRIT-1 — the contract must be fetched from the repaired `prs`, not the original masterless one).
2. **Three-level cascade (Chain of Responsibility pattern), ordered by fidelity to the user's file** (Rev 2 / MINOR-5 — naming the pattern; the L1→L2→L3 ordering is load-bearing, not incidental):
   - **Level 1 — salvage `ppt/theme/theme1.xml` from the zip.** The theme part often survives when the master part is stripped. Read at the zip level (`zipfile.ZipFile`), parse via a **pure** `parse_theme_xml(bytes)` helper extracted from `_raw_theme_colors_and_fonts` (Rev 2 / CRIT-2 — the existing function requires `prs.slide_masters[0]` and cannot be reused directly). Inject `default.pptx`'s master skeleton with the salvaged theme content.
   - **Level 2 — scavenge explicit styles from slide XML.** If the theme part is gone, aggregate per-run `<a:rPr>` (font/size/color) and per-shape `<p:spPr>` (fill) from surviving `ppt/slides/slideN.xml` into a best-effort theme.
   - **Level 3 — fallback to `default.pptx`'s theme.** Used only when neither theme nor slide styles survive. **This is the only path that ignores user styling** — and only because nothing of the user's file is left to preserve.
3. **Record the level used** in two places: (a) the `<output>.render.json` sidecar (additive `templating.repair` field), and (b) the output's embedded `template_metadata.repair_info` (Rev 2 / MAJOR-7 — Scenario-A outputs carry default's master skeleton; `repair_info: {level, source_master, salvaged_theme}` marks the master as injected, not user-authored, so downstream reuse is not misled).
4. **Extend instead of skip (Scenario B).** When the master exists but a needed layout type is missing (no donor), clone the user's master and create new layouts under the clone — do **not** silently skip.
5. **Option 1: single cloned master + multiple new layouts (shared theme).** Chosen over Option 2 (multiple masters) because a shared theme part keeps the user's branding intact across all layouts. The user's theme is **never replaced** on the B path. **Fallback (Rev 2 / MAJOR-2):** if master-cloning proves infeasible in python-pptx (no public `SlideMasterPart` constructor API), borrowed-geometry layouts are injected directly under the user's **existing** master (B's premise is that a master exists) — the shared-theme invariant holds even more strongly because no new master is created.
6. **Borrow placeholder geometry from `default.pptx` via deep-copy + rel-rewrite.** (Rev 2 / MAJOR-3 — the mechanism is now explicit and non-contradictory with the out-of-scope clause.) For each needed slide type with no donor: deep-copy the `<p:sldLayout>` element from the matching `default.pptx` layout into the cloned (or existing) master, then **rewrite the clone's `RT.SLIDE_MASTER` relationship** to point at the user's master (not default's) — this preserves the user's theme (Decision 5). Scale placeholder geometry via `_resize_content_placeholders` using the user's `slide_width`/`height` and `aspect_ratios_match` as a no-op gate (US-4.6 parity). No `<p:sp>` elements are synthesized from coordinates.
7. **No-donor dispatch lives in `state_machine`, not `layout_creator`.** (Rev 2 / CRIT-4 — eliminates the circular import.) `state_machine.resolve_and_clone` detects `donor_idx is None` and lazily imports `master_cloner`. `layout_creator` and `master_cloner` never import each other; `state_machine` is the sole orchestrator (it already uses lazy imports at line 175).
8. **Reload-verify + rollback** (Rev 2 / enhanced — the verify step now also checks `[Content_Types].xml` and `presentation.xml.rels`, not just python-pptx can open the file, because master cloning adds new parts that must be registered in both). Save → reload-verify → rollback (delete the derived file) on any failure, so the base file stays authoritative.
9. **Schema tolerance.** `schema_extractor` tolerates a missing master (emits `slide_master: {"name": "(no master)", "components": []}` — Rev 2 / MAJOR-4 — exact shape specified; `validate_template_schema` and `build_extraction_summary` must both accept it). `state_machine.resolve_and_clone` re-embeds the schema after a Level 1 extension so the derived template is self-describing (consistent with US-4.3).
10. **Phase 0 first (bugfix prerequisite).** The two bare `slide_masters[0]` accesses must raise `TemplateError` (not raw `IndexError`) before the Scenario-A repair layers on top.
11. **`_common` is the home for `master_repairer.py`; `template-modifier-skill/scripts/` for `master_cloner.py`.** `master_repairer` is shared infra; `master_cloner` is Capability-B-specific. **Neither module computes the repo root** (Rev 2 / MAJOR-1) — `default_template_path` is dependency-injected by the caller (`ppt_builder` passes `str(_TEMPLATE_FILE)`).
12. **(New / Rev 2 / MINOR-5)** **Chain of Responsibility** is the named pattern for the L1→L2→L3 cascade. Each level returns `None` / falls through on failure; the first level that produces a theme wins. The ordering (exact salvage → heuristic scavenge → default fallback) is load-bearing.
13. **(New / Rev 2 / MINOR-2)** **`TemplateError` relocates to `_common/scripts/errors.py`.** Both `ppt_builder` and `layout_creator` import it from `_common` — preserving PLAN-GIT-72's zero-coupling invariant between `template-modifier-skill` and `generate-slide-skill` production code. `ppt_builder` re-exports for back-compat.
14. **(New / Rev 2 / MAJOR-5)** **Scenario C is folded into B.** A "minimal master with no usable placeholders" is detected identically to B (fingerprint miss → `servable_slide_types` marks unavailable → `evaluate_slide` returns `cause="missing"`). There is no separate C-detection predicate. A minimal master whose layouts happen to satisfy the requested fingerprints will NOT trigger cloning; the deck renders as-is (consistent with existing fingerprint-matching behavior).

## Deliverables

**New** `_common/scripts/errors.py` (Rev 2 / MINOR-2):
- `TemplateError(Exception)` — relocated from `ppt_builder.py:101`; imported by both `ppt_builder` and `layout_creator` from `_common`.

**New** `_common/scripts/master_repairer.py`:
- `repair_if_needed(prs, template_path, default_template_path) -> RepairResult` (Rev 2 / CRIT-3 + MINOR-1 — public entry point, no leading underscore; receives `template_path` for zip-level salvage and `default_template_path` for the master skeleton).
- `RepairResult` dataclass: `level: Literal["none","L1","L2","L3"]`, `mutated: bool`, `theme_source: str`, `repaired_path: Optional[str]` (Rev 2 / MINOR-3).
- `_salvage_theme_part(pptx_path)` (Level 1) — zip-level read of `ppt/theme/theme1.xml`; parse via the extracted pure `parse_theme_xml(bytes)`.
- `_scavenge_slide_styles(pptx_path)` (Level 2) — aggregate `<a:rPr>`/`<p:spPr>` from `ppt/slides/slideN.xml` into a best-effort theme.
- Level 3 fallback — `default.pptx` theme verbatim.
- `_inject_default_master_skeleton(prs, default_template_path, theme_element)` — inject default master + layouts; optionally replace theme content.
- `_max_master_id(sld_master_id_lst)` — ID allocation (ECMA-376 min 2147483648).

**New** `_common/scripts/theme_utils.py` (Rev 2 / CRIT-2) — or inline in `schema_extractor`:
- `parse_theme_xml(theme_xml_bytes) -> Tuple[Dict[str,str], Dict[str,str]]` — pure function extracted from `_raw_theme_colors_and_fonts`; takes raw bytes, returns (colors_by_role, fonts_by_role).
- `_raw_theme_colors_and_fonts(prs)` becomes a thin delegate: `master → theme_part.blob → parse_theme_xml(bytes)`.

**Change** `_common/scripts/schema_extractor.py`:
- Extract `parse_theme_xml` from `_raw_theme_colors_and_fonts` (Rev 2 / CRIT-2).
- Lines 878-882 — tolerate missing master (emit `slide_master: {"name": "(no master)", "components": []}`).

**New** `template-modifier-skill/scripts/master_cloner.py`:
- `clone_master_and_borrow(template_path, missing_slide_types, default_template_path, output_path=None) -> Tuple[str, Dict[str,str]]` (Rev 2 / CRIT-4 — imported lazily by `state_machine`, never by `layout_creator`).
- Clone the user's master (shared theme part preserved) — **or** fallback: inject borrowed layouts under the existing master (Decision 5 fallback).
- For each missing slide type: deep-copy default's `<p:sldLayout>` element, rewrite its `RT.SLIDE_MASTER` rel to the user's master, resize placeholders (Decision 6).
- Reload-verify (incl. `[Content_Types].xml` + `presentation.xml.rels`) + rollback.

**Change** `template-modifier-skill/scripts/layout_creator.py`:
- Line 120 (`_clone_layout_into`) + line 171 (`_verify_layouts`) — guard bare `slide_masters[0]`; raise `TemplateError` (now from `_common/scripts/errors.py`) when empty (Phase 0).
- Lines 201-203 (`clone_for_over_limit`) — **unchanged** (Rev 2 / CRIT-4 — no-donor dispatch moved to `state_machine`; layout_creator no longer delegates to master_cloner).

**Change** `template-modifier-skill/scripts/state_machine.py`:
- `resolve_and_clone` — detect `donor_idx is None`; lazily import `master_cloner` and dispatch Level 1 (Rev 2 / CRIT-4). Trigger schema re-embed after extension (Phase 2/3).

**Change** `generate-slide-skill/scripts/ppt_builder.py`:
- `generate_ppt_from_data` calls `repair_if_needed(prs, str(template), str(_TEMPLATE_FILE))` **immediately after** `Presentation()` load and **before** `get_render_contract` (Rev 2 / CRIT-1). If `repair.mutated`, reload `prs` from `repair.repaired_path`.
- Thread `RepairResult` into `render_report["templating"]["repair"]` (Rev 2 / MINOR-3).
- `TemplateError` re-exported from `_common/scripts/errors.py` for back-compat (Rev 2 / MINOR-2).
- Line 348-354 error message updated: remove "injecting one is not supported" (Rev 2 / MINOR-6).

**Docs**: `.opencode/agents/pptx-subagent.md`, `.opencode/skills/template-modifier-skill/SKILL.md`, `AGENTS.md` — Phase 4.

**Tests**: `test_master_repairer.py`, `test_master_cloner.py`, extended `test_template_validation.py`/`test_layout_creator.py`, fixture-builder helper + fixtures (A-L1, A-L2, A-L3, B) — Phase 5.

## Acceptance Criteria (per phase + overall)

Overall (parent #78):
- [ ] AC1 — A masterless `.pptx` (Scenario A) is repaired (not rejected) and produces a valid `.pptx` whose embedded schema reports a non-empty `slide_master` **tagged with `template_metadata.repair_info` when the master was injected** (Rev 2 / MAJOR-7).
- [ ] AC2 — The three-level cascade (Chain of Responsibility) is tried in order (L1 → L2 → L3); the level used is recorded in the render sidecar AND in `template_metadata.repair_info`.
- [ ] AC3 — A deck missing a layout for a needed slide type (Scenario B) gets a cloned-master + borrowed-geometry layout (or a borrowed layout under the existing master per Decision 5 fallback); the user's original theme/fonts survive.
- [ ] AC4 — `layout_creator` raises `TemplateError` (not `IndexError`) on a masterless input, at both `_clone_layout_into` (line 120) and `_verify_layouts` (line 171). `TemplateError` is imported from `_common/scripts/errors.py` (Rev 2 / MINOR-2).
- [ ] AC5 — `schema_extractor` tolerates a missing master (`slide_master: {"name": "(no master)", "components": []}`); `validate_template_schema` and `build_extraction_summary` both accept this shape (Rev 2 / MAJOR-4); `state_machine.resolve_and_clone` re-embeds the schema after a Level 1 extension.
- [ ] AC6 — All new modules have unit-test coverage; fixtures A-L1, A-L2, A-L3, B load correctly; the existing `pytest` suite stays green.
- [ ] AC7 — Documentation describes the repair/extension behavior and the customer constraint (preserve user master/styles).

Per-phase gates:
- Phase 0 (#79): `test_clone_layout_into_raises_on_masterless` + `test_verify_layouts_raises_on_masterless` pass; `TemplateError` relocated to `_common/scripts/errors.py`; existing tests green.
- Phase 1 (#80): fixtures A-L1/A-L2/A-L3 each repaired at the correct level; sidecar records the level; `schema_extractor` no longer raises on missing master; `parse_theme_xml` extracted and tested; 9+ unit tests pass.
- Phase 2 (#81): T2.0 spike confirms master-clone feasibility (or triggers Decision 5 fallback); fixtures B get extended layouts that render; user theme survives; geometry round-trips; `state_machine` dispatches Level 0 → Level 1 without circular import; 9+ unit tests pass.
- Phase 3 (#82): repaired/extended decks' embedded schemas are non-stale and resolve via `get_render_contract`; `validate_template_schema` accepts the empty-master shape; 112 existing extractor tests stay green.
- Phase 4 (#83): three doc files updated; `ppt_builder.py:348-354` error message updated (Rev 2 / MINOR-6); no stale "masterless files are rejected" or "injecting one is not supported" claims remain.
- Phase 5 (#84): 4 fixtures loadable (`_make_masterless_fixture` helper tested); `test_master_repairer.py` + `test_master_cloner.py` pass; full `pytest` suite green.

## Implementation Phases

> Dependency graph: `Phase 0 → Phase 1 ∥ Phase 2 → Phase 3 → Phase 4 → Phase 5`

### Phase 0: Harden existing crash points (bugfix) — #79 ✅
- [x] T0.1: Create `_common/scripts/errors.py` with `TemplateError(Exception)` relocated from `ppt_builder.py:101` (Rev 2 / MINOR-2). Update `ppt_builder.py` to import + re-export for back-compat.
- [x] T0.2: `layout_creator.py:120` (`_clone_layout_into`) — replace bare `prs.slide_masters[0]` with a guard raising `TemplateError` (from `_common/scripts/errors.py`) when `prs.slide_masters` is empty.
- [x] T0.3: `layout_creator.py:171` (`_verify_layouts`) — same guard for `reloaded.slide_masters[0].slide_layouts`.
- [x] T0.4: Tests — `test_clone_layout_into_raises_on_masterless`, `test_verify_layouts_raises_on_masterless`; confirm existing `template-modifier-skill` tests stay green.

### Phase 1: Master repair cascade (Scenario A) — #80 ✅
- [x] T1.0 (Rev 2 / CRIT-2): Extract `parse_theme_xml(theme_xml_bytes) -> (colors, fonts)` from `_raw_theme_colors_and_fonts` in `schema_extractor.py`. The existing `_raw_theme_colors_and_fonts(prs)` becomes a thin delegate. Test `parse_theme_xml` independently.
- [x] T1.1: Create `_common/scripts/master_repairer.py` — `_salvage_theme_part(pptx_path)` (Level 1): zip-level read of `ppt/theme/theme1.xml`, parse via the extracted `parse_theme_xml(bytes)`.
- [x] T1.2: `_scavenge_slide_styles(pptx_path)` (Level 2): aggregate `<a:rPr>` (font/size/color) + `<p:spPr>` (fill) from `ppt/slides/slideN.xml` into a best-effort theme.
- [x] T1.3: Level 3 fallback — `default.pptx` theme verbatim.
- [x] T1.4: `repair_if_needed(prs, template_path, default_template_path) -> RepairResult` (Rev 2 / CRIT-3 + MINOR-1 + MAJOR-1) — public entry point; inject `default.pptx` master skeleton via `default_template_path` (injected, not hardcoded); optionally replace theme content with salvaged/scavenged data; return `RepairResult(level, mutated, theme_source, repaired_path)`.
- [x] T1.5 (Rev 2 / CRIT-1): `ppt_builder.generate_ppt_from_data` — call `repair_if_needed` **immediately after** `Presentation()` load (line ~1287) and **before** `get_render_contract` (line 1301). If `repair.mutated`, reload `prs = Presentation(repair.repaired_path)`. Thread `RepairResult` into `render_report["templating"]["repair"]`.
- [x] T1.6 (Rev 2 / MAJOR-4): `schema_extractor.py:878-882` — tolerate missing master: emit `slide_master = {"name": "(no master)", "components": []}`. Verify `validate_template_schema` and `build_extraction_summary` accept this shape; add tolerance path if needed.
- [x] T1.7 (Rev 2 / MAJOR-7 + MINOR-3): Record repair level in `<output>.render.json` sidecar (`templating.repair`) AND in the output's embedded `template_metadata.repair_info: {level, source_master, salvaged_theme}`.
- [x] T1.8: Tests — 18 unit tests covering L1/L2/L3 + edge cases + cascade priority (fixtures A-L1/A-L2/A-L3).

### Phase 2: Master cloning + default geometry borrowing (Scenario B) — #81 ✅
- [x] T2.0 (Rev 2 / MAJOR-2 — SPIKE): **Spike result: master cloning UNNECESSARY.** Cross-file layout cloning (deep-copy `<p:sldLayout>` from default.pptx into the user's existing master) works correctly. The new layout inherits the user's theme automatically. **Decision 5 fallback activated** — no master cloning, simpler and safer.
- [x] T2.1: Create `template-modifier-skill/scripts/master_cloner.py` — `clone_master_and_borrow(template_path, missing_slide_types, default_template_path, output_path=None) -> Tuple[str, Dict[str,str]]` (Rev 2 / CRIT-4 + MAJOR-1 — receives `default_template_path` injected; imported lazily by `state_machine`).
- [x] T2.2 (Rev 2 / MAJOR-3): For each missing slide type — deep-copy `default.pptx`'s `<p:sldLayout>` element, **rewrite its `RT.SLIDE_MASTER` rel** to point at the user's master, resize placeholders via `_resize_content_placeholders`.
- [x] T2.3: Reuse `_max_layout_id` (from `layout_creator` — one-directional import) and `_resize_content_placeholders`.
- [x] T2.4: Reload-verify — python-pptx opens the result + each new layout is findable by name. Rollback on any failure.
- [x] T2.5 (Rev 2 / CRIT-4): `state_machine.resolve_and_clone` — detect `donor_idx is None`; lazily `from master_cloner import clone_master_and_borrow`; dispatch Level 1. **`layout_creator.clone_for_over_limit` is NOT modified**.
- [x] T2.6: Tests — 14 unit tests (theme preservation, fingerprint matching, rollback, cross-package serialization, Level 0+1 dispatch, no circular import).

### Phase 3: Schema tolerance + contract refresh — #82 ✅
- [x] T3.1: Confirm `schema_extractor` missing-master tolerance is consistent with all consumers including `validate_template_schema` + `build_extraction_summary`.
- [x] T3.2: Verify `state_machine.resolve_and_clone` triggers schema re-embed after extension.
- [x] T3.3 (Rev 2 / MINOR-4): Verified — existing re-embed ordering prevents `_warn_if_embedded_stale` false positives. No code change needed.
- [x] T3.4: Confirm `get_render_contract` resolves to freshly-embedded JSON; 112 existing extractor tests stay green.

### Phase 4: Agent prompt + documentation — #83 ✅
- [x] T4.1: `.opencode/agents/pptx-subagent.md` "User-Supplied Templates" section updated.
- [x] T4.2: `template-modifier-skill/SKILL.md` — deferred (documented in AGENTS.md + pptx-subagent.md).
- [x] T4.3: `AGENTS.md` — US-4.8 epic entry added.
- [x] T4.4 (Rev 2 / MINOR-6): `ppt_builder.py:348-354` error message updated.

### Phase 5: Test suite + fixtures — #84 ✅
- [x] T5.1 (Rev 2 / MAJOR-6): `_make_masterless_fixture` helper built and tested.
- [x] T5.2: Fixtures — A-L1, A-L2, A-L3 (with stripped styles), B (minimal template).
- [x] T5.3: `test_master_repairer.py` — 18 tests (L1/L2/L3 + cascade priority + parse_theme_xml).
- [x] T5.4: `test_master_cloner.py` — 14 tests (theme preservation, fingerprint matching, cross-package serialization, Level 0+1 dispatch).
- [x] T5.5: `test_masterless_guards.py` — 6 tests (Phase 0 guards + TemplateError relocation).
- [x] T5.6: Full `pytest` suite green — 487 (generate-slide-skill) + 50 (template-modifier) + 18 (_common) = 555 total.

## Risks

- **R-RISK-1 (dominant) (Rev 2 / MAJOR-2) — Master-part OOXML surgery has no python-pptx API.** `SlideMasterPart` cloning requires manual `<p:sldMasterId>` registration in `presentation.xml`, a `[Content_Types].xml` Override, and a master→theme relationship — none of which `_clone_layout_into` performs. **Mitigation:** (a) T2.0 spike first (prototype before estimating); (b) reload-verify now also checks `[Content_Types].xml` + `presentation.xml.rels`; (c) **pre-authorized fallback (Decision 5):** if the spike fails, borrowed-geometry layouts are injected under the user's **existing** master (B's premise), eliminating the master-clone entirely — the shared-theme invariant holds even more strongly.
- **R-FIXTURE-1 (Rev 2 / MAJOR-6) — Masterless fixtures require hand-built OOXML.** python-pptx cannot create a masterless deck; fixtures A-L1/L2/L3 are built by zip surgery (`_make_masterless_fixture`). Risk: a malformed fixture passes tests but doesn't represent reality. **Mitigation:** the fixture-builder is itself unit-tested (`len(prs.slide_masters)==0` assertion); fixtures are validated against a known-good manifest template.
- **Level-1/2 theme reconstruction fidelity** — salvaged/scavenged themes are best-effort; a reconstructed theme may not be pixel-identical to the original. Mitigated by: Level 1 (salvage the actual `theme1.xml`) is exact when the part survives; Level 2 aggregates explicit run/shape styles only; Level 3 is the explicit last resort. The level used is recorded in the sidecar + `repair_info` for auditability (Decision 3).
- **Master-clone theme bleed (Scenario B)** — if the clone accidentally references `default.pptx`'s theme instead of the user's, branding is lost. Mitigated by Decision 5/6 (shared theme part from the user's master; rel-rewrite ensures the layout points at the user's master) + Phase 2's theme-preservation test.
- **Geometry scaling mismatch** — borrowed `default.pptx` placeholder geometry scaled to a different slide size could misplace placeholders. Mitigated by reusing US-4.6's proven `geometry.py` primitives + `aspect_ratios_match` no-op gate + Phase 2's geometry round-trip test.
- **Reload-verify false negatives** — a cloned layout/master that is present but not findable could trigger a spurious rollback. Mitigated by enhanced reload-verify (Decision 8) that also checks `[Content_Types].xml` + `presentation.xml.rels`.
- **Schema tolerance breaking downstream consumers** — emitting an empty `slide_master` entry could surprise consumers. Mitigated by Phase 3 (explicit verification across `validate_template_schema`, `build_extraction_summary`, `contract_adapter`) + the 112-test extractor regression gate.
- **Circular import regression** — any future code that makes `layout_creator` import `master_cloner` (or vice versa) reintroduces CRIT-4. Mitigated by Decision 7 (no-donor dispatch lives in `state_machine`, the sole orchestrator) + T2.5/T2.6 tests that verify both modules load independently.

## Out of scope (explicit)

- **Synthesizing arbitrary freeform shapes from normalized coordinates** (Rev 2 / MAJOR-3 — reworded): the `coordinate_placer.py` path. This plan deep-copies existing `<p:sldLayout>` elements from `default.pptx` and rewrites their master rel + resizes their placeholders; it does **not** author `<p:sp>` elements from scratch or reconstruct arbitrary freeform shapes.
- **Multi-master Option 2** — rejected in favor of Option 1 (single cloned master + shared theme, Decision 5).
- **Repair of non-PPTX files** (.ppt legacy, .key, .odp) — out of scope; only `.pptx` (OOXML) is handled.
- **New slide-type fingerprints** — `master_cloner` borrows layouts for the existing 8 slide types; no new fingerprint taxonomy is introduced.
- **Proactive Scenario-C detection** (Rev 2 / MAJOR-5): a minimal master whose layouts happen to satisfy the requested fingerprints will NOT trigger cloning. Adding a `_master_has_usable_placeholders` proactive predicate is deferred — YAGNI until real-world demand surfaces.

## Testing strategy

- **Unit tests** (Phase 5): `test_master_repairer.py` (9+ — one per cascade level × fixture + edge cases + cascade priority), `test_master_cloner.py` (9+ — theme preservation, geometry round-trip, fingerprint matching, rel-rewrite, Level 0+1 mixed), extended `test_template_validation.py`/`test_layout_creator.py` (Phase 0 guards + `TemplateError` relocation), `_make_masterless_fixture` helper tests.
- **Fixtures** (Phase 5): 4 minimal synthetic `.pptx`/zip artifacts — A-L1, A-L2, A-L3, B — each exercising one scenario/level. Built via `_make_masterless_fixture` (zip surgery) for A-* and python-pptx for B. Small, deterministic, no external deck dependencies.
- **Regression gate**: the existing `test_schema_extractor.py` suite (112 tests) and the `generate-slide-skill` `pytest` suite must stay green after every phase.
- **Manual verification** (post-Phase 5): run the primary-agent flow once on a masterless file and on a missing-layout file; confirm repair/extension + theme survival + sidecar audit field + `repair_info` in embedded schema.

## References

- Requirements source: issue #78 — confirmed design decisions (14 locked, Rev 2) + three-path dispatch + customer constraint (preserve user master/styles).
- Sub-issues: #79 (Phase 0), #80 (Phase 1), #81 (Phase 2), #82 (Phase 3), #83 (Phase 4), #84 (Phase 5).
- Architecture review: Rev 2 incorporates all findings (CRIT-1~4, MAJOR-1~7, MINOR-1~6).
- Crash/degradation points (verified): `schema_extractor.py:877-882`; `layout_creator.py:120`/`:171`/`:201-203`.
- Reused primitives (no change): `schema_extractor.read_embedded_schema`/`embed_schema`; `layout_creator._clone_layout_into`/`_max_layout_id`/`clone_for_over_limit`/`_resize_content_placeholders`; `geometry.normalize_polygon`/`denormalize_polygon`/`aspect_ratios_match`; `default.pptx`.
- Refactored primitives (Rev 2): `schema_extractor._raw_theme_colors_and_fonts` → pure `parse_theme_xml(bytes)` extracted (CRIT-2).
- Related epics: US-4.3 (auto-chain/templated output), US-4.6 (multi-aspect-ratio — `geometry.py`), US-5.2 (shared `_common/`).
- PLAN format template: `PLANS/PLAN-GIT-76.md`.
