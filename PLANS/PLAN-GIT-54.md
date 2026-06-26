# PLAN-GIT-54 — US-1.4: Font Detection & Availability Checking

**Issue**: #54
**Branch**: GIT-54 (base: dev)
**Priority**: Must Have (P0)
**Status**: Implemented (v2 + code-review follow-ups; all phases complete, 80 tests green)

## Goal

Bring US-1.4 from **❌ Not met** to ✅ Met by adding per-textbox font detection to
`schema_extractor`. Every text-bearing component (`textbox`/`placeholder`) must
carry a populated `font` object — `family`, `size_pt`, `weight`, `color`,
`alignment`, `is_available`, `fallback` — plus a `runs[]` array, and the deck
must aggregate a deduped top-level `missing_fonts[]` (surfaced as a non-fatal
**`ValidationIssue` warning**, mirroring US-1.2/1.3, when non-empty). Today the
extractor emits only empty stubs (`font: {}`, `runs: []`, `missing_fonts: []`),
so there is **no per-textbox font detection** at all.

## Strategic Context

Builds on US-1.1 (issue #48, PR #49), US-1.2 (issue #50, PR #51), and US-1.3
(issue #52, PR #53). `schema_extractor.py` already extracts theme-level fonts
(`major_latin`/`minor_latin`) via `_raw_theme_colors_and_fonts()` and lays down
the `font`/`runs`/`missing_fonts` scaffolding — but leaves them empty
(`schema_extractor.py:415`, `:468-469`). This plan fills that scaffolding with
real per-run font metadata and surfaces non-built-in dependencies. The component
model is **metadata-only** (the renderer does not yet consume per-component
fonts), so this is about schema fidelity / contract completeness. GAP-ANALYSIS §2
grades this ❌ Not met (one of the P0 stories in §4).

## Font Inheritance — Explicit-Only (correctness rule)

PPTX font properties cascade: run → paragraph → body → layout → master → theme.
`python-pptx` exposes only the **explicit** value on a run/paragraph (returns
`None` for inherited values). Resolving the full cascade would require walking
the layout/master/theme inheritance chain — that is **out of scope** for US-1.4.

**Decision:** capture explicit-only. Inherited (`None`) values are emitted as
`null`. The already-captured `theme.font_palette` (`major_latin`/`minor_latin`)
provides the default context for any consumer that needs it. This means many
fields will be `null` in practice — this is **expected and by design**, not a
gap. Full-cascade resolution is deferred.

## Script Scope — Latin / English Only (architecture review MAJOR-1)

`python-pptx`'s `run.font.name` reads **only** the `<a:latin>` typeface; fonts
set via `<a:ea>` (East Asian — e.g. 微软雅黑/思源黑体) or `<a:cs>` (complex
script) return `None`. **This project targets English/Latin fonts only** (per
the maintainer's scope decision), so US-1.4 inspects `<a:latin>` exclusively and
**CJK / complex-script typeface detection is explicitly out of scope** (see
§"Out of Scope"). Consequence: a `null`/inherited latin family is treated as
available (`is_available:true`, no warning) — it falls back to the theme
default. (If CJK support is later needed, extend `_extract_text_fonts` to also
probe `<a:ea>`/`<a:cs>` on `run.font._rPr`, mirroring `research_theme.py`'s
theme-level ea read.)

## Built-in Allowlist & Fallback Mapping

- **`_BUILTIN_FONTS`** — a curated allowlist of ~25 MS-Office/universal fonts
  (Calibri, Calibri Light, Arial, Times New Roman, Cambria, Consolas, Courier
  New, Georgia, Verdana, Tahoma, Trebuchet MS, Comic Sans MS, Impact, Segoe UI,
  Candara, Corbel, Constantia, Century Gothic, …). Membership drives
  `is_available` and whether a family lands in `missing_fonts`. Subjective but
  curated and **extensible**.
- **`_FONT_FALLBACK_MAP`** — maps common non-built-in families to a built-in
  substitute (Helvetica→Arial, Roboto/Inter/Open Sans/Lato→Arial, Helvetica
  Neue/SF Pro→Arial, Garamond Premier Pro→Garamond, …). The **default** for an
  unmapped non-built-in family is the **theme body font** when it is itself
  built-in (`theme.font_palette.body`), else `"Arial"` (architecture review
  MINOR-3). A fallback is **always** a built-in name (satisfies AC4). `fallback`
  is `null` when the family is built-in or `null`.
- **`is_available`** — `family in _BUILTIN_FONTS`. A `null`/inherited family is
  treated as available (`true`) and emits **no** warning.

## Architecture Decisions (locked, v2)

1. **Font inheritance — explicit-only.** Capture only values directly set on the
   run/paragraph; inherited (`None`) values → `null`. `theme.font_palette`
   provides the default context. Full-cascade resolution is deferred.
2. **Component `font` summary — first run.** The component-level `font` object
   summarizes the **first text run** (explicit props) + the first paragraph's
   `alignment`; `runs[]` captures every run. An empty textbox (zero runs) emits
   a fully-keyed `font` with all-`null` fields and an empty `runs[]` (NIT-1).
3. **Built-in allowlist.** Curated ~25 fonts in `_BUILTIN_FONTS` (extensible).
4. **Fallback via `_FONT_FALLBACK_MAP` + theme-aware default (MINOR-3).** Mapped
   family → mapped built-in; unmapped non-built-in → `theme.font_palette.body`
   if built-in else `"Arial"`; `null`/built-in family → `null`. Always a built-in
   name (AC4).
5. **`is_available` semantics (Latin/English scope).** `family in _BUILTIN_FONTS`;
   a `null`/inherited (latin) family → `true`, no warning.
6. **`runs[]` shape — nested (MINOR-1).** Each run is
   `{text, font: {family, size_pt, weight, color}}` for runs that carry text
   (matches the Reference schema in `chenyu-user-stories.md`).
7. **`missing_fonts[]` + ValidationIssue warning (MAJOR-3).** Deduped non-built-in
   families, each `{family, is_available:false, fallback, download_url:null}`.
   `validate_template_schema` emits a **non-fatal `ValidationIssue(severity="warning")`**
   per missing font (mirrors US-1.2 degenerate-polygon / US-1.3 shape-low patterns);
   `extract_schema` also `logger.warning(...)`s. AC3 is provable via
   `result.warnings`. The subagent agent definition is **not** modified.
8. **Scope — text types only / C1 preserved.** Font detection applies to
   `_TEXT_TYPES = {"textbox","placeholder"}`; non-text components keep the C1
   rule (no `font`). Incremental edits, zero render-path changes.
9. **AC4 enforcement + invariant (MINOR-5).** A non-null `fallback` that is not a
   built-in name is a validator **ERROR** (blocks `is_valid`). Also enforce the
   cheap invariant `is_available == (fallback is None)` (warning).
10. **Color/size/weight/alignment conversions (MINOR-2/6/7).**
    - `run.font.color`: capture hex **only** when `color.type == MSO_COLOR_TYPE.RGB`
      (guard required — `.rgb` raises on theme/None colors); else `null`.
    - `run.font.size` (Length/EMU) → `size.pt` (number).
    - `run.font.bold` (tri-state `bool|None`) → `weight`: `True→"bold"`,
      `False/None→null` (italic is out of scope).
    - `paragraph.alignment` (`PP_ALIGN`) → `"left"/"center"/"right"/"justify"`/`null`.
11. **Theme-before-components ordering.** `_build_theme(prs)` has no component
    dependency, so build it **before** components and thread the resolved
    `default_body` into `_extract_text_fonts` (via `_extract_components`/
    `_build_component`) so component `font.fallback` is theme-aware.

## Deliverables (all incremental edits, zero new files)

- `scripts/schema_extractor.py`: `_BUILTIN_FONTS`; `_FONT_FALLBACK_MAP`;
  `_font_fallback(family, default_body)`; `_extract_text_fonts(shape, default_body)`
  (walk `text_frame.paragraphs`→`.runs`; guarded RGB color; size→pt;
  bold→weight; alignment→string); reorder `extract_schema` to build theme first;
  in `_build_component` (`:467-469`) replace the `font:{}`/`runs:[]` stubs with
  populated values (nested `runs[]`); aggregate `missing_fonts` after components;
  `_validate_component` adds font type checks + AC4 ERROR + the
  `is_available⟺fallback` invariant; `validate_template_schema` emits the
  per-missing-font WARNING.
- `scripts/schemas/template_schema.json`: tighten the `font` description
  ("populated by the extractor; sub-fields optional"); define the `missing_fonts`
  item shape (`family`/`is_available`/`fallback`/`download_url` with
  `download_url` nullable).
- `scripts/tests/test_schema_extractor.py`: integration test (deck with custom +
  built-in runs); unit tests for `_font_fallback`/`_BUILTIN_FONTS`; **update
  `test_text_components_carry_font_stub`** (MAJOR-2) to assert a *populated*
  `font` (rename → `test_text_components_carry_populated_font`); add the
  missing-font `result.warnings` assertion (AC3) and the AC4-error / invariant
  tests.
- `docs/user-stories/GAP-ANALYSIS.md`: US-1.4 → ✅ Met (Met 5→6).
- `docs/user-stories/chenyu-user-stories.md`: US-1.4 ACs → `[x]`.

## Acceptance Criteria (US-1.4) — to deliver

- [x] Every textbox component has a `font` object with all specified fields
      (`family`, `size_pt`, `weight`, `color`, `alignment`, `is_available`,
      `fallback`).
- [x] `missing_fonts` array is empty when all fonts are built-in.
- [x] When non-built-in fonts are found, a user-facing warning lists them
      (delivered as a `ValidationIssue(severity="warning")` in
      `validate_template_schema`).
- [x] `fallback` is always a built-in font name (enforced as a validator ERROR).

## Implementation Phases

### Phase 1: Font data model + helpers (schema_extractor.py)

- [x] Task 1: Add `_BUILTIN_FONTS` (~25 MS-Office/universal fonts).
- [x] Task 2: Add `_FONT_FALLBACK_MAP` (Helvetica→Arial, Roboto/Inter/Open
      Sans/Lato→Arial, …) and `_font_fallback(family, default_body)` — mapped
      built-in name, else `default_body` (theme body or `"Arial"`), or `null`
      when family is built-in/`null`.
- [x] Task 3: Add `_extract_text_fonts(shape, default_body)` — walk
      `text_frame.paragraphs`→`.runs`; read `run.font.name/size/bold/color` +
      `paragraph.alignment`; **color guarded by `color.type == MSO_COLOR_TYPE.RGB`**
      (else `null`); `size→.pt`; `bold→weight`; `alignment PP_ALIGN→string`; return
      first-run `font` summary + nested `runs[]` (`{text, font:{...}}`); empty
      textbox → all-null keyed `font` + `runs:[]`.

### Phase 2: Component wiring + missing_fonts aggregation

- [x] Task 4: Reorder `extract_schema` — build `_build_theme(prs)` **before**
      components; resolve `default_body = theme.font_palette.body if in
      _BUILTIN_FONTS else "Arial"`; thread into `_extract_components`/
      `_build_component`/`_extract_text_fonts`.
- [x] Task 5: In `_build_component` (`:467-469`), for `_TEXT_TYPES` replace the
      `font:{}`/`runs:[]` stubs with `_extract_text_fonts` output (nested runs).
- [x] Task 6: In `extract_schema`, aggregate `missing_fonts` after components
      (deduped non-built-in families, each
      `{family, is_available:false, fallback, download_url:null}`); set
      `template_metadata.missing_fonts`; `logger.warning(...)` when non-empty.
- [x] Task 7: Non-text components keep the C1 rule (no `font`).

### Phase 3: Schema + validator

- [x] Task 8: `template_schema.json` — tighten `font` description; define
      `missing_fonts` item shape with `download_url` nullable. Keep `missing_fonts`
      **optional** (not in `_METADATA_REQUIRED`).
- [x] Task 9: `_validate_component` — font-field type checks; AC4: non-null
      `fallback` not built-in → **ERROR**; invariant
      `is_available == (fallback is None)` → warning. In
      `validate_template_schema`: emit one non-fatal `ValidationIssue(severity="warning")`
      per `missing_fonts` entry (AC3, mirror US-1.2/1.3).

### Phase 4: Tests

- [x] Task 10: Unit tests for `_font_fallback`/`_BUILTIN_FONTS` (built-in→null;
      Helvetica→Arial; unmapped→default_body).
- [x] Task 11: Unit tests for `_extract_text_fonts` (inherited→null; RGB→hex;
      theme/None color→null; size→pt; bold→weight; alignment→string; empty
      textbox→all-null keyed font + `runs:[]`).
- [x] Task 12: Integration test — deck with custom-font run + built-in-font run;
      assert populated `font`/nested `runs[]`/`is_available`/`fallback`/
      `missing_fonts`/`result.warnings`.
- [x] Task 13: **MAJOR-2** — update `test_text_components_carry_font_stub` →
      `test_text_components_carry_populated_font` (assert populated `font`, not
      `== {}`). Non-text C1 regression (no `font`). AC4-error + invariant tests.

### Phase 5: Docs

- [x] Task 14: Update `GAP-ANALYSIS.md` US-1.4 → ✅ Met (Met 5→6);
      `chenyu-user-stories.md` US-1.4 ACs → `[x]`.

## Test matrix

| Case | Expected |
| --- | --- |
| text component with a Calibri run | `family=Calibri`, `is_available=true`, `fallback=null`; `missing_fonts=[]`; no warning |
| text component with a mapped custom-font run (e.g. "Roboto") | `is_available=false`, `fallback="Arial"`; in `missing_fonts`; `result.warnings` non-empty |
| text component with an unmapped custom-font run | `fallback=theme body if built-in else "Arial"`; in `missing_fonts`; warning |
| textbox with multiple runs | `font` = first-run summary; nested `runs[]` has all runs |
| text with no explicit font (inherited) | inherited fields `null`; `is_available=true`; not in `missing_fonts` |
| empty textbox (zero runs) | `font` fully-keyed all-`null`; `runs:[]` |
| non-text component | no `font` (C1 preserved) |
| non-null `fallback` not built-in | validator ERROR (`is_valid=False`) |

## Verification

```bash
# from .opencode/skills/ppt-template-filler/scripts
python -m pytest tests/test_schema_extractor.py -q
python -m pytest tests/ -q
python schema_extractor.py --input templates/template.pptx --output /tmp/s.json
```

## Out of Scope (deferred) / OPEN QUESTIONS

- **Full-cascade font resolution** (run→paragraph→body→layout→master→theme) —
  deferred. Explicit-only capture (decision 1) is the scoped deliverable;
  `theme.font_palette` provides default context.
- **CJK / complex-script typefaces** (`<a:ea>`/`<a:cs>`) — out of scope (Latin/
  English only). A `null` latin family is treated as available. Extend later by
  probing `run.font._rPr` for ea/cs (precedent: `research_theme.py`).
- **Theme-color → hex resolution** — `run.font.color` is hex only when
  `type==RGB`; theme/None stay `null`.
- **`download_url` auto-discovery** — stays `null`.
- **Font detection on non-text components** — out of scope (C1 preserved).
- **Italic** — not captured (only `bold→weight`).
- **Validator↔schema divergence** (NIT-5) — each new hand-rolled check widens the
  manual-sync surface with `template_schema.json`; consolidating is US-5.2 scope.
- **OPEN** — should `missing_fonts` ever be promoted to `_METADATA_REQUIRED`? (v2
  answer: no — keep optional for backward-compatible externally-supplied schemas.)

## Risks

- **Explicit-only capture → many `null` fields** — expected by design
  (`theme.font_palette` gives defaults).
- **Color access variance** — `run.font.color.rgb` raises on theme/None; the
  `.type==RGB` guard (decision 10) is **mandatory**, not optional.
- **Theme-aware fallback sequencing** — requires building theme before components
  and threading `default_body` (decision 11); signature change on
  `_extract_components`/`_build_component`/`_extract_text_fonts`.
- **Test regression (MAJOR-2)** — `test_text_components_carry_font_stub` breaks on
  first run; Task 13 updates it.
- **"Built-in" is subjective** — curated, extensible allowlist.
- **Backward compatibility — low** — `font:{}` stub → populated object; output is
  generated (not committed), so no committed artifact changes shape.

## Code-review follow-ups

Post-implementation code review (verdict: Approve with revisions — 0 Critical, 1
Major). The applied fixes (behavior-preserving except the widened validator):

- [x] **M1 — invariant guard hole.** `_validate_component` dropped the
  `fb is not None` clause from the `is_available == (fallback is None)` check, so
  the `is_available=False`/`fallback=None` quadrant is now caught (warning). Added
  `test_invariant_false_available_null_fallback`.
- [x] **m1 — dead `families` output.** `_extract_text_fonts` now returns
  `(summary, runs)` (the `families` list was discarded by every caller); the
  `missing_fonts` aggregator keeps its single derivation from emitted
  `font`/`runs`.
- [x] **m4 — symmetric font type checks.** `_validate_component` now type-checks
  `family`/`weight`/`color`/`alignment` (string-or-null), matching the existing
  `is_available`/`size_pt` checks. Added `test_font_string_fields_type_checked`.
- [x] **m6/n2 — doc accuracy.** `_extract_text_fonts` docstring now says "first
  **explicit** paragraph alignment" (matches code); the stale
  `# populated in US-1.4` comment is now "populated below in extract_schema".
- Not applied (justified): **m3** (AC4 data-invariant assert), **m5** (decouple
  bundled-template test) — deferred.

Tests: `test_schema_extractor.py` → **80 passed** (was 78). No extraction-output
change; only the validator catches more (M1/m4) and dead code was removed (m1).

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 1, US-1.4.
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 US-1.4 (✅ Met), §4 P0.
- GitHub issue: #54 (`[US-1.4] Font Detection & Availability Checking`).
- Predecessors: US-1.1 (issue #48, PR #49), US-1.2 (issue #50, PR #51), US-1.3
  (issue #52, PR #53).
- Format template: `PLANS/PLAN-GIT-52.md`.
- Architecture review: findings MAJOR-1 (CJK scope-out), MAJOR-2 (test update),
  MAJOR-3 (ValidationIssue warning), MINOR-1/3/5/6/7, NIT-1/3/4/5 incorporated
  above; MINOR-2/4 (aggregate helper) not adopted.
- Code review: M1, m1, m4, m6/n2 applied — see §"Code-review follow-ups"; m3/m5
  not applied (justified).
- After implementation: update `GAP-ANALYSIS.md` US-1.4 → ✅ Met (Met 5→6);
  `chenyu-user-stories.md` US-1.4 ACs → `[x]`.
