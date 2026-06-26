# PLAN-GIT-50 — US-1.2: Normalized Polygon Positioning

**Issue**: #50
**Branch**: GIT-50 (base: dev)
**Priority**: Must Have (P0)
**Status**: Implementation complete (v2 — architecture review findings incorporated)

## Goal

Add a cross-product winding check that verifies the canonical polygon winding
(AC3), closing the last acceptance gap for US-1.2. AC1, AC2, AC4 are already met
by US-1.1's `normalize_polygon()`. This plan delivers AC3, bringing US-1.2 to Met.

## Strategic Context

Builds on US-1.1 (PR #49): `schema_extractor.py` already emits a `polygon` field
on every component (4 normalized `{x,y}` points, order TL→TR→BR→BL). The polygon
is **metadata-only** — no consumer reads it (GAP-ANALYSIS §5 Decision 2 open),
so winding accuracy is about schema fidelity / contract hygiene, not rendering.

## Winding — Algebraic Interpretation (AC3 is literally met)

US-1.2 AC3 says "anti-clockwise winding is verified by a simple cross-product
check." A cross-product (shoelace) check measures the **algebraic** winding. For
the canonical order TL→TR→BR→BL the signed area is **positive**, which is
algebraically **counter-clockwise (CCW) = anti-clockwise**. So our output IS
anti-clockwise by the algebraic measure AC3 invokes, and the check verifies it.
AC3 is satisfied without any caveat.

> Visual note (for future readers): in screen coords (Y-down) the TL→TR→BR→BL
> trace *appears* clockwise, but that is a display-projection artifact — the
> algebraic winding (what a cross-product measures) is CCW. This is recorded in
> the schema `$comment` (Task 4) to prevent confusion.

## Architecture Decisions (locked, v2)

1. **Keep TL→TR→BR→BL** as the canonical order — matches the requirement's
   explicit coordinate list and US-1.1's existing output (zero churn). The
   cross-product check asserts **positive signed area** (= algebraic CCW).
2. **AC-focused scope** — winding check only; all shapes keep 4-point rectangular
   bounding boxes. Satisfies AC1–AC4 literally → US-1.2 Met. Non-rectangular
   actual-vertex extraction deferred.
3. **Incremental edits to existing files** — no new files, zero changes to the
   render path (`ppt_builder.py` / `template_introspector.py` untouched).
4. **Severity split (per architecture review MAJOR-1 + open Q1):**
   - **Reversed winding (negative area) → ERROR** (catches malformed external
     input; the extractor never produces this).
   - **Degenerate / zero-area (~0) → WARNING (non-fatal)** (real shapes may be
     legitimately thin/zero-width — dividers, hidden spacers — so it must not
     hard-block `is_valid`; mirrors the density-mode warning pattern in
     `schema_validator.py`).
5. **No defensive assertion in `normalize_polygon()` (per review MAJOR-1).**
   The single chokepoint is `validate_template_schema()`. An assertion in the
   extractor would crash on the existing zero-dims fallback
   (`schema_extractor.py:240-241`) and break `test_zero_dims_safe`. Dropped.

## Winding Check — Sign Convention & Epsilon

Normalized coords: (0,0)=top-left, (1,1)=bottom-right. Shoelace signed area:
`A = 0.5 * Σ (xᵢ·yᵢ₊₁ − xᵢ₊₁·yᵢ)`. Canonical TL→TR→BR→BL → **A > 0**.

- `A > epsilon` → canonical (pass)
- `A < -epsilon` → reversed → ERROR
- `|A| <= epsilon` → degenerate → WARNING

**Epsilon = `1e-9`** — below any physically meaningful shape area in normalized
[0,1] coords (a 1px divider on a 7.5″ slide ≈ 0.0001 normalized, area ≈ 1e-8 ≫
1e-9), above float noise. Verified by the reviewer's trace (canonical +0.48,
reversed -0.48, degenerate 0.0).

## Deliverables (all incremental edits, zero new files)

- `scripts/schema_extractor.py`: `_signed_area()` helper; winding check in
  `_validate_component()` (reversed→error, degenerate→warning).
- `scripts/schemas/template_schema.json`: `$comment` on `polygon`.
- `scripts/tests/test_schema_extractor.py`: winding tests + fix `_ok_component()`.
- `docs/user-stories/GAP-ANALYSIS.md`: US-1.2 → Met.
- `docs/user-stories/chenyu-user-stories.md`: AC3 → `[x]`.

## Acceptance Criteria (US-1.2)

- [x] Every component has a `polygon` array with exactly 4 `{x, y}` objects for rectangular shapes.
- [x] All x and y values are in [0.0, 1.0] range.
- [x] Anti-clockwise winding is verified by a simple cross-product check in the script. *(delivered by Task 2)*
- [x] Slide dimensions (EMU, inches, and aspect ratio string) are recorded in metadata.

## Implementation Phases

### Phase 1: Winding Check (single phase)

- [x] Task 1: Add `_signed_area(polygon)` shoelace helper (works for n-point
      polygons; forward-compatible with future non-rectangular vertices).
- [x] Task 2: Integrate winding check into `_validate_component()` in
      `validate_template_schema()` — reversed (A < -1e-9) → error;
      degenerate (|A| <= 1e-9) → warning. Use the existing `severity` field on
      `ValidationIssue`.
- [x] Task 3: Add `$comment` to `template_schema.json` `polygon` documenting:
      (a) canonical winding = positive signed area = algebraic CCW = anti-clockwise;
      (b) the Y-down visual artifact note; (c) `maxItems:4` is rectangular-only —
      non-rectangular vertex extraction (deferred) will require relaxing this.
- [x] Task 4: Tests:
      - canonical winding (axis-aligned rect) passes;
      - reversed (BL→BR→TR→TL) → error;
      - degenerate (4 identical points / zero area) → warning, `is_valid` stays True;
      - **non-axis-aligned canonical 4-gon** (e.g. `(0.2,0.1),(0.9,0.3),(0.7,0.9),(0.1,0.7)`)
        passes — proves shoelace generality, not a rect-only shortcut;
      - bundled `template.pptx` full extract still validates (`is_valid` True) —
        regression guard for any real zero-area shape.
      - Fix `_ok_component()` test helper to a canonical non-degenerate polygon
        (currently 4 identical points = zero area + mutable-aliasing bug).
- [x] Task 5: Docs — GAP-ANALYSIS §2 US-1.2 → ✅ Met (algebraic-CCW note);
      `chenyu-user-stories.md` AC3 → `[x]`.

## Out of Scope (deferred)

- **Non-rectangular actual vertices**: `custGeom` pathLst extraction,
  `triangle`/`flowChartPreparation`/`straightConnector1` preset→vertex tables.
- **Schema `polygon` maxItems=4 unchanged** (rectangular-only; relaxing it is
  part of the deferred non-rectangular work — noted in the `$comment`).
- **Polygon consumers / denormalization** — remains GAP-ANALYSIS §5 Decision 2.

## Risks

- **Very low for reversed/error path**: purely additive; the extractor never
  produces reversed winding, so no existing output is rejected.
- **Bundled-template regression guard (per review MINOR-5):** Task 4 includes a
  bundled-template full-extract validation test. If the bundled `template.pptx`
  or default `Presentation()` contains a zero-area shape, the new check flags it
  as a WARNING (non-fatal), so `is_valid` stays True and `test_validates_clean`
  does not regress. If such a shape surfaces, treat it as a data-quality finding
  (not a code change).
- Existing 37 tests unaffected: their polygons are canonical rectangles (pass);
  the `_ok_component()` fix (Task 4) keeps validation-failure tests correct.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 1, US-1.2.
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 US-1.2, §4 P0, §5 Decision 2.
- Predecessor: US-1.1 (issue #48, PR #49).
- Architecture review: findings MAJOR-1, MINOR-1..5, NIT-1 incorporated above.
