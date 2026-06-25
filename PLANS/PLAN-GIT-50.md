# PLAN-GIT-50 — US-1.2: Normalized Polygon Positioning

**Issue**: #50
**Branch**: GIT-50 (base: dev)
**Priority**: Must Have (P0)
**Status**: Planning

## Goal

Add a cross-product winding check that verifies the canonical polygon winding
(AC3), closing the last acceptance gap for US-1.2. AC1, AC2, AC4 are already met
by US-1.1's `normalize_polygon()`. This plan delivers AC3 and reconciles a
self-contradiction in the requirement, bringing US-1.2 to Met.

## Strategic Context

Builds on US-1.1 (PR #49): `schema_extractor.py` already emits a `polygon` field
on every component (4 normalized `{x,y}` points, order TL→TR→BR→BL). The polygon
is **metadata-only** — no consumer reads it (GAP-ANALYSIS §5 Decision 2 open),
so winding accuracy is about schema fidelity / contract hygiene, not rendering.

## Requirement Contradiction (must be resolved)

US-1.2 says "anti-clockwise" but explicitly lists the order
"top-left → top-right → bottom-right → bottom-left" (TL→TR→BR→BL). A shoelace
check confirms TL→TR→BR→BL traces a **clockwise** loop in screen coords (Y-down,
positive signed area). These are mutually exclusive — the "anti-clockwise" label
is a requirement misnomer.

## Architecture Decisions (locked)

1. **Keep TL→TR→BR→BL** as the canonical order. It matches the requirement's
   explicit coordinate list and US-1.1's existing output (zero churn to current
   data). The cross-product check verifies this canonical winding (assert
   positive signed area). The "anti-clockwise" label is documented as a
   requirement misnomer in the schema `$comment`.
2. **AC-focused scope.** Add the cross-product winding check only; all shapes
   keep 4-point rectangular bounding boxes. This satisfies AC1–AC4 literally,
   making US-1.2 Met. Non-rectangular actual-vertex extraction is deferred (see
   Out of Scope).
3. **Incremental edits to existing files** — no new files, zero changes to the
   render path (`ppt_builder.py` / `template_introspector.py` untouched).

## Winding Check — Sign Convention

Normalized coords: (0,0)=top-left, (1,1)=bottom-right → Y increases downward
(screen coords). Shoelace signed area for TL→TR→BR→BL is **positive**. The check
asserts `signed_area > epsilon`:
- positive area → canonical winding (pass)
- negative area → reversed winding (BL→BR→TR→TL) → error
- ~0 area → degenerate/collinear polygon → error

## Deliverables (all incremental edits, zero new files)

- `scripts/schema_extractor.py`: `_signed_area()` helper; winding check in
  `validate_template_schema()`; defensive assertion in `normalize_polygon`.
- `scripts/schemas/template_schema.json`: `$comment` on `polygon`.
- `scripts/tests/test_schema_extractor.py`: winding tests.
- `docs/user-stories/GAP-ANALYSIS.md`: US-1.2 → Met.
- `docs/user-stories/chenyu-user-stories.md`: AC3 → `[x]`.

## Acceptance Criteria (US-1.2)

- [x] Every component has a `polygon` array with exactly 4 `{x, y}` objects for rectangular shapes.
- [x] All x and y values are in [0.0, 1.0] range.
- [ ] Anti-clockwise winding is verified by a simple cross-product check in the script.
- [x] Slide dimensions (EMU, inches, and aspect ratio string) are recorded in metadata.

## Implementation Phases

### Phase 1: Winding Check (single phase — small change)

- [ ] Task 1: Add `_signed_area(polygon)` shoelace helper (normalized coords).
- [ ] Task 2: Integrate winding check into `validate_template_schema()` — assert
      `signed_area > epsilon`; reversed (negative) and degenerate (~0) windings
      emit a ValidationIssue error.
- [ ] Task 3: Add a defensive self-check assertion in `normalize_polygon()` (the
      extractor's own output is canonical by construction; assert it).
- [ ] Task 4: Add `$comment` to `template_schema.json` on the `polygon` property
      documenting the canonical winding (TL→TR→BR→BL, clockwise in screen coords)
      and flagging the "anti-clockwise" requirement misnomer.
- [ ] Task 5: Tests — canonical winding passes; reversed (BL→BR→TR→TL) fails;
      degenerate/collinear (zero area) fails; bundled `template.pptx` full extract
      still validates (no regression). Existing 37 tests stay green.
- [ ] Task 6: Docs — GAP-ANALYSIS §2 US-1.2 → ✅ Met (with convention note);
      `chenyu-user-stories.md` AC3 → `[x]`.

## Out of Scope (deferred)

- **Non-rectangular actual vertices**: `custGeom` pathLst extraction,
  `triangle`/`flowChartPreparation`/`straightConnector1` preset→vertex tables.
  Polygon stays a 4-point rectangular bounding box for all shapes.
- **Schema `polygon` maxItems=4 unchanged** (still rectangular-only).
- **Polygon consumers / denormalization** — remains GAP-ANALYSIS §5 Decision 2
  (open).

## Risks

Very low: purely additive (one helper + validator enhancement + one assertion),
does not alter existing polygon data. Existing 37 tests unaffected (their
polygons are all canonical rectangles, which pass the new check).

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 1, US-1.2.
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 US-1.2, §4 P0 US-1.2 row.
- Predecessor: US-1.1 (issue #48, PR #49) — the `schema_extractor.py` this builds on.
