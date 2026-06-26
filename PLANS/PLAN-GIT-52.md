# PLAN-GIT-52 — US-1.3: Component Type Enumeration

**Issue**: #52
**Branch**: GIT-52 (base: dev)
**Priority**: Must Have (P0)
**Status**: Implemented (v2 — architecture review findings incorporated; all phases complete, 64 tests green)

## Goal

Close the two deferred **Details** items that keep US-1.3 at 🟡 Partial, bringing
it to ✅ Met. All three Acceptance Criteria are already satisfied by
`schema_extractor.py map_shape_type()` (US-1.1 / PR #49); this plan delivers the
remaining Details: (1) emit `type_confidence` (always; `"low"` for unrecognized /
indeterminate), and (2) make the `"audio"` enum value reachable by splitting
MEDIA into audio/video via OOXML inspection.

## Strategic Context

Builds on US-1.1 (issue #48, PR #49) and US-1.2 (issue #50, PR #51).
`map_shape_type()` already applies the full 10-value enum to all elements
(placeholders and freeform shapes). The `type` field is always present and never
null/"unknown" — the ACs are met. The component model is **metadata-only** (no
consumer reads `type_confidence`), so this is about schema fidelity / contract
hygiene, not rendering. GAP-ANALYSIS §2 US-1.3 grades this 🟡 Partial solely on
the two Details.

## Component Type — Recognition vs Fallback (correctness rule, v2)

The current `map_shape_type` `else` branch returns `"shape"` for BOTH recognized
preset-geometry shapes (`AUTO_SHAPE`, `FREEFORM`, `LINE` — rectangles, arrows,
connectors) AND genuinely unrecognized types. A naive `type_confidence: "low"`
on the `else` branch would **mislabel recognized shapes as low-confidence**.

**Architecture review (MAJOR-1)** established that `MSO_SHAPE_TYPE` has **26
members**, so an explicit whitelist (`AUTO_SHAPE`/`FREEFORM`/`LINE`) would dump
~13 *other recognized* members (`LINKED_PICTURE`, `TEXT_EFFECT`, `CALLOUT`,
`DIAGRAM`, `CANVAS`, `INK`, …) into `low` — exactly the silent mislabeling US-1.3
exists to prevent. **v2 decision (MAJOR-1 option a): drop the whitelist.** The
rule is now a clean 2-way split keyed on whether `shape_type` is readable:

- **Any non-`None` `shape_type`** → the element *was* recognized by python-pptx →
  confidence **`"high"`**. Mapped members (`PICTURE`/`GROUP`/`TEXT_BOX`/
  `IGX_GRAPHIC`/`MEDIA`/`WEB_VIDEO`) get their specific type; everything else
  falls through to `"shape"/"high"` (recognized shape, not further sub-typed).
- **`shape_type` is `None` OR `.shape_type` access raises** → genuinely
  unrecognized → `"shape"` + **`"low"`**.

This keeps the prose ("low ⟺ unrecognized") and the code in agreement, and makes
production `low` rare (effectively: MEDIA with no marker).

## Media Subtype — Audio/Video Split

`python-pptx` collapses audio and video into `MSO_SHAPE_TYPE.MEDIA`; it exposes
no audio/video subtype. The subtype is recoverable from OOXML: a media `<p:pic>`
carries `<a:audioFile>` or `<a:videoFile>` under `<p:nvPicPr>/<p:nvPr>`. Rule:
- `<a:audioFile>` present → `"audio"` + `"high"`.
- `<a:videoFile>` present → `"video"` + `"high"`.
- MEDIA with **neither marker** → `"video"` + `"low"` (ambiguous guess; preserves
  the historical default).

> **MINOR-4 (reachability note):** real python-pptx shapes always carry
> `_element`, so the `"no _element"` guard is reachable **only** via `FakeShape`
> in tests. The *production-reachable* `low` path is narrower: a real MEDIA shape
> whose `_element` lacks both `<a:audioFile>` and `<a:videoFile>` (e.g. some
> embedded-media / placeholder-media representations). The Task-11 integration
> test should therefore assert at least one `low` end-to-end if such a shape
> exists in the bundled template, not merely membership in `{high, low}`.

## Architecture Decisions (locked, v2)

1. **`type_confidence` is always emitted** by the extractor ("high" default,
   "low" for unrecognized/indeterminate), but **optional in the JSON Schema** so
   existing/old extracted data still validates (backward compatible).
2. **Keep `map_shape_type(shape) -> str` as a backward-compatible wrapper**; add
   internal `_classify_shape(shape) -> (type, confidence)`. Existing ~9 unit
   assertions on `map_shape_type` stay green.
3. **No whitelist (MAJOR-1 option a).** Non-`None` `shape_type` → confidence
   `"high"`; only `None`/unreadable → `"low"`. `_SHAPELIKE_TYPES` is **not**
   introduced.
4. **`WEB_VIDEO` → `("video", "high")`** — a free one-liner under decision 3
   (was deferred in v1); pulled in to avoid mislabeling.
5. **Audio detection via OOXML** (`shape._element` + `qn("a:audioFile")`/
   `qn("a:videoFile")`); synthetic-XML unit test (bundled template has no audio).
   **NIT-2:** `_media_classification` uses the idiomatic `qn()`; the existing
   theme code (`_NS_A`/`etree.QName`) is **not** refactored (out of scope).
6. **shape/low WARNING (MINOR-2).** A `type=="shape"` + `type_confidence=="low"`
   component emits a **non-fatal `ValidationIssue(severity="warning")`** so
   `is_valid` stays True but the unrecognized element is surfaced ("flagged for
   review") — mirroring the US-1.2 degenerate-polygon pattern
   (`schema_extractor.py:653-658`). This delivers the field's stated purpose.
7. **Validator enum check** — `type_confidence` (when present) must be
   `"high"`/`"low"` (error otherwise); do NOT add it to `_COMPONENT_REQUIRED`.
8. **Incremental edits to existing files** — no new files, zero changes to the
   render path (`ppt_builder.py` / `template_introspector.py` untouched).
   `component_type_enum` already includes `"audio"` — no enum change.

## Deliverables (all incremental edits, zero new files)

- `scripts/schema_extractor.py`: `from pptx.oxml.ns import qn`;
  `_media_classification()`; `_classify_shape()`; `map_shape_type` wrapper; emit
  `type_confidence` in `_build_component`; enum check + shape/low WARNING in
  `_validate_component`; **update the stale MEDIA-branch comment at `:196-198`**
  (NIT-1 — D2 makes "audio unreachable" false).
- `scripts/schemas/template_schema.json`: optional `type_confidence` property on
  `$defs/component`.
- `scripts/tests/test_schema_extractor.py`: import `_classify_shape`;
  `FakeShape._element`; `_ok_component()` update; `_media_element()` helper;
  `TestTypeConfidence`, `TestAudioVideo`, validation tests (incl. shape/low
  WARNING), integration test.
- `docs/user-stories/GAP-ANALYSIS.md`: US-1.3 → ✅ Met (Met 3→4, Partial 6→5);
  §4 P2 US-1.3 row stale note.
- `docs/user-stories/chenyu-user-stories.md`: (US-1.3 ACs already `[x]`; no
  checkbox change — Details closure noted in GAP).

## Acceptance Criteria (US-1.3) — already satisfied

- [x] `type` field is always present and always one of the defined enum values.
- [x] No component has `type: null` or `type: "unknown"`.
- [x] A mapping table from OOXML tags to enum values is included in the skill
      source.

## Details to implement (this plan's scope)

- [x] D1: emit `type_confidence` always ("high" default; "low" for unrecognized
      + indeterminate MEDIA).
- [x] D2: make `"audio"` reachable (audio/video split via
      `<a:audioFile>`/`<a:videoFile>`).
- [x] D3: non-`None` `shape_type` → `"high"` (no whitelist); only `None`/
      unreadable → `"shape"/"low"` (MAJOR-1 option a).
- [x] D4: optional `type_confidence` in `template_schema.json`.
- [x] D5: validator enum check + **shape/low WARNING (MINOR-2)** + tests.
- [x] D6: map `WEB_VIDEO` → `("video","high")`.
- [x] D7: update the stale MEDIA-branch comment (NIT-1).

## Implementation Phases

### Phase 1: Classification refactor + type_confidence

- [x] Task 1: Add `from pptx.oxml.ns import qn`. (No `_SHAPELIKE_TYPES` —
      decision 3.)
- [x] Task 2: Add `_media_classification(shape) -> (subtype, had_marker)`
      (subtype `audio`/`video`; searches `shape._element` via
      `el.find(".//" + qn("a:audioFile"))` / `qn("a:videoFile")`; returns
      `("video", False)` when absent or no `_element`).
- [x] Task 3: Add `_classify_shape(shape) -> (type, confidence)` with precedence:
      placeholder/table/chart → high; `shape_type` access exception →
      `(shape, low)`; `st is None` → `(shape, low)`; `PICTURE`/`GROUP`/
      `TEXT_BOX`/`IGX_GRAPHIC` → high; `WEB_VIDEO` → `(video, high)`; MEDIA →
      `(sub,high) if had else ("video",low)`; **else (any other non-None) →
      `(shape, high)`**.
- [x] Task 4: Convert `map_shape_type` to a wrapper returning
      `_classify_shape(shape)[0]`.
- [x] Task 5: In `_build_component`, emit
      `component["type_confidence"] = confidence`.

### Phase 2: Schema + validator

- [x] Task 6: Add optional `type_confidence` (`enum:["high","low"]`) to
      `template_schema.json` `$defs/component` (NOT in `required`).
- [x] Task 7: In `_validate_component`, after the type-enum check: (a) if
      `type_confidence` present and not in `("high","low")` → error; (b) **if
      `type=="shape"` and `type_confidence=="low"` → non-fatal warning**
      ("unrecognized element — flagged for review").

### Phase 3: Tests

- [x] Task 8: Import `_classify_shape`, `etree`; add `self._element = None` to
      `FakeShape`; add `"type_confidence": "high"` to `_ok_component()`; add
      `_media_element(marker)` helper.
- [x] Task 9: `TestTypeConfidence` (recognized→high; placeholder→high;
      AUTO_SHAPE/FREEFORM/LINE→(shape,high); **any other non-None e.g.
      CALLOUT/LINKED_PICTURE→(shape,high)**; WEB_VIDEO→(video,high); unknown
      st=None→(shape,low); shape_type raises→(shape,low); MEDIA no marker→
      (video,low)).
- [x] Task 10: `TestAudioVideo` (audioFile→(audio,high); videoFile→(video,high);
      `map_shape_type` returns `"audio"`).
- [x] Task 11: Validation tests — bad `type_confidence`→error; missing→valid;
      **shape/low → valid + 1 warning (MINOR-2)**; integration (bundled
      template: every component `type_confidence` ∈ {high,low}, and ≥1 `low`
      flagged if such a shape exists).

### Phase 4: Docs

- [x] Task 12: Update the stale MEDIA-branch comment in `map_shape_type`
      (`schema_extractor.py:196-198`) to reflect the audio/video split (NIT-1).
- [x] Task 13: GAP-ANALYSIS §2 US-1.3 → ✅ Met (Met 3→4, Partial 6→5); note §4
      P2 US-1.3 row is partly stale ("extend extraction beyond placeholders"
      already done).

## Test matrix (expected (type, confidence))

| Case | Expected |
| --- | --- |
| placeholder / table / chart / image / group / textbox / smartart | (*, high) |
| WEB_VIDEO | (video, high) |
| any other non-None (AUTO_SHAPE / FREEFORM / LINE / CALLOUT / LINKED_PICTURE / INK / …) | (shape, high) |
| shape_type=None or unreadable | (shape, low) + WARNING |
| MEDIA + audioFile | (audio, high) |
| MEDIA + videoFile | (video, high) |
| MEDIA no marker | (video, low) |

## Verification

```bash
# from .opencode/skills/ppt-template-filler/scripts
python -m pytest tests/test_schema_extractor.py -q
python -m pytest tests/ -q
python schema_extractor.py --input templates/template.pptx --output /tmp/s.json
```

## Out of Scope (deferred) / OPEN QUESTIONS

- **`LINKED_PICTURE`/`TEXT_EFFECT`/`DIAGRAM`/`CANVAS`/`INK`…** → remain
  `"shape"/"high"` (recognized but not further sub-typed). Finer promotion
  (e.g. `LINKED_PICTURE`→`image`) is intentionally deferred — the enum has no
  finer types and this keeps scope minimal (MAJOR-1 option a, not "a +
  promotions").
- **Cross-field confidence consistency** beyond the shape/low WARNING — e.g.
  flagging externally-supplied schemas where `type=="shape"` and
  `type_confidence=="high"` — not enforced; the WARNING (decision 6) is the
  scoped deliverable.
- **Real audio fixture** — covered by synthetic XML; python-pptx exposes no
  `add_audio`, so no audio `.pptx` is added. (A real *video* fixture via
  `add_movie` was a review suggestion; deferred — synthetic XML guards the
  descendant-search path.)
- **Future promotion of `type_confidence` to `_COMPONENT_REQUIRED`** — kept
  optional now (backward compat); revisit once all in-flight extracted artifacts
  are regenerated by the new extractor.
- **Pre-existing validator/schema divergence** — `template_schema.json` declares
  `additionalProperties:false` and `pattern:"^comp_\\d{3,}$"` that the
  hand-rolled `_validate_component` does not enforce. Each new optional field
  widens the manual-sync surface; consolidating (load the schema / generate the
  validator) is **US-5.2** scope, not US-1.3.

## Risks

- **Backward compatibility — low.** schema_extractor output is generated (not
  committed); `type_confidence` is optional in the schema, so
  previously-extracted/old data still validates. No committed artifact changes
  shape.
- **Existing tests unaffected.** The `map_shape_type` wrapper preserves the ~9
  `== "video"`/`== "shape"` assertions; `FakeShape` MEDIA (no `_element`) still
  returns `"video"`.
- **Recognized-shape mislabeling (avoided — v2).** The default-`high` rule
  (decision 3) ensures every non-`None` `shape_type` is `high`; `low` is
  reserved strictly for `None`/unreadable + indeterminate MEDIA. Covered by
  `TestTypeConfidence` (incl. `CALLOUT`/`LINKED_PICTURE` → high).
- **Audio reachability is synthetic-only.** No bundled audio shape exists, so
  the audio branch is exercised solely by the synthetic-XML test (intentional;
  documented). The video descendant-search path is also synthetic; a real
  `add_movie` fixture is deferred.

## Code-review follow-ups

Post-implementation code review (verdict: Approve with revisions — 0 Critical/0
Major). All findings were Minor/Nit; the 4 applied are behavior-preserving polish:

- [x] **MINOR-4 — confidence constants.** Added `_CONFIDENCE_HIGH`/`_CONFIDENCE_LOW`
  module constants; `_classify_shape` and `_validate_component` now reference them,
  removing the three-way duplication of `"high"`/`"low"` (classifier ↔ validator ↔
  schema enum).
- [x] **MINOR-5 — `map_shape_type` docstring.** Trimmed to delegate to
  `_classify_shape` (no longer narrates the moved precedence chain).
- [x] **MINOR-1 + MINOR-3 — shape-only warning scope locked.** Added a comment at
  the `shape/low` WARNING documenting that indeterminate MEDIA (`video/low`) is
  intentionally NOT flagged (locked decision 6); added the negative test
  `test_video_low_emits_no_warning` to pin the boundary.
- Not applied (justified): **MINOR-2** (assert a real `low` on the bundled
  template) — the bundled template extracts all-`high` (140 image / 293
  placeholder / 60 shape / 5 textbox), so there is no `low` to assert.

Tests: `test_schema_extractor.py` → **65 passed** (was 64 after the +1 negative
test). No production-logic change.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 1, US-1.3.
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 US-1.3 (✅ Met), §4 P2.
- GitHub issue: #52 (`[US-1.3] Component Type Enumeration`).
- Predecessors: US-1.1 (issue #48, PR #49), US-1.2 (issue #50, PR #51).
- Format template: `PLANS/PLAN-GIT-50.md`.
- Architecture review: findings MAJOR-1, MINOR-2, MINOR-4(note), NIT-1/2/3
  incorporated above; OPEN QUESTIONS noted in §"Out of Scope".
- Code review: MINOR-4, MINOR-5, MINOR-1(+MINOR-3) applied — see
  §"Code-review follow-ups"; MINOR-2 not applied (justified).
