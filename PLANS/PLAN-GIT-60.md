# PLAN-GIT-60 — US-4.2: Reactive Text-Fitting (Auto-Shrink to Fit)

**Issue**: #60
**Branch**: GIT-60 (base: dev)
**Priority**: Must Have (P0)
**Status**: Planned

## Goal

Close the **single hard gap** in US-4.2: reactive per-textbox font auto-shrink (−2pt steps down to an 8pt floor). The existing `density_mode.py` (per-slide word budget) + `constraint_checker` (`content_area_in2`) are **preventive**; US-4.2 adds the **reactive** backstop — at render time, estimate whether text overflows its placeholder polygon and, if so, shrink from the template's base font size in steps until it fits. It also makes body base size / line spacing **derived from the template** (fixing the pre-existing `_set_body_text` hardcode of Pt(14)/Pt(12)), and records a `font_size_adjusted` flag into an accompanying render report `<output>.render.json`.

Satisfies the three ACs: ① no overflow ② size only ever reduced (≤ template's defined size) ③ a `font_size_adjusted` flag is set when auto-sizing occurs.

## Strategic Context

GAP-ANALYSIS §US-4.2 (L121-123): a grep for `font_size` / `auto.*shrink` / `adjust` in `density_mode.py` returns no matches — there is no reactive shrink and no flag is set. Epics 1/3/4.1 are complete and the embedded schema is now consumable; its components carry `font.size_pt` (`schema_extractor._extract_text_fonts`, L582-628) and the placeholder polygon carries geometry. US-4.1 routes generation through `add_slide(layout)`, so the **live placeholder** (`ph.width` / `ph.height` in EMU) is also an available geometry source.

**Core technical constraint:** python-pptx has **no text-measurement engine** — it cannot measure rendered text height at a given font size (PowerPoint does the real layout). Text-fitting therefore uses a **calibrated heuristic estimator** (the story's "−2pt steps until it fits" implies exactly this).

## Architecture Decisions (locked)

1. **New `text_fit.py` pure estimator** (no PPTX dependency → unit-testable in isolation). `fit_font_size(text, base_pt, w_in, h_in, …) -> FontFit{applied_pt, adjusted, fits, base_pt}`: starting from `base_pt`, if it does not fit, step −2pt until it fits or the 8pt floor is reached; if still not fit at 8pt, stay at 8pt with `fits=False` (log an overflow WARNING; do not go smaller).
2. **Character-width model:** Latin char advance ≈ 0.5 × pt / 72 in; CJK ≈ 1.0 × pt / 72 in (reuse `density_mode._CJK_RE`). `chars_per_line ≈ (w_in − 2×padding) × 72 / (ratio × pt)`; compute `lines` per segment (respecting explicit `\n`); `est_height_in = lines × pt × line_spacing_factor / 72`; fits ⟺ `est_height ≤ h_in − 2×padding`. padding = 0.1 in; default line-spacing factor 1.2. **Calibration self-check:** a 9 in × 4 in body box @ 18 pt → ~72 chars/line × ~13 lines ≈ 150 words, matching the `text-heavy` budget.
3. **Base font size derived from the template (decision Q2 = Yes):** resolution chain = ① embedded-schema matching component `font.size_pt` (explicit, non-null) → ② live placeholder inherited size (python-pptx best-effort) → ③ role default (body 18 / title 32 / subtitle 20). At the top of the render loop, call `read_embedded_schema(template)` once and build a `layout_name → {role → font.size_pt}` lookup.
4. **`font_size_adjusted` written to an accompanying render report (decision Q1 = sidecar):** at end of render, write `<output>.render.json`: `{slides:[{index, slide_type, placeholders:[{role, field, template_size_pt, applied_size_pt, font_size_adjusted, fits, lines_estimated}]}]}`. **Do not change `generate_ppt_from_data`'s return type** (it still returns the path string) → fully backward compatible; subagent / SKILL reads the report on demand.
5. **Line spacing derived from the template:** at render time, read the live layout placeholder's `paragraph.line_spacing` (python-pptx can read it); otherwise default 1.2. **Do not modify `schema_extractor`** (avoids a re-extract + re-embed cycle) — listed as an optional enhancement.
6. **Bullet style preserved via inheritance:** `add_slide(layout)` already inherits the layout's bullet style; this story **does not actively override** it (word_wrap=True, no `buNone`); only apply a clean default (•) when the template provides none. Low priority; verification-focused.
7. **Scope:** all text placeholders (title / subtitle / body / body_left / body_right). title/subtitle rarely trigger shrink, but are included so AC1 (no overflow) holds globally.
8. **Floor 8 pt (decision Q3):** below 8 pt do not shrink further; log a WARNING.

## Deliverables

**New** `.opencode/skills/ppt-template-filler/scripts/text_fit.py`
- Constants: `MIN_FONT_SIZE_PT=8`, `FONT_STEP_PT=2`, `LINE_SPACING_DEFAULT=1.2`, `TEXT_PADDING_IN=0.1`, `LATIN_RATIO=0.5`, `CJK_RATIO=1.0`.
- `FontFit` dataclass (`applied_size_pt, base_size_pt, adjusted, fits, lines_estimated`).
- `estimate_lines(text, font_pt, width_in, …) -> int`; `estimate_height_in(lines, font_pt, line_spacing) -> float`; `fits_at_size(...) -> bool`; `fit_font_size(...) -> FontFit`.

**Change** `.opencode/skills/ppt-template-filler/scripts/ppt_builder.py`
- New `_resolve_base_font_pt(shape, role, schema_font_map, layout_name) -> float` (decision 3 resolution chain).
- New `_apply_text_with_fit(shape, text, role, *, schema_font_map, layout_name, box_in) -> FontFit | None`: resolve base size → `fit_font_size` → write runs at the applied size (preserve `_set_body_text`'s bold-title/desc split, but with template-derived sizes: title-run = applied_pt, desc-run = applied_pt × 0.85) → set `tf.word_wrap=True` → return the FontFit.
- Refactor `_set_text` / `_set_body_text` to delegate to `_apply_text_with_fit` (accept geometry: `shape.width` / `shape.height` EMU → inches).
- Render loop: at top, `schema = read_embedded_schema(template)` (best-effort; on absence, all role-defaults); collect a FontFit per fill of title/subtitle/body/two-body → `_write_render_report(output, report)`.
- New `_write_render_report(pptx_path, report)`: writes `<stem>.render.json` (failure is debug-log only; never blocks the render).

**Change** `pptx-subagent.md` / `ppt-template-filler/SKILL.md`: Stage 4 mentions the render-report path + the meaning of `font_size_adjusted` (one line each).

**Tests** `tests/test_text_fit.py` (pure) + extend `tests/test_render_contract.py` / new `test_text_fitting.py` (integration).

## Acceptance Criteria (US-4.2) — to deliver

- [ ] AC1 — No generated slide has text overflowing its bounding box (guaranteed by the estimator; if it still overflows at 8 pt, a WARNING is logged — acceptable degradation, but should not trigger on default inputs).
- [ ] AC2 — Font size is only reduced when necessary — never increased beyond the template's defined size (base = template-derived; only −2pt downward steps).
- [ ] AC3 — A `font_size_adjusted` flag is set in the accompanying render report when auto-sizing occurs.

## Implementation Phases

### Phase 1: Estimator + pure unit tests (text_fit.py)
- [ ] T1: `text_fit.py` — `FontFit` + `estimate_lines` + `estimate_height_in` + `fits_at_size` + `fit_font_size` (−2pt stepping, 8pt floor, CJK/Latin mixed-width model, explicit-`\n` respected).
- [ ] T2: pure unit tests — fit / no-fit boundaries, stepping sequence (24→22→…→8), floor clamp (not fit at 8 → fits=False/applied=8), CJK occupies two width units, explicit newline, empty text, line-spacing-factor influence.

### Phase 2: Wire into the render path + template base size + word_wrap + render report
- [ ] T3: `_resolve_base_font_pt` (embedded-schema component font.size_pt → live placeholder inheritance → role default); at the top of the render loop build `schema_font_map` (best-effort `read_embedded_schema`).
- [ ] T4: `_apply_text_with_fit` refactors `_set_text` / `_set_body_text` (template-derived sizes, bold-title/desc split preserved, word_wrap=True, returns FontFit).
- [ ] T5: `_write_render_report` (`<output>.render.json`) + render loop collects a FontFit per slide per placeholder; return type unchanged.
- [ ] T6: integration tests — long body triggers shrink and report `font_size_adjusted=true`; short text `adjusted=false`; size never exceeds template; render-report schema validated.

### Phase 3: Spacing / bullets + regression + end-to-end
- [ ] T7: line spacing read from the live layout placeholder `paragraph.line_spacing` (else 1.2); bullet-inheritance verification (do not override; add default • when absent).
- [ ] T8: regression — compare existing body / title / two-content / chart output appearance; the 112 schema_extractor + density + render suites stay green.
- [ ] T9: end-to-end — render a long-content deck from the bundled `template.pptx`, verify no overflow + report generated + output stays 100% native/editable (python-pptx).

### Phase 4: Docs
- [ ] T10: `chenyu-user-stories.md` US-4.2 AC → `[x]`; `GAP-ANALYSIS.md` Rev 10 (US-4.2 ✅; counts Met 12 / Partial 3); `AGENTS.md` / `README.md` add a text-fitting + render-report note.

## Test matrix

| Case | Expected |
| --- | --- |
| `fit_font_size` fit boundary | text just fits → applied=base, adjusted=False |
| shrink stepping sequence | 24→22→20→… until fit, adjusted=True, applied=base−2k |
| 8 pt floor clamp | still not fit at 8 → applied=8, fits=False, adjusted=True |
| CJK width | at same size/width, CJK lines ≥ Latin (each char = two width units) |
| template base size derivation | non-null embedded font.size_pt → used; null → role default; never exceeds template |
| long body integration | applied<base, `font_size_adjusted=true`, word_wrap=True |
| render-report schema | `<output>.render.json` exists + each placeholder carries 6 fields + render not blocked |
| regression | 112 extractor + density + render green; no unexpected output-appearance change |

## Verification

```powershell
python -m pytest .opencode/skills/ppt-template-filler/scripts/tests/ -q
python -c "import sys; sys.path.insert(0,'.opencode/skills/ppt-template-filler/scripts'); from text_fit import fit_font_size; print(fit_font_size('x'*400, 18, 9.0, 4.0))"
```

## Out of Scope / Open Questions

- **Extending `schema_extractor` to capture line_spacing / space_before / space_after** (optional enhancement; reading the live placeholder is sufficient this round).
- **True text measurement** (requires LibreOffice/COM rendering) — out of scope; the calibrated estimator suffices.
- **In-chart / in-table text fitting** — placeholder text only; chart font sizes are governed by the `_CHART_*` constants.
- **US-4.6 multi-aspect-ratio** — separate story (coordinate placement); not here.

## Risks

- **Estimator drift** — the heuristic is imperfect; `word_wrap=True` is the hard backstop (guarantees the box width is never exceeded), and the height estimate is kept conservative (prefer shrinking an extra line). The calibration self-check aligns with the `text-heavy` budget.
- **Regressing `_set_body_text` font sizes** (decision Q2) — the current Pt(14)/Pt(12) → template-derived change alters output appearance; regression test T8 compares + the bold-title/desc ratio is preserved (desc = 0.85 × title).
- **Embedded `font.size_pt` often null (inherited)** — falls back to the role default; deterministic and testable, but the "template's defined size" semantics are weaker (the inherited value is not explicitly extracted). Mitigation: T3 also reads the live placeholder's inherited size as a secondary source.
- **Render-report write failure** — debug-log only; never blocks the render (same pattern as outline cleanup).

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` US-4.2 (L303-319); GAP-ANALYSIS §US-4.2 (L121-123, Rev 9).
- Predecessors: US-4.1 (#58, embedded schema now consumable); US-1.4 (component font.size_pt).
- Key code: `ppt_builder.py:464-525` (_set_text/_set_body_text), `:965-1033` (render loop), `schema_extractor.py:582-628` (font extraction), `density_mode.py` (CJK regex reuse), `contract_adapter._build_placeholders` (geometry width_in/height_in).
- PLAN format template: `PLANS/PLAN-GIT-58.md`.
