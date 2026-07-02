# PLAN-GIT-63 — US-4.3: Auto-Chain Extraction (Templated Output)

**Issue**: #63
**Branch**: GIT-63 (base: dev)
**Priority**: Must Have (P0)
**Status**: Planned (rev 2 — architecture-review APPROVE-WITH-CHANGES incorporated: M1 source-from-input, M2 staleness-aware, m3 atomic-simplify, m4 field-rename, m6 agent exception-safety)

## Goal

Make every generated `.pptx` **self-describing / templated** (carrying `ppt/template_schema.json`), and surface a one-line "extracting first, then generating" status message to the user when the input template lacks embedded JSON. This closes US-4.3's only real gaps: **AC2** (output carries the embedded schema) and **AC3** (status message). **AC1** (single prompt, no error) is **already functionally met** after US-4.1 — `get_render_contract` falls back to sidecar introspection, so any PPTX renders in one `generate_ppt_from_data` call.

## Strategic Context

GAP-ANALYSIS §US-4.3 is rated "⚪ Architecture differs" — the sidecar fallback makes the "no JSON → extract first" mechanism unnecessary. But the implicit path does not satisfy the ACs: confirmed that existing output PPTXs are all `embedded: NO` (python-pptx drops the unmodeled `ppt/template_schema.json` part on `prs.save`, PLAN-GIT-58 C1), and there is no user-facing message. Constraint: the subagent's `task` permission is `ppt-template-filler: allow, "*": deny` (`pptx-subagent.md:10-12`), and generate-template-skill is **interactive** (US-3.2 title confirmation) — so literally chaining the skill as a subagent in a headless one-shot flow is infeasible → use **engine inline** (decision Q1).

## Architecture Decisions (locked)

1. **Engine inline auto_template (decision Q1):** `generate_ppt_from_data` gains `auto_template: bool = True`. After `prs.save(output)`, if `auto_template`: resolve a schema source → embed into the output → record the result. `embed_schema` is **already atomic** (it writes to its own temp in `out_path.parent` then `os.replace`), so it is called directly in place (`embed_schema(output, schema, output)` — no extra temp hop, arch-review m3). Non-fatal (try/except; failure is debug-log only, never blocks render). The generate-template-skill's interactive title-confirmation is skipped (appropriate for headless auto-chain).
2. **Always re-embed (decision Q2):** regardless of whether the input is templated, the output carries the embedded schema — every generated deck becomes a reusable template. The AC3 message fires only when the **input** lacked embedded JSON.
3. **Schema source resolution chain (arch-review M1 + M2):** always source from the **input template**, never the rendered output:
   - ① input has a **valid, non-stale** embedded schema (present AND the M5 staleness guard does not fire — embedded layout signature matches the live template) → copy it (`read_embedded_schema(template)`) → `schema_source="copied_input"`.
   - ② otherwise (absent, **or stale** per M2) → `extract_schema(template)` → `schema_source="extracted_input"`. **Crucially this extracts from the *input template*, not the output** (M1): the input's master + layouts are byte-identical to the output's (python-pptx does not mutate them on save), and `_infer_title` reads the *template's* own `slides[0]`/core title — the template's identity — **not the rendered deck's cover slide**. (Extracting from the output would seed `template_metadata.title` with the deck's cover title — a content leak + wrong reuse identity. Confirmed: `_infer_title` reads `prs.slides[0]`, `schema_extractor.py:472-479`.) The M2 stale case also routes here so a stale embedded schema is **not** laundered into a "trusted" output; the fresh extraction reflects the actually-rendered structure.
4. **Render report gains a `templating` field (additive, backward compatible; field name per arch-review m4):** `{input_template_embedded:bool (was `input_source`), output_templated:bool, schema_source:"copied_input"|"extracted_input"|"failed", message:str}`. Engine return type unchanged (still returns the path string).
5. **AC3 is emitted by the agent at Stage 0:** the agent runs a `read_embedded_schema` check **wrapped in try/except `TemplateExtractionError`** (a corrupt-zip template raises; arch-review m6 — treat a raise as "absent" + warn, never crash Stage 0 before the engine's own robust path runs); when the input lacks embedded JSON, it surfaces "No template found — extracting first, then generating slides..." before proceeding. The engine's auto_template does the actual extraction/embedding; the agent only informs.
6. **Ordering:** `save → auto_template (re-embed output) → write render report (incl. templating field) → cleanup_temp → return path`.
7. **Do not touch input-side resolution:** US-4.3 does not template the input (sidecar fallback yields identical layout results); US-4.2's `schema_font_map` for a non-templated input still falls to role ceilings (already supported).

## Deliverables

**Change** `.opencode/skills/ppt-template-filler/scripts/ppt_builder.py`
- `generate_ppt_from_data(..., auto_template: bool = True)`: after save, call new `_ensure_output_templated(output, template, auto_template) -> dict` and fold its result into the render report.
- New `_ensure_output_templated(...)`: resolve schema from the **input template** — `emb = read_embedded_schema(template)`; if `emb` is non-None **and** not stale (reuse the `_warn_if_embedded_stale` check) → `schema = emb`, `schema_source="copied_input"`; **else** → `schema = extract_schema(template)`, `schema_source="extracted_input"` (M1/M2). Then `embed_schema(output, schema, output)` in place (m3 — `embed_schema` is already atomic; no extra temp hop). Return `{input_template_embedded, output_templated, schema_source, message}`; on failure → `{output_templated:False, schema_source:"failed", message:<err>}`.
- Merge the `templating` section into the report written by `_write_render_report`.

**Change** `.opencode/agents/pptx-subagent.md`
- Stage 0: add an "is the template templated?" detection step (`read_embedded_schema`, **wrapped in try/except `TemplateExtractionError`** — arch-review m6); when embedded JSON is absent, emit the AC3 message before proceeding.
- "What NOT to Handle": add a boundary clarification — generating slides **from a non-templated file** is this agent's job (US-4.3 auto-tempates the output); pure **extraction / fingerprinting** intent still routes to generate-template-skill.

**Tests** `tests/test_auto_template.py`
- Non-templated input → `read_embedded_schema(output)` non-None; report `input_template_embedded=False`, `output_templated=True`, `schema_source="extracted_input"`.
- Templated input → output still templated; report `input_template_embedded=True`, `schema_source="copied_input"`.
- Stale templated input (M2) → output schema is fresh-extracted (`schema_source="extracted_input"`), not the stale embedded copy.
- Output schema identity (M1) → the embedded schema's `template_metadata.title` is the **template's** identity, not the rendered deck's cover (`title_source` from the template, not `"slide1"` of the deck).
- `auto_template=False` → output NOT templated (opt-out preserved).
- Extraction failure (synthetic corrupt) → non-fatal; render still succeeds; report `output_templated=False` / `schema_source="failed"`.

## Acceptance Criteria (US-4.3) — to deliver

- [ ] AC1 — Single user prompt triggers the full flow without error (functionally met post-US-4.1; engine auto_template guarantees a templated output).
- [ ] AC2 — Intermediate JSON is embedded in the output PPTX (post-save re-embed; `read_embedded_schema(output)` non-None).
- [ ] AC3 — User sees a status message (agent Stage 0 emits it when the input lacks embedded JSON).

## Implementation Phases

### Phase 1: Engine auto_template + tests
- [ ] T1: `_ensure_output_templated` (schema-source chain + atomic re-embed + non-fatal); `generate_ppt_from_data` gains `auto_template=True` and wires it (save → auto_template → report).
- [ ] T2: render report gains the `templating` field.
- [ ] T3: `test_auto_template.py` — four cases (non-templated / templated / opt-out / failure non-fatal).

### Phase 2: Agent detection + message
- [ ] T4: `pptx-subagent.md` Stage 0 detection + AC3 message; "What NOT to Handle" boundary clarification.

### Phase 3: Docs
- [ ] T5: `chenyu-user-stories.md` US-4.3 AC → `[x]`; `GAP-ANALYSIS.md` Rev 11 (US-4.3 ✅ Met; counts Met 12 / Partial 4 / Not met 3 / Differs 0); AGENTS.md / README.md one-line note.

## Test matrix

| Case | Expected |
| --- | --- |
| non-templated input render | `read_embedded_schema(output)` non-None; report `input_template_embedded=False`, `schema_source="extracted_input"` |
| templated input render | output still templated; report `input_template_embedded=True`, `schema_source="copied_input"` |
| stale templated input (M2) | output carries a **fresh** schema (`schema_source="extracted_input"`), not the stale embedded one |
| output schema describes the template, not the deck (M1) | `extract_schema(template)`-sourced schema's `template_metadata.title` is the **template's** identity (not the rendered cover); `title_source != "slide1"` from the rendered deck |
| `auto_template=False` | output NOT templated; report `output_templated=False` |
| extraction failure | non-fatal; render succeeds; report `schema_source="failed"` |
| output schema excludes rendered slides | `extract_schema(output)` covers master + layouts only (no slide-content leak) |

## Verification

```powershell
python -m pytest .opencode/skills/ppt-template-filler/scripts/tests/ -q
python -c "import sys; sys.path.insert(0,'.opencode/skills/ppt-template-filler/scripts'); from schema_extractor import read_embedded_schema; print(read_embedded_schema('<output>.pptx') is not None)"
```

## Out of Scope / Open Questions

- **Templating the input side** (making layout resolution use embedded) — sidecar fallback yields identical results; unnecessary.
- **Reusing generate-template-skill's interactive title confirmation** — inappropriate for headless auto-chain; the inferred title is accepted.
- **`title_source` user confirmation on the output schema** — US-3.2's interactive flow; US-4.3 accepts the inferred value.

## Risks

- **python-pptx dropping the part again** — the re-embed happens **after** `prs.save` (an order-preserving zip rewrite via `embed_schema`, never re-opened by python-pptx), so it is not stripped.
- **`extract_schema(template)` cost / failure** — parses the input template's master + layouts only (lightweight, O(layouts)); failure is non-fatal and degrades to "output not templated" with a report note. Sourcing from the **input** (not the output) avoids `_infer_title` reading the rendered deck's cover slide (M1).
- **Report `templating` field breaking existing consumers** — additive top-level key; existing `slides` consumers are unaffected.
- **AC1 "triggers both skills" literal mismatch** — the mechanism is engine inline rather than literally chaining the skill; GAP-ANALYSIS already accepts "architecture differs, function met". Documented.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` US-4.3 (L323-337); GAP-ANALYSIS §US-4.3 (L131-133).
- Predecessors: US-4.1 (#58, `get_render_contract` sidecar fallback); US-1.5 (`embed_schema` / `read_embedded_schema`); US-3.1 (generate-template-skill).
- Key code: `ppt_builder.generate_ppt_from_data` (save + render report), `schema_extractor.embed_schema` / `extract_schema` / `read_embedded_schema`.
- PLAN format template: `PLANS/PLAN-GIT-58.md` / `PLANS/PLAN-GIT-60.md`.
