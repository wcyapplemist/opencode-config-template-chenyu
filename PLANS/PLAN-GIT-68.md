# PLAN-GIT-68 — US-2.1: Header & Footer Detection

**Issue**: #68
**Branch**: GIT-68 (base: dev)
**Priority**: Must Have (P0)
**Status**: Planned (rev 2 — architecture-review APPROVE-WITH-CHANGES incorporated: M1 read-path, M2 batch-prompts, m1 polygon-literal, m2 test-assertions)

## Goal

Detect whether the slide master carries header/footer placeholders, record `has_header`/`has_footer` booleans in `template_metadata.header_footer` (AC1), prompt the user when both are absent (AC2), and inject a user-requested default header zone into the schema (AC3, schema-only). This closes the last fully-unmet Must-Have.

## Strategic Context

`schema_extractor` already classifies HEADER/FOOTER/DATE/SLIDE_NUMBER as "chrome" (`map_placeholder_type`) but only to **filter them out** of fingerprints — never recorded. The `template_metadata.header_footer = {}` placeholder (and the `template_schema.json: {additionalProperties:true}` slot) await US-2.1. The bundled master has `FOOTER+SLIDE_NUMBER+DATE` but **no HEADER** → `has_header=False, has_footer=True`. No test asserts the empty dict, so populating it is safe.

> **OOXML note**: chenyu's "`<p:hdr>`/`<p:ftr>` elements" is a simplification — the slide master's header/footer are `<p:ph type="hdr"/>`/`type="ftr"/>` placeholders (python-pptx → `PP_PLACEHOLDER.HEADER`/`FOOTER`). Detection is by placeholder type (robust, native).

## Architecture Decisions (locked)

1. **Detection** (shared, `schema_extractor`): a new `_detect_header_footer(prs)` scans `prs.slide_masters[0].placeholders` for HEADER/FOOTER → returns `{has_header, has_footer}`. **Header/footer only** (strict AC1; date/slide_number not recorded). Master only (per chenyu's wording; layout-only chrome is out of scope).
2. **Prompt (AC2)** — **extraction skill primary**: `generate-template-skill` Stage 2 (after extract+validate, before embed): if both are false → primary-agent mode issues a `question`; headless/subagent mode skips. When **both** the title-source==filename condition **and** the header/footer-absent condition fire, batch them into a **single** `question` call (project convention; arch-review M2). Headless mode skips both. **+ generation agent light note** (arch-review M1): `pptx-subagent` Stage 0 reads `header_footer` via `read_embedded_schema` (**NOT** `get_render_contract` — the `contract_adapter` strips `template_metadata`). The note is **scoped to templated inputs only**; for a non-templated input `read_embedded_schema` returns `None` at Stage 0 (US-4.3's auto_template only produces the schema post-render), so the note is **deferred/skipped** — do NOT add a redundant `extract_schema` call (conflicts with US-4.3's "extract once" design). No injection on the generation path (it produces no template schema).
3. **Injection (AC3)** — `generate-template-skill`: if the user says yes, call `inject_default_header_zone(schema)` (in `schema_extractor`) → sets `header_footer.header = {source:"user_default", polygon:[{x:0,y:0},{x:1,y:0},{x:1,y:0.05},{x:0,y:0.05}], note:"Default header zone (top strip); metadata only — not rendered into the PPTX until a real header placeholder is added"}`. Persisted before embed. Schema-only, never touches the PPTX. The polygon is exactly 4 normalized `{x,y}` points (US-1.2 model; arch-review m1).
4. The "both-false → prompt" decision and the injection are **pure helpers** (unit-testable); the actual prompt is agent/LLM behaviour (SKILL.md instructions).

## Deliverables

**Change** `schema_extractor.py`:
- `_detect_header_footer(prs) -> {has_header, has_footer}` (scan master placeholders).
- `needs_header_footer_prompt(schema) -> bool` (both false).
- `inject_default_header_zone(schema)` (default zone marker; **English** note).
- Wire `_detect_header_footer` into `_build_metadata` (replaces `"header_footer": {}`).

**Change** `generate-template-skill/SKILL.md`: Stage 2 adds the header/footer check + prompt (primary-agent mode); when BOTH the title-confirm and header/footer conditions fire, batch into one `question` call (arch-review M2). Add the inject step.

**Change** `pptx-subagent.md`: Stage 0 reads `header_footer` via `read_embedded_schema` (NOT `get_render_contract` — adapter strips it; arch-review M1). If both false AND the input is templated, surface a one-line informational note (primary-agent mode; headless skips). Non-templated inputs: note deferred (no schema at Stage 0).

**Tests** (`test_schema_extractor.py` extend or new file):
- Bundled template → `has_header=False, has_footer=True`.
- Synthetic chrome-less master → both false; `needs_header_footer_prompt` True.
- `inject_default_header_zone` → `header_footer.header` has the 4-point polygon + English note; **explicit polygon assertions** (len==4, `{x,y}` keys, [0,1] range — arch-review m2); `validate_template_schema` still passes (no top-level breakage).

## Acceptance Criteria (US-2.1) — to deliver

- [ ] AC1 — `template_metadata.header_footer.has_header` / `.has_footer` are booleans reflecting actual detection.
- [ ] AC2 — When both are `false`, a user-facing question is output before continuing (generate-template-skill; pptx-subagent light note).
- [ ] AC3 — User "yes" → injects a default header zone into the JSON (schema-only; English note).

## Implementation Phases

### Phase 1: Detection + helpers + tests
- [ ] T1: `_detect_header_footer` + `needs_header_footer_prompt` + `inject_default_header_zone`; wire into `_build_metadata`.
- [ ] T2: tests — bundled (has_header=False/has_footer=True), synthetic chrome-less (both false), inject (default zone + English note + schema validates).

### Phase 2: Agent/skill wiring
- [ ] T3: `generate-template-skill/SKILL.md` Stage 2 check + prompt (primary-agent) + inject.
- [ ] T4: `pptx-subagent.md` Stage 0 light informational note (primary-agent; headless skips).

### Phase 3: Docs
- [ ] T5: `chenyu-user-stories.md` US-2.1 ACs → `[x]`; `GAP-ANALYSIS.md` Rev 13 (US-2.1 ✅; counts Met 13 / Partial 4 / Not met 2); AGENTS.md / README.md one-line note.

## Test matrix

| Case | Expected |
| --- | --- |
| bundled template detection | `header_footer.has_header=False`, `has_footer=True` |
| synthetic chrome-less master | both false; `needs_header_footer_prompt` True |
| inject default header | `header_footer.header` = 4-point polygon + English note; explicit assertions (len==4, {x,y}, [0,1]); `validate_template_schema` passes |
| master-only scope | layout-only header/footer not double-counted (master is canonical) |

## Verification

```powershell
python -m pytest .opencode/skills/generate-slide-skill/scripts/tests/ -q
```

## Risks

- **Layout-only chrome**: if a template's header/footer live only on layouts (not the master), master-based detection reports false-negative. The bundled master has them; layout-only cases are out of scope (documented).
- **pptx-subagent light-note scope (arch-review M1)**: `read_embedded_schema` returns `None` for non-templated inputs at Stage 0 (US-4.3 produces the schema post-render). The note is therefore scoped to **templated inputs only**. Do NOT add a redundant `extract_schema` at Stage 0 — it conflicts with US-4.3's "extract once inline" design and doubles cost. The full prompt+inject path lives in `generate-template-skill` (which always has the extracted schema).
- **Prompt is agent behaviour**: AC2's automated test covers the `needs_header_footer_prompt` helper + SKILL.md instructions; the live conversational prompt is verified by SKILL.md reasoning (headless skips).
- **`header_footer` shape**: the schema is `additionalProperties:true`, so any shape validates; keep has_header/has_footer + the optional injected zone (polygon pinned to US-1.2 4-point model).

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` US-2.1 (L144-158); GAP-ANALYSIS §US-2.1 (L95-97) + remediation (L202).
- Key code: `schema_extractor._build_metadata` (L485-497, the `header_footer:{}` placeholder), `map_placeholder_type`/chrome classification (L228-246).
- PLAN format template: `PLANS/PLAN-GIT-58.md` / `PLAN-GIT-60.md`.
