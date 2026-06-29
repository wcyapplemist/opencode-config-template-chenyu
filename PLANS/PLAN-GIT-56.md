# PLAN-GIT-56 — US-3.1: generate-template-skill (Standalone Template Generator)

**Issue**: #56
**Branch**: GIT-56 (base: dev)
**Priority**: Must Have (P1)
**Status**: Planned

## Goal

Bring US-3.1 from ❌ Not met to ✅ Met by shipping a **standalone `generate-template-skill`** invocable by natural-language intent detection (e.g. "extract the template from this PPTX" / "generate template"), which orchestrates the full pipeline **end-to-end** with no manual intermediate steps:

`extract → validate → (title confirm) → embed → return templated PPTX + human-readable summary`

The engine already exists (`schema_extractor.py` CLI + `embed_schema` / `read_embedded_schema`, delivered by Epic 1 / US-1.5). This plan adds only the **orchestration + interaction layer (the skill)** plus two small additive engine enhancements. As natural collateral the skill's workflow also satisfies **US-3.2 AC2/AC3** (title user-prompt + confirmation) and **US-3.3 AC1/AC2** (downloadable templated PPTX + extraction summary). US-3.3 AC3 (round-trip) already exists. On completion **all 4 Epic 3 stories are Met**.

## Strategic Context

This is the highest-priority Must-Have gap per `docs/user-stories/GAP-ANALYSIS.md` §4 P1: there is no standalone template-generator skill — introspection is embedded in the render path, and the CLI (`schema_extractor.py`) lacks the orchestration + interaction layer (NL routing, title confirmation, extraction summary, surfaced download). The engine capabilities are complete (US-1.1–1.5: extraction, normalized polygon, type confidence, fonts, zip-embed); the gap is purely the skill wrapper.

It also corrects the **GAP-ANALYSIS Rev-7 stale rating of US-3.3**: after US-1.5 landed `embed_schema` + the round-trip test, US-3.3 should have been re-graded (AC3 met; AC1 partially met via CLI `--embed --output-pptx`). This plan completes the remaining interaction-layer ACs.

Aligned with GAP-ANALYSIS §5 **Decision 1 = Coexist**: the skill wraps `schema_extractor`; the renderer's sidecar read path (`template_introspector.py`) is **unchanged**. The skill is a peer of `ppt-template-filler` (fill) and `template-modifier-skill` (extend) — orthogonal responsibilities, all three coexist.

## Architecture Decisions (locked)

1. **New standalone skill** — `.opencode/skills/generate-template-skill/SKILL.md`. Naming aligns with the user story and the US-5.1 `generate-template` / `generate-slides` decomposition intent. Clean separation from the fill and extend skills.
2. **NL intent routing via SKILL.md `description`** — trigger phrases include `extract template`, `generate template`, `templated PPTX`, `template from this pptx`, `what layouts`, `template schema`. Invoked **directly by the primary conversation agent via the Skill tool**. `pptx-subagent.md` (a slide generator) is **not modified** — it does not absorb template extraction.
3. **Skill wraps the existing engine; no rewrites** — calls `extract_schema` / `embed_schema` / `build_extraction_summary` via `python -c` (same pattern as `ppt-template-filler` calling `ppt_builder`). Function-level calls (not single-shot CLI) where a mid-pipeline pause (title confirm) is needed.
4. **Two additive engine enhancements** (pure, non-breaking):
   - (a) **`title_source`**: `_infer_title` returns `(title, source)`, `source ∈ {"core_xml","slide1","filename"}`; `_build_metadata` emits `title_source`. The skill uses the `question` tool when `source == "filename"`, then writes back `title` and sets `title_source = "user"`.
   - (b) **`build_extraction_summary(schema) -> str`**: pure function producing a multi-line human-readable summary (title + source, slide dimensions, layout count + names, master/layout component counts, theme colors + font palette, `missing_fonts` count + names). CLI gains `--summary` (prints to stdout after extraction).
5. **Spec sync** — add optional `title_source` (enum of 4: `core_xml | slide1 | filename | user`) to `template_metadata.properties` in `schemas/template_schema.json`. `template_metadata` declares `additionalProperties: false`, so the spec must track the output to avoid widening the US-5.2 divergence. The hand-rolled `validate_template_schema` does not enforce `additionalProperties` at runtime (safe).
6. **Output paths** — templated PPTX → `output/<input_stem>.templated.pptx`; schema JSON → `output/<input_stem>.schema.json` (matches the existing `output/` convention).
7. **Error reporting (AC3)** — `extract_schema` raises `TemplateExtractionError` (incl. "no slide master found"); the skill catches it and **restates it structurally** to the user (file problem / validation failure / runtime), each with an actionable fix. Reuses the CLI exit-code semantics (1 = validation, 2 = runtime).
8. **Headless safety** — when the skill runs as a subagent (no user channel), Stage 2 skips the `question` prompt and accepts the filename fallback (never hangs), mirroring the `pptx-subagent` headless convention.

## Deliverables

**New file**
- `.opencode/skills/generate-template-skill/SKILL.md` — skill definition (frontmatter + What I do + When to use me + Stage 0–4 workflow + error handling + output).

**Additive code changes** (all under `.opencode/skills/ppt-template-filler/scripts/`)
- `schema_extractor.py`:
  - `_infer_title(prs, path) -> Tuple[str, str]` (adds source return); constant for the source values.
  - `_build_metadata`: emit `title_source` (immediately after `title`).
  - `extract_schema`: adapt the single call site (unpack the tuple).
  - New `build_extraction_summary(schema) -> str` (pure, no side effects).
  - CLI `main()`: add `--summary` (store_true) → `print(build_extraction_summary(schema))` after extraction.
- `schemas/template_schema.json`: add `title_source` (`{"type":"string","enum":["core_xml","slide1","filename","user"]}`, optional) under `template_metadata.properties`.
- `tests/test_schema_extractor.py`:
  - `title_source` across all three inference branches; skill write-back sets `"user"`.
  - `build_extraction_summary` content assertions (title, layout count, missing_fonts, theme colors present).
  - CLI `--summary` smoke (stdout contains key summary lines).
  - Regression: schema with `title_source` still passes `validate_template_schema`.

**Docs**
- `docs/user-stories/chenyu-user-stories.md` — check US-3.1 all 3 ACs; collateral US-3.2 AC2/AC3 and US-3.3 AC1/AC2.
- `docs/user-stories/GAP-ANALYSIS.md` (Revision 8): US-3.1 ❌→✅; US-3.2 🟡→✅; US-3.3 stale ❌→✅ (note the Rev-7 miss); counts **Met 7→10 / Partial 6→5 / Not met 5→3 / Differs 1**; **Epic 3 complete**.
- `AGENTS.md` (project root) — add `generate-template-skill` to the Project-Level Resources table + structure diagram; one-line responsibility note.

## Acceptance Criteria (US-3.1) — to deliver

- [ ] AC1 — Skill is invocable by natural-language intent detection (SKILL.md `description` triggers; no special command needed).
- [ ] AC2 — Full pipeline runs end-to-end without manual intermediate steps.
- [ ] AC3 — Validation errors (e.g. "no slide master found") are reported clearly to the user.

**Collateral Must-Have ACs (same workflow)**
- [ ] US-3.2 AC2 — Title inference order: `core.xml` → slide 1 title → user prompt.
- [ ] US-3.2 AC3 — The title is displayed to the user for confirmation after extraction.
- [ ] US-3.3 AC1 — A downloadable templated PPTX (original + embedded JSON) is returned.
- [ ] US-3.3 AC2 — A human-readable extraction summary is printed (layouts, components, fonts, theme).
- (US-3.3 AC3 — round-trip — already met.)

## Implementation Phases

### Phase 1: Engine enhancements (schema_extractor.py + spec)
- [ ] Task 1: `_infer_title` returns `(title, source)`; source-value constant.
- [ ] Task 2: `_build_metadata` emits `title_source`; `extract_schema` adapts the single call site.
- [ ] Task 3: `build_extraction_summary(schema) -> str`.
- [ ] Task 4: `schemas/template_schema.json` adds `title_source` (enum of 4, optional).
- [ ] Task 5: CLI `main()` adds `--summary`.

### Phase 2: Skill (new SKILL.md)
- [ ] Task 6: Frontmatter (name / description with triggers / license / compatibility / metadata).
- [ ] Task 7: "What I do" / "When to use me" (trigger phrases) / "Do NOT use" (cross-reference filler + modifier).
- [ ] Task 8: Stage 0–4 workflow (with `python -c` invocation templates + headless branch).
- [ ] Task 9: Error-handling table + output-path section.

### Phase 3: Tests
- [ ] Task 10: `title_source` — three inference branches (core.xml / slide1 / filename).
- [ ] Task 11: skill write-back sets `title_source == "user"`.
- [ ] Task 12: `build_extraction_summary` content assertions.
- [ ] Task 13: regression — schema with `title_source` still validates.
- [ ] Task 14: CLI `--summary` smoke test.

### Phase 4: Docs
- [ ] Task 15: `chenyu-user-stories.md` AC checkboxes; `GAP-ANALYSIS.md` Revision 8 (re-grade + counts + Epic 3 complete); `AGENTS.md` resources table.

## Test matrix

| Case | Expected |
| --- | --- |
| `_infer_title`: core.xml has title | `("…", "core_xml")` |
| `_infer_title`: core.xml empty, slide 1 has text | `("…", "slide1")` |
| `_infer_title`: neither | `(filename_stem, "filename")` |
| skill writes back a user title | metadata `title_source == "user"` |
| `build_extraction_summary` | contains title, `N layouts`, missing_fonts count, theme colors |
| `template_metadata` with `title_source` | `validate_template_schema` still passes |
| CLI `--summary` | stdout contains key summary lines; exit 0 |
| existing round-trip | still green (no regression) |

## Verification

```powershell
# Engine unit tests
python -m pytest .opencode/skills/ppt-template-filler/scripts/tests/test_schema_extractor.py -q
python -m pytest .opencode/skills/ppt-template-filler/scripts/tests/ -q
# CLI self-sufficiency (US-5.1 AC2)
python .opencode/skills/ppt-template-filler/scripts/schema_extractor.py `
  --input .opencode/skills/ppt-template-filler/scripts/templates/template.pptx `
  --output output/templated.schema.json --embed --summary
# Skill end-to-end (primary agent via Skill tool: extract -> confirm title -> embed -> summary -> return path)
```

## Out of Scope / Open Questions

- **`pptx-subagent.md` unchanged** — `generate-template` is triggered directly by the primary agent; if subagent routing is later desired, a separate issue handles it.
- **`generate-slides` skill (US-5.1's other half) is out** — US-5.1 remains 🟡 Partial (improved, not fully Met).
- **Renderer migration to `read_embedded_schema`** (deprecate sidecar) — deferred from US-1.5, separate issue.
- **Real-PowerPoint "no repair prompt"** — not auto-testable; proxy + manual, per US-1.5.

## Risks

- **Spec↔output consistency** — `title_source` must be added to `template_schema.json` (`additionalProperties: false`); Task 4 + regression test mitigate.
- **NL trigger overlap with `ppt-template-filler`** — both mention "template"; disambiguate via trigger phrases (generate-template = extract/generate/templated; filler = populate/fill/slides/deck) and mutual "Do NOT use" cross-references.
- **Headless hang** — Stage 2 explicitly skips the `question` prompt in subagent mode and falls back to filename.
- **Small regression surface** — `_infer_title` has a single internal call site; `--summary` is purely additive.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 3 (US-3.1–3.4).
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 Epic 3, §4 P1, §5 Decision 1.
- Predecessor: US-1.5 (issue #55) — `embed_schema` / `read_embedded_schema` / CLI `--embed`.
- Skill templates: `.opencode/skills/ppt-template-filler/SKILL.md`, `.opencode/skills/template-modifier-skill/SKILL.md`.
- PLAN format template: `PLANS/PLAN-GIT-55.md`.
