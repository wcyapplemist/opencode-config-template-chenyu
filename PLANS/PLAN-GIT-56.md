# PLAN-GIT-56 — US-3.1: generate-template-skill (Standalone Template Generator)

**Issue**: #56
**Branch**: GIT-56 (base: dev)
**Priority**: Must Have (P1)
**Status**: Implemented (all phases complete; 112 schema_extractor tests green. The 26 failures in the broader suite are pre-existing in `template_introspector`/`resource_pipeline`, confirmed via `git stash` — unrelated to US-3.1.)

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
2. **NL intent routing via SKILL.md `description` + agent deferral (MAJOR-1)** — the new skill's `description` **leads with extraction verbs** (`extract template`, `generate template`, `templated PPTX`, `template schema`, `what layouts`) and explicitly excludes slide generation. **Routing collision resolved:** `pptx-subagent.md` triggers greedily on `pptx` / `.pptx` / `presentation` (`pptx-subagent.md:75-79`); a request like "extract the template from this pptx" matches it first, and per the workspace "agents-first" delegation rule it would win over a skill — but the subagent is **forbidden** from extracting templates, so AC1 would silently fail. **Therefore `pptx-subagent.md` IS modified**: add one line to its `## What NOT to Handle` section routing template-extraction / "generate template" requests to `generate-template-skill`. (This **reverses** the earlier "pptx-subagent.md not modified" stance flagged by architecture review MAJOR-1.) Invoked directly by the primary conversation agent via the Skill tool.
3. **Skill wraps the existing engine; no rewrites** — calls `extract_schema` / `embed_schema` / `build_extraction_summary` via `python -c` (same pattern as `ppt-template-filler` calling `ppt_builder`). Function-level calls (not single-shot CLI) where a mid-pipeline pause (title confirm) is needed.
4. **Two additive engine enhancements** (pure, non-breaking):
   - (a) **`title_source` via NamedTuple (MINOR-5)** — new `TitleInference(NamedTuple)` with `title: str, source: str`; `_infer_title` returns it (mirrors the existing `EmbeddedSchemaResult(NamedTuple)` pattern at `schema_extractor.py:70`). Module constant `TITLE_SOURCES = frozenset({"core_xml", "slide1", "filename", "user"})` is the **single source of truth** consumed by `_infer_title`, the validator (Decision #5), and the spec enum. `_build_metadata` emits `title_source` (`source ∈ {"core_xml","slide1","filename"}` from inference). The skill uses the `question` tool when `source == "filename"`, then writes back `title` and sets `title_source = "user"`.
   - (b) **`build_extraction_summary(schema) -> str`**: pure function producing a multi-line human-readable summary (title + source, slide dimensions, layout count + names, master/layout component counts, theme colors + font palette, `missing_fonts` count + names). CLI gains `--summary` (prints to stdout after extraction).
5. **Spec sync + validator enforcement (MAJOR-2)** — add optional `title_source` (enum of 4: `core_xml | slide1 | filename | user`) to `template_metadata.properties` in `schemas/template_schema.json` (`template_metadata` declares `additionalProperties: false`, so the spec must track the output). **Critically, `title_source` IS enforced at runtime** — unlike the general case: `validate_template_schema` gains a ~3-line enum check keyed off the shared `TITLE_SOURCES` constant (reject any value not in the set). This **closes** the US-5.2 gap for this field rather than widening it (Task 13 now tests **both** directions). Other `additionalProperties` remain advisory (US-5.2 stays 🟡 Partial overall).
6. **Output paths** — templated PPTX → `output/<input_stem>.templated.pptx`; schema JSON → `output/<input_stem>.schema.json` (matches the existing `output/` convention).
7. **Error reporting (AC3)** — `extract_schema` raises `TemplateExtractionError` (incl. "no slide master found"); the skill catches it and **restates it structurally** to the user (file problem / validation failure / runtime), each with an actionable fix. Reuses the CLI exit-code semantics (1 = validation, 2 = runtime).
8. **Headless safety** — when the skill runs as a subagent (no user channel), Stage 2 skips the `question` prompt and accepts the filename fallback (never hangs), mirroring the `pptx-subagent` headless convention.

## Deliverables

**New file**
- `.opencode/skills/generate-template-skill/SKILL.md` — skill definition (frontmatter + What I do + When to use me + Stage 0–4 workflow + error handling + output).

**Additive code changes** (all under `.opencode/skills/ppt-template-filler/scripts/`)
- `schema_extractor.py`:
  - `TitleInference(NamedTuple)` (`title`, `source`); `TITLE_SOURCES = frozenset({...})` constant; `_infer_title(prs, path) -> TitleInference` (adds source return).
  - `_build_metadata`: emit `title_source` (immediately after `title`).
  - `validate_template_schema`: add the `title_source` enum check keyed off `TITLE_SOURCES` (MAJOR-2) — reject values outside the set.
  - `extract_schema`: adapt the single call site (unpack `TitleInference`).
  - New `build_extraction_summary(schema) -> str` (pure, no side effects).
  - CLI `main()`: add `--summary` (store_true) → `print(build_extraction_summary(schema))` after extraction.
- `schemas/template_schema.json`: add `title_source` (`{"type":"string","enum":["core_xml","slide1","filename","user"]}`, optional) under `template_metadata.properties`.

**Agent routing change (MAJOR-1)**
- `.opencode/agents/pptx-subagent.md`: add one line to the `## What NOT to Handle` section routing template-extraction / "generate template" requests to `generate-template-skill`. (No change to the `task` allow-list; the primary agent invokes the skill directly.)

**Tests** (`.opencode/skills/ppt-template-filler/scripts/tests/test_schema_extractor.py`)
- `title_source` across all three inference branches (core_xml / slide1 / filename); NamedTuple unpacking.
- `validate_template_schema` **enforces** `title_source`: valid values pass; an invalid value (e.g. "garbage") → validation error (MAJOR-2, both directions).
- skill write-back sets `title_source == "user"`.
- `build_extraction_summary` content assertions (title, layout count, missing_fonts, theme colors present).
- CLI `--summary` smoke (stdout contains key summary lines).
- Regression: existing schema-extractor tests unaffected by the `TitleInference`/`title_source` change.

**Docs**
- `docs/user-stories/chenyu-user-stories.md` — check US-3.1 all 3 ACs; collateral US-3.2 AC2/AC3 and US-3.3 AC1/AC2.
- `docs/user-stories/GAP-ANALYSIS.md` (Revision 8): US-3.1 ❌→✅; US-3.2 🟡→✅; US-3.3 stale ❌→✅ (note the Rev-7 miss); counts **Met 7→10 / Partial 6→5 / Not met 5→3 / Differs 1**; **Epic 3 complete**.
- `AGENTS.md` (project root) — add `generate-template-skill` to the Project-Level Resources table + structure diagram; one-line responsibility note.

## Acceptance Criteria (US-3.1) — delivered

- [x] AC1 — Skill is invocable by natural-language intent detection (SKILL.md `description` triggers; no special command needed).
- [x] AC2 — Full pipeline runs end-to-end without manual intermediate steps.
- [x] AC3 — Validation errors (e.g. "no slide master found") are reported clearly to the user.

**Collateral Must-Have ACs (same workflow)**
- [x] US-3.2 AC2 — Title inference order: `core.xml` → slide 1 title → user prompt.
- [x] US-3.2 AC3 — The title is displayed to the user for confirmation after extraction.
- [x] US-3.3 AC1 — A downloadable templated PPTX (original + embedded JSON) is returned.
- [x] US-3.3 AC2 — A human-readable extraction summary is printed (layouts, components, fonts, theme).
- (US-3.3 AC3 — round-trip — already met.)

## Implementation Phases

### Phase 1: Engine enhancements (schema_extractor.py + spec + validator)
- [x] Task 1: `TitleInference(NamedTuple)` (`title: str`, `source: str`); `TITLE_SOURCES = frozenset({"core_xml","slide1","filename","user"})` constant; `_infer_title(prs, path) -> TitleInference` (MINOR-5).
- [x] Task 2: `_build_metadata` emits `title_source` (after `title`); `extract_schema` adapts the single call site.
- [x] Task 3: `build_extraction_summary(schema) -> str` (pure).
- [x] Task 4: `schemas/template_schema.json` adds `title_source` (enum of 4, optional).
- [x] Task 5: `validate_template_schema` adds the `title_source` enum check keyed off `TITLE_SOURCES` (MAJOR-2) — reject values outside the set.
- [x] Task 6: CLI `main()` adds `--summary`.

### Phase 2: Skill + agent routing
- [x] Task 7: SKILL.md frontmatter (name / description **leading with extraction verbs** / license / compatibility / metadata).
- [x] Task 8: "What I do" / "When to use me" (extraction-verb triggers) / "Do NOT use" (cross-reference filler + modifier).
- [x] Task 9: Stage 0–4 workflow (with `python -c` invocation templates + headless branch); explicitly call `validate_template_schema` between Stage 1 and Stage 2.
- [x] Task 10: Error-handling table + output-path section.
- [x] Task 11: `pptx-subagent.md` — add one line to `## What NOT to Handle` routing template-extraction requests to `generate-template-skill` (MAJOR-1).

### Phase 3: Tests
- [x] Task 12: `title_source` — three inference branches (core.xml / slide1 / filename); NamedTuple unpacking.
- [x] Task 13: `validate_template_schema` **enforces** `title_source` — valid values pass; invalid value → error (both directions, MAJOR-2).
- [x] Task 14: skill write-back sets `title_source == "user"`.
- [x] Task 15: `build_extraction_summary` content assertions.
- [x] Task 16: CLI `--summary` smoke test.
- [x] Task 17: regression — existing schema-extractor tests unaffected.

### Phase 4: Docs
- [x] Task 18: `chenyu-user-stories.md` AC checkboxes; `GAP-ANALYSIS.md` Revision 8 (re-grade + counts + Epic 3 complete); `AGENTS.md` resources table.

## Test matrix

| Case | Expected |
| --- | --- |
| `_infer_title`: core.xml has title | `TitleInference("…", "core_xml")` |
| `_infer_title`: core.xml empty, slide 1 has text | `TitleInference("…", "slide1")` |
| `_infer_title`: neither | `TitleInference(filename_stem, "filename")` |
| skill writes back a user title | metadata `title_source == "user"` |
| `validate_template_schema` + valid `title_source` | passes |
| `validate_template_schema` + invalid `title_source` (e.g. "garbage") | **fails** with an error (MAJOR-2) |
| `build_extraction_summary` | contains title, `N layouts`, missing_fonts count, theme colors |
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

- **No deeper `pptx-subagent` integration** — the agent receives only a one-line `## What NOT to Handle` deferral (MAJOR-1, in scope); it does **not** gain the skill in its `task` allow-list, and forward-delegation routing is a separate concern.
- **`generate-slides` skill (US-5.1's other half) is out** — US-5.1 remains 🟡 Partial (improved, not fully Met).
- **Renderer migration to `read_embedded_schema`** (deprecate sidecar) — deferred from US-1.5, separate issue.
- **Real-PowerPoint "no repair prompt"** — not auto-testable; proxy + manual, per US-1.5.

## Risks

- **Spec↔output consistency (MAJOR-2)** — `title_source` is added to `template_schema.json` (`additionalProperties: false`) **and** enforced by `validate_template_schema` (enum check keyed off `TITLE_SOURCES`); Task 4 + Task 5 + the bidirectional test (Task 13) mitigate. Other `additionalProperties` remain advisory (US-5.2 Partial).
- **NL trigger collision (MAJOR-1)** — `pptx-subagent.md` triggers greedily on `pptx` / `.pptx`; mitigated by (a) the new skill leading with extraction verbs and (b) a `## What NOT to Handle` deferral line in `pptx-subagent.md` (Task 11, in scope). Residual overlap with `ppt-template-filler` (both say "template") is disambiguated via trigger phrases + mutual "Do NOT use".
- **Headless hang** — Stage 2 explicitly skips the `question` prompt in subagent mode and falls back to filename.
- **Small regression surface** — `_infer_title` has a single internal call site; `--summary` is purely additive.

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 3 (US-3.1–3.4).
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 Epic 3, §4 P1, §5 Decision 1.
- Predecessor: US-1.5 (issue #55) — `embed_schema` / `read_embedded_schema` / CLI `--embed`.
- Skill templates: `.opencode/skills/ppt-template-filler/SKILL.md`, `.opencode/skills/template-modifier-skill/SKILL.md`.
- PLAN format template: `PLANS/PLAN-GIT-55.md`.
- Architecture review (rev 2 basis): incorporated MAJOR-1 (agent routing collision), MAJOR-2 (validator enforcement of `title_source`), MINOR-5 (`TitleInference` NamedTuple). Other review findings (MAJOR-3 write-only drift, MINOR-4/6/7 prose-testability) are noted but out of scope for this revision.
