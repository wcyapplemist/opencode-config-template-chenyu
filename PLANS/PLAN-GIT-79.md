# PLAN-GIT-79 — pptx-subagent prompt hardening + simplification + three-skill unification

**Issue**: #86 (parent) — sub-issues #88 (Phase 1), #87 (Phase 2), #90 (Phase 3), #89 (Phase 4)
**Branch**: GIT-79 (base: dev)
**Priority**: Medium
**Status**: Planned (Rev 2 — architecture-review fixes adopted: CRIT-1/2, MAJ-1~7, MIN-1~5).

## Goal

Harden the `pptx-subagent` prompt so the agent **never** asks a `question()` before producing the first `.pptx`, unify all three skills under explicit frontmatter permissions + a detection-based routing table, and slim the prompt from 450 → ~200 lines by relocating reference material to `generate-slide-skill/SKILL.md`.

This is a **prompt/docs-only change** (2 files). No engine code, no tests. The primitives it relies on already exist: `read_embedded_schema` (US-4.3 detection), `repair_if_needed` (US-4.8), the engine's `auto_template` (US-4.3 inline extraction), generate-first defaults + post-generation refinement `question` (GIT-76), and `servable_slide_types` layout detection.

## Strategic Context

The agent prompt (`pptx-subagent.md`, 450 lines) has three structural weaknesses this plan resolves:

1. **Premature-questioning risk.** GIT-76 made the first generation zero-prompt, but the prompt still carries ~4 paraphrased "generate-first" statements scattered across Stage 0/1/2. These restatements are soft — a model can misread any one of them as "sometimes OK to ask". There is **no single, top-of-prompt hard rule** that bans `question()` outright before the first output.
2. **Three-skill fragmentation.** The frontmatter `permission.task` allows only `generate-slide-skill`. There is no Skill Routing table and no explicit detection gate.
3. **Bloat.** ~125 lines of reference content plus duplicated tables live in the agent prompt but already (mostly) exist in `generate-slide-skill/SKILL.md`.

**Risk profile is low** because no behavioral primitive is new — this is a prompt restructure layered on already-shipped engine behavior.

## Architecture Decisions (locked — Rev 2: 10 decisions)

1. **First-generation questions are BANNED (D1).** Rule #1 of the new ABSOLUTE RULES block: NEVER call `question()` between the user's initial prompt and the first `.pptx` output. Honor user-stated preferences; auto-determine unstated parameters. This subsumes and hardens GIT-76's "generate-first" philosophy into one top-of-prompt absolute rule; the ~4 scattered paraphrases are deleted. *(MAJ-1: this is prompt-level enforcement, best-effort; human test T5 (≥3 runs) is the regression gate.)*
2. **All three skills are permitted (D2).** `permission.task` opens `generate-slide-skill`, `generate-template-skill`, and `template-modifier-skill` as `allow`. The `"*": deny` catch-all stays. *(MAJ-5: `template-modifier-skill` permission is opened per user intent, but the routing table explicitly prohibits direct task dispatch — bash import only.)*
3. **`template-modifier-skill` is invoked via bash import, not direct task dispatch (D3).** The routing table row says explicitly: *"do NOT dispatch template-modifier-skill as a task; use `resolve_and_clone` via bash import."* *(MAJ-5 mitigation.)*
4. **Non-templated `.pptx` → informational detection only (D4, Rev 2).** *(CRIT-1 fix: Path B DROPPED.)* When Stage -1 detects the input is not templated (`read_embedded_schema` returns None), the agent emits a one-line status message (*"No template found — extracting first, then generating slides..."*) and proceeds normally. The engine's `auto_template` (US-4.3, `generate_ppt_from_data(auto_template=True)`) handles extraction + embedding into the **output** inline during render — no separate `generate-template-skill` invocation. This preserves the locked US-4.3 "engine-inline" decision (GAP-ANALYSIS:136, PLAN-GIT-63 Q1).
5. ~~**Extraction is zero-prompt (D5).**~~ **DROPPED (Rev 2).** No longer needed — Path B is dropped; extraction is engine-inline, not a separate skill invocation.
6. **Prompt length target = ~200 lines (D6, Rev 2).** Down from 450. *(MAJ-2: decision-time content — Speaker Notes Style Guide structure + Self-Critique Rubric dimensions — stays condensed in the prompt; only full examples and reference tables move to SKILL.md.)*
7. **Missing master/theme is NOT in the routing table (D7).** A corrupt/masterless template is **engine-internal** — `repair_if_needed` runs automatically inside `generate_ppt_from_data`. The agent prompt does not route on it.
8. **Repair result is surfaced in Stage 5 (D8).** After render, Stage 5 reads the `<output>.render.json` sidecar; if `templating.repair` is present, the agent informs the user.
9. **Routing basis is code detection (D9, Rev 2).** The Skill Routing table has **2 rows** (Rev 2 — Path B row dropped): missing layouts → template-modifier-skill (bash import); all pass → generate-slide-skill. Non-templated detection is informational only (one-line status message; engine handles inline).
10. **Language strategy is split (D10).** Slide content = English ONLY (Rule #3). Agent↔user interaction = match the user's prompt language. *(MAJ-7: includes a concrete example — "Chinese prompt → Chinese status/outline/refinement text; slide titles+bodies stay English." For Stage 5 question: translate `header` + `question` text, keep option `label`s English (they map to engine params) with translated `description`.)*

## Deliverables

**Change** `.opencode/agents/pptx-subagent.md` (Phases 1–3):
- **Frontmatter** — add `generate-template-skill: allow` and `template-modifier-skill: allow`.
- **ABSOLUTE RULES block** (new, top of prompt body) — 5 rules + interaction-language note with example.
- **Skill Routing table** (2 rows, Rev 2) — missing layouts → template-modifier-skill (bash import, do NOT dispatch as task); all pass → generate-slide-skill. Note: non-templated detection is informational; engine handles inline.
- **Stage -1: Template Check** (new) — `read_embedded_schema` detection **wrapped in try/except** (CRIT-2); if not templated → one-line status message → proceed normally (engine auto_template handles embedding).
- **Stage 5 repair-report check** — read `templating.repair` from render.json sidecar; notify user.
- **Decision-time content stays condensed in prompt** (MAJ-2): Speaker Notes 4-part structure + word budget (~5 lines); Self-Critique 6 dimensions (~3 lines). Full examples move to SKILL.md.
- **Preserved one-liners** (MAJ-3): sign-off default ("first-gen: unset → engine removes placeholder") in Stage 5; density-soft ("warnings never block") in Stage 2; fingerprint ("layouts matched by composition, not name") in Stage 0.
- **"What NOT to Handle" edit** (MAJ-4): keep pure-extraction bullet (*"Pure extraction/fingerprint, no slides wanted → generate-template-skill"*); remove only the "rendering from non-templated is not my job" clause (it IS this agent's job, handled inline).
- **Delete** the 4 redundant "generate-first" paraphrases (~25 lines).
- **Remove** ~100 lines of reference content (full Speaker Notes example, full Example Interaction, Template Introspection bash snippets, duplicated tables) — relocated to SKILL.md.
- **Condense** each stage in-place.
- Final line count ~200 (±20).

**Change** `.opencode/skills/generate-slide-skill/SKILL.md` (Phase 4):
- Append **4 new sections**: `## Speaker Notes Style Guide` (full 4-part structure + GOOD example), `## Example Interaction`, `## Template Introspection Commands`, `## Self-Critique Rubric` (full 6-dimension detail).

## Acceptance Criteria

- [ ] AC1 — No `question()` is issued between the user's initial prompt and the first `.pptx` output (ABSOLUTE RULE #1, instructed + verified by multi-run human test T5). *(MAJ-1: "instructed + verified", not "enforced" — prompt-level is best-effort.)*
- [ ] AC2 — All three skills are permitted in frontmatter; routing is detection-based via the Skill Routing table (2 rows). `template-modifier-skill` is never dispatched as a `task` (explicit prohibition in routing table).
- [ ] AC3 (Rev 2) — A non-templated `.pptx` emits a one-line status message and renders normally; the engine's `auto_template` handles extraction inline. No separate `generate-template-skill` invocation for rendering.
- [ ] AC4 — The prompt is ~200 lines (±20); reference content relocated to `SKILL.md` and not duplicated. Decision-time content (notes structure, rubric dimensions) retained in condensed form.
- [ ] AC5 — Agent-user interaction language matches the user's prompt language; slide content is always English. Stage 5 question: `header`/`question` translated, option `label`s English.
- [ ] AC6 — Stage 5 surfaces a repair notice when the render.json sidecar reports `templating.repair`.
- [ ] AC7 — The 4 redundant "generate-first" paraphrases are removed; one ABSOLUTE RULE (#1) governs.
- [ ] AC8 (Rev 2, MAJ-3) — No behavioral rule from the old Absolute Constraints is lost: sign-off default, density-soft, and fingerprint resolution each survive as a one-liner in the appropriate stage.
- [ ] AC9 (Rev 2, CRIT-2) — Stage -1 wraps `read_embedded_schema` in try/except; corrupt/missing/non-PPTX paths never crash.

## Implementation Phases

> Dependency graph: `Phase 1 → Phase 2 → Phase 3 → Phase 4`

### Phase 1: Permissions + ABSOLUTE RULES + interaction language — #88
- [ ] T1.1: Frontmatter — add `generate-template-skill: allow` and `template-modifier-skill: allow`.
- [ ] T1.2: Add **ABSOLUTE RULES** block (5 rules) at the top of the prompt body, replacing `## Absolute Constraints`: Rule #1 (no `question()` before first output), Rule #2 (no building from scratch), Rule #3 (English-only slide content), Rule #4 (speaker notes mandatory), Rule #5 (validate before render). Add interaction-language note with concrete example (MAJ-7).
- [ ] T1.3: Delete the 4 redundant "generate-first" paraphrases in Stage 0 / Stage 1 / Stage 2 (~25 lines).
- [ ] T1.4 (MAJ-3): Verify each old Constraint (#3 fingerprint, #6 density-soft, #7 sign-off) has a surviving one-liner in the appropriate condensed stage.

### Phase 2: Skill routing + Stage -1 (informational) + repair report — #87
- [ ] T2.1: Add **Skill Routing table** (2 rows, Rev 2): missing layouts → template-modifier-skill via bash import (**do NOT dispatch as task** — MAJ-5); all pass → generate-slide-skill. Note: non-templated detection is informational (engine handles inline); missing master/theme is engine-internal.
- [ ] T2.2 (CRIT-2): Add **Stage -1: Template Check** — run `read_embedded_schema(template_path)` **wrapped in try/except TemplateExtractionError** (treat exception as NOT_TEMPLATED); if not templated → emit one-line status message → proceed normally. The engine's `auto_template` handles extraction + embedding in the output during render. *(Path B dropped — CRIT-1.)*
- [ ] T2.3: Add **Stage 5 repair-report check** (D8) — read `templating.repair` from the render.json sidecar; notify the user (in their prompt's language).
- [ ] T2.4: Add **interaction-language note to the Stage 5 question template** — translate `header` + `question` text; keep option `label`s English with translated `description` (MAJ-7).
- [ ] T2.5 (MAJ-4): Update **"What NOT to Handle"** — keep pure-extraction bullet (*"Pure extraction/fingerprint, no slides wanted → generate-template-skill"*); remove only the "rendering from non-templated is not my job" clause. Keep non-PPT deferrals (docx, PDF, spreadsheets).

### Phase 3: Simplify prompt (450 → ~200 lines) — #90
- [ ] T3.1: Remove ~100 lines of reference content from `pptx-subagent.md` (destination: Phase 4): full Speaker Notes Style Guide example (~20 lines), full Example Interaction (~29 lines), Template Introspection bash snippets (~23 lines), full Self-Critique Rubric detail (~8 lines), duplicated tables (~20 lines).
- [ ] T3.2 (MAJ-2): Keep **condensed decision-time content** in the prompt: Speaker Notes 4-part structure + word budget (~5 lines, no full example); Self-Critique 6 dimensions as one-liner list (~3 lines, no detail).
- [ ] T3.3: Condense each stage in-place — Stage 0: 35→15, Stage 1: 35→8, Stage 2: 39→12, Stage 3: 42→15, Stage 4: 50→25 (MIN-5: slightly relaxed), Stage 5: 36→20.
- [ ] T3.4: Verify final line count ~200 (±20); no behavioral rule lost (AC8); Stage -1 exception-safe (AC9).

### Phase 4: SKILL.md receives moved content — #89
- [ ] T4.1: Append `## Speaker Notes Style Guide` to `SKILL.md` (full four-part structure + GOOD example).
- [ ] T4.2: Append `## Example Interaction` to `SKILL.md`.
- [ ] T4.3: Append `## Template Introspection Commands` to `SKILL.md`.
- [ ] T4.4: Append `## Self-Critique Rubric` to `SKILL.md` (full 6-dimension detail).
- [ ] T4.5: Verify the agent prompt cross-references SKILL.md where content was moved — no dangling pointer.

## Risks

- **Over-condensation loses behavior (R1).** **Mitigation:** reference content is relocated, not deleted; decision-time content stays condensed in prompt (MAJ-2); each old rule has a surviving one-liner (MAJ-3/AC8).
- **Premature-question regression (R2).** **Mitigation:** Rule #1 is the first item in the first block; 4 paraphrases deleted (single source of truth). Human test T5 (≥3 runs) is the gate. *(MAJ-1: prompt-level enforcement is best-effort.)*
- ~~**Routing-table miskey (R3).**~~ **Simplified (Rev 2):** routing table is 2 rows (Path B dropped); `read_embedded_schema` is wrapped in try/except (CRIT-2); engine handles non-templated files inline.
- **Permission-opened-but-unused-skill drift (R4).** **Mitigation (MAJ-5):** routing table explicitly says "do NOT dispatch template-modifier-skill as a task; use bash import."
- **Interaction-language confusion (R5).** **Mitigation (MAJ-7):** Rule #3 (English-only slides) sits above the interaction-language note; concrete example provided; Stage 5 question labels kept English (engine params).

## Out of scope (explicit)

- **Engine code / tests** — untouched.
- **Path B (extract-then-render via generate-template-skill)** — *(CRIT-1: DROPPED.)* The engine's `auto_template` handles non-templated files inline. A separate extraction-then-render flow is not introduced.
- **`generate-template-skill` SKILL.md updates** — the stale no-master handling (MAJ-6) is noted but out of scope for this issue; tracked separately.
- **New skill creation** — no skills added or split.
- **`render.json` schema changes** — Stage 5 only reads existing fields.

## Testing strategy (human testing only — no Python changes)

| ID | Scenario | Expected | AC |
|----|----------|----------|----|
| T1 | Chinese prompt, no template | zero-prompt, Chinese interaction, English slides | AC1, AC5 |
| T2 | English prompt, templated template | skip Stage -1 silently, direct render | AC2 |
| T3 | English prompt, non-templated template | one-line status message, engine auto-extracts inline, normal render | AC3 |
| T4 | English prompt, masterless template | auto-repair + repair report in Stage 5 | AC6, AC9 |
| T5 | Multiple runs of T1 (≥3) | ABSOLUTE RULES prevents premature questions | AC1, AC7 |

## References

- Requirements source: issue #86 (parent) — 10 locked design decisions (Rev 2).
- Architecture review: Rev 2 incorporates CRIT-1 (Path B dropped), CRIT-2 (exception safety), MAJ-1~7, MIN-1~5.
- Sub-issues: #88 (Phase 1), #87 (Phase 2), #90 (Phase 3), #89 (Phase 4).
- Reused primitives (no change): `read_embedded_schema` (US-4.3); `auto_template` (US-4.3 inline extraction); `repair_if_needed` (US-4.8); generate-first defaults (GIT-76); `servable_slide_types`; `resolve_and_clone` (template-modifier-skill).
- PLAN format template: `PLANS/PLAN-GIT-76.md`, `PLANS/PLAN-GIT-78.md`.
