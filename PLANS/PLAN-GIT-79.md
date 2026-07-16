# PLAN-GIT-79 — pptx-subagent prompt hardening + simplification + three-skill unification

**Issue**: #86 (parent) — sub-issues #88 (Phase 1), #87 (Phase 2), #90 (Phase 3), #89 (Phase 4)
**Branch**: GIT-79 (base: dev)
**Priority**: Medium
**Status**: Planned (prompt/docs-only change — 2 files, no Python code, no tests).

## Goal

Harden the `pptx-subagent` prompt so the agent **never** asks a `question()` before producing the first `.pptx`, unify all three skills under explicit frontmatter permissions + a detection-based routing table, and slim the prompt from 450 → ~180 lines by relocating ~125 lines of reference material to `generate-slide-skill/SKILL.md`.

This is a **prompt/docs-only change** (2 files). No engine code, no tests, no subagent autonomous-path behavioral change. The primitives it relies on already exist: `read_embedded_schema` (US-4.3 detection), `repair_if_needed` (US-4.8, PLAN-GIT-78), `_infer_title` (US-3.2 zero-prompt title), the generate-first defaults + post-generation refinement `question` (GIT-76), and `servable_slide_types` layout detection.

## Strategic Context

The agent prompt (`pptx-subagent.md`, 450 lines) has three structural weaknesses this plan resolves:

1. **Premature-questioning risk.** GIT-76 made the *first* generation zero-prompt, but the prompt still carries ~4 paraphrased "generate-first" statements scattered across Stage 0/1/2 (e.g. L97, L118, L198, L233). These restatements are soft — a model can misread any one of them as "sometimes OK to ask". There is **no single, top-of-prompt hard rule** that bans `question()` outright before the first output. A model that asks once invalidates the entire generate-first contract.
2. **Three-skill fragmentation.** The frontmatter `permission.task` allows only `generate-slide-skill` (`pptx-subagent.md:10-12`). `generate-template-skill` and `template-modifier-skill` are reachable only implicitly, and non-templated `.pptx` files are routed by **user-intent interpretation** (the agent guesses whether the user wants extraction vs. rendering) rather than by **code detection** (`read_embedded_schema` return value). This is brittle. There is no Skill Routing table and no Stage -1 detection gate.
3. **Bloat.** ~125 lines of reference content — Speaker Notes Style Guide (L380–404), Example Interaction (L406–435), Template Introspection Commands (the Stage 0 bash snippets), Self-Critique Rubric (Stage 2) — plus duplicated tables (resource placeholders, image presets, slide-type table, density table) live in the agent prompt but already (mostly) exist in `generate-slide-skill/SKILL.md`. Every line in the agent prompt is re-read on every turn; reference material that the agent consults *on demand* belongs in the skill, not the system prompt. Target: **~180 lines** (60% reduction).

**Risk profile is low** because no behavioral primitive is new — this is a prompt restructure (permissions + routing + relocation) layered on already-shipped engine behavior.

## Architecture Decisions (locked — 10 decisions)

1. **First-generation questions are BANNED (D1).** Rule #1 of the new ABSOLUTE RULES block: NEVER call `question()` between the user's initial prompt and the first `.pptx` output. Honor user-stated preferences; auto-determine unstated parameters. This subsumes and hardens GIT-76's "generate-first" philosophy into one top-of-prompt absolute rule; the ~4 scattered paraphrases are deleted.
2. **All three skills are permitted (D2).** `permission.task` opens `generate-slide-skill`, `generate-template-skill`, and `template-modifier-skill` as `allow`. The `"*": deny` catch-all stays.
3. **`template-modifier-skill` is invoked via bash import, not direct task dispatch (D3).** Permission is *opened* (so the agent is authorized), but the actual invocation path stays the existing bash `import` pattern (consistent with how the engine already calls into it). Opening the permission removes a latent denial; it does not change the call mechanism.
4. **Non-templated `.pptx` routing = Path B (D4).** When Stage -1 detects the input is not templated (`read_embedded_schema` returns None), the agent invokes `generate-template-skill` to extract → embed first, *then* renders via `generate-slide-skill`. This replaces the current intent-interpretation routing with deterministic code detection.
5. **Extraction is zero-prompt (D5).** `generate-template-skill`'s `_infer_title` auto-determines the title (US-3.2); the subagent has no user channel, so no title-confirmation prompt is issued. The templated file is produced silently and the pipeline continues.
6. **Prompt length target = ~180 lines (D6).** Down from 450. Achieved by relocating ~125 lines of reference content to `SKILL.md` (Phase 4) + removing duplicated tables + condensing each stage in-place.
7. **Missing master/theme is NOT in the routing table (D7).** A corrupt/masterless template is **engine-internal** — `repair_if_needed` (US-4.8, PLAN-GIT-78) runs automatically inside `generate_ppt_from_data` before `get_render_contract`. The agent prompt does not route on it; it is not a Skill Routing row.
8. **Repair result is surfaced in Stage 5 (D8).** After render, Stage 5 reads the `<output>.render.json` sidecar; if `templating.repair` is present (level L1/L2/L3), the agent informs the user that the template was auto-repaired. This is a *notification*, not a routing decision.
9. **Routing basis is code detection, not intent interpretation (D9).** The Skill Routing table keys off `read_embedded_schema`'s return value and `servable_slide_types`'s layout availability — never on guessing what the user "wants". Detection is deterministic; intent is not.
10. **Language strategy is split (D10).** Slide content = English ONLY (unchanged, Rule #3). Agent↔user interaction = match the user's prompt language (Chinese prompt → Chinese conversation; the slides are still English). This is added as an explicit instruction (Stage 0 + Stage 5 question template).

## Deliverables

**Change** `.opencode/agents/pptx-subagent.md` (Phases 1–3):
- **Frontmatter** (`permission.task`) — add `generate-template-skill: allow` and `template-modifier-skill: allow` (keep `"*": deny`).
- **ABSOLUTE RULES block** (new, top of prompt body) — replaces `## Absolute Constraints` (L82). Five rules (D1): (1) no `question()` before first output; (2) no building from scratch; (3) English-only slide content; (4) speaker notes mandatory; (5) validate before render.
- **Interaction-language instruction** (D10) — communicate with the user in their prompt's language; slides always English. Wired into Stage 0 and the Stage 5 question template.
- **Skill Routing table** (D9, 3 rows) — not-templated → generate-template-skill; missing layouts → template-modifier-skill (bash import); all pass → generate-slide-skill.
- **Stage -1: Template Check** (new, D4/D5) — `read_embedded_schema` detection; Path B if not templated (zero-prompt extract).
- **Stage 5 repair-report check** (D7/D8) — read `templating.repair` from render.json sidecar; notify user.
- **"What NOT to Handle" edit** — remove the template-extraction routing bullet (now internal to the agent).
- **Delete** the 4 redundant "generate-first" paraphrases in Stage 0/1/2 (~25 lines) — consolidated into Rule #1.
- **Remove** ~125 lines of reference content (Speaker Notes Style Guide, Example Interaction, Template Introspection Commands, Self-Critique Rubric) + duplicated tables (resource placeholders, image presets, slide-type, density) — relocated to `SKILL.md` (Phase 4).
- **Condense** each stage in-place (Stage 0: 35→15, Stage 1: 35→8, Stage 2: 39→12, Stage 3: 42→15, Stage 4: 50→20, Stage 5: 36→20).
- Final line count ~180 (±15).

**Change** `.opencode/skills/generate-slide-skill/SKILL.md` (Phase 4):
- Append **4 new sections** receiving the relocated content: `## Speaker Notes Style Guide`, `## Example Interaction`, `## Template Introspection Commands`, `## Self-Critique Rubric`.

## Acceptance Criteria

- [ ] AC1 — No `question()` is ever issued between the user's initial prompt and the first `.pptx` output (ABSOLUTE RULE #1, enforced).
- [ ] AC2 — All three skills are permitted in frontmatter; routing is detection-based (`read_embedded_schema`) via the Skill Routing table + Stage -1 gate.
- [ ] AC3 — A non-templated `.pptx` triggers Path B (extract via `generate-template-skill`, then render); extraction is zero-prompt (D5).
- [ ] AC4 — The prompt is ~180 lines (±15); reference content relocated to `SKILL.md` and not duplicated.
- [ ] AC5 — Agent-user interaction language matches the user's prompt language; slide content is always English (D10).
- [ ] AC6 — Stage 5 surfaces a repair notice when the render.json sidecar reports `templating.repair` (D8).
- [ ] AC7 — The 4 redundant "generate-first" paraphrases are removed; one ABSOLUTE RULE (#1) governs (D1).

## Implementation Phases

> Dependency graph: `Phase 1 → Phase 2 → Phase 3 → Phase 4`

### Phase 1: Permissions + ABSOLUTE RULES + interaction language — #88
- [ ] T1.1: Frontmatter — add `generate-template-skill: allow` and `template-modifier-skill: allow` to `permission.task` (D2/D3 — `template-modifier-skill` invoked via bash import, not direct task dispatch).
- [ ] T1.2: Add **ABSOLUTE RULES** block (5 rules) at the top of the prompt body, replacing `## Absolute Constraints`: Rule #1 (no `question()` before first output — D1), Rule #2 (no building from scratch), Rule #3 (English-only slide content), Rule #4 (speaker notes mandatory ~120–180 words), Rule #5 (validate before render).
- [ ] T1.3: Add **interaction-language instruction** (D10) — communicate with the user in their prompt's language; slides always English. Wire into Stage 0 + Stage 5.
- [ ] T1.4: Delete the 4 redundant "generate-first" paraphrases in Stage 0 / Stage 1 / Stage 2 (~25 lines). Consolidated into Rule #1; no information lost.

### Phase 2: Skill routing + Stage -1 + repair report — #87
- [ ] T2.1: Add **Skill Routing table** (3 rows, detection-based — D9): not-templated → `generate-template-skill`; missing layouts → `template-modifier-skill` (bash import, D3); all pass → `generate-slide-skill`.
- [ ] T2.2: Add **Stage -1: Template Check** (D4/D5) — `read_embedded_schema` detection; Path B (extract via `generate-template-skill`, zero-prompt via `_infer_title`) if not templated.
- [ ] T2.3: Add **Stage 5 repair-report check** (D7/D8) — read `templating.repair` from the render.json sidecar; notify the user. Missing master/theme is engine-internal (`repair_if_needed`), NOT a routing row (D7).
- [ ] T2.4: Add **interaction-language note to the Stage 5 question template** (D10) — refinement `question` rendered in the user's prompt language.
- [ ] T2.5: Update **"What NOT to Handle"** — remove the template-extraction routing bullet (now handled internally by Stage -1 / the routing table); keep non-PPT deferrals (docx, PDF, spreadsheets).

### Phase 3: Simplify prompt (450 → ~180 lines) — #90
- [ ] T3.1: Remove ~125 lines of reference content from `pptx-subagent.md` (destination append is Phase 4): Speaker Notes Style Guide (~24 lines, L380–404), Example Interaction (~29 lines, L406–435), Template Introspection Commands (~23 lines, Stage 0 bash snippets), Self-Critique Rubric (~11 lines, Stage 2).
- [ ] T3.2: Remove duplicated tables already in `SKILL.md` (resource placeholders, image presets, slide-type table, density table).
- [ ] T3.3: Condense each stage in-place — Stage 0: 35→15, Stage 1: 35→8, Stage 2: 39→12, Stage 3: 42→15, Stage 4: 50→20, Stage 5: 36→20.
- [ ] T3.4: Verify final line count ~180 (±15); prompt keeps frontmatter, ABSOLUTE RULES, Skill Routing table, condensed Stages -1–5, "What NOT to Handle", Error Handling.

### Phase 4: SKILL.md receives moved content — #89
- [ ] T4.1: Append `## Speaker Notes Style Guide` to `SKILL.md` (four-part KEY MESSAGE / dialogue / TRANSITION / COACHING structure + GOOD example, verbatim from the agent prompt).
- [ ] T4.2: Append `## Example Interaction` to `SKILL.md` (AI-accounting worked example: autonomous standard + self-critique + post-generation refinement).
- [ ] T4.3: Append `## Template Introspection Commands` to `SKILL.md` (Stage 0 bash snippets: `servable_slide_types`, `get_render_contract`, `read_embedded_schema`).
- [ ] T4.4: Append `## Self-Critique Rubric` to `SKILL.md` (6 dimensions: consistency / flow / coverage gaps / redundancy / length / template fit).
- [ ] T4.5: Verify the agent prompt cross-references `SKILL.md` where these were removed — no dangling pointer.

## Risks

- **Over-condensation loses behavior (R1).** Aggressively shrinking stages (e.g. Stage 4 50→20 lines) risks dropping an instruction the model relied on. **Mitigation:** reference content is *relocated* to `SKILL.md` (Phase 4), not deleted; the agent prompt keeps a one-line cross-reference per moved section. Each stage's acceptance is "no behavioral information lost — recoverable from SKILL.md".
- **Premature-question regression (R2).** If the ABSOLUTE RULES block is placed too low or Rule #1 is diluted, the model may still ask before the first output. **Mitigation:** Rule #1 is the *first* item in the *first* block of the prompt body; the 4 redundant paraphrases are deleted so there is exactly one statement of the rule (single source of truth). Human test T5 (multiple runs) is the regression gate.
- **Routing-table miskey (R3).** A detection predicate that fires wrongly (e.g. treating a corrupt-but-templated file as "not templated") could loop extract→render unnecessarily. **Mitigation:** routing keys off `read_embedded_schema` (exception-safe, US-4.3) and `servable_slide_types` (proven in PLAN-GIT-78) — both deterministic and already battle-tested.
- **Permission-opened-but-unused-skill drift (R4).** Opening `template-modifier-skill` permission could invite the model to dispatch it as a task when bash import is the intended path (D3). **Mitigation:** the routing table explicitly says "via bash import" for that row; the skill's own SKILL.md documents Capability B. No code coupling changes.
- **Interaction-language confusion (R5).** A model could over-apply D10 and produce Chinese *slide* content. **Mitigation:** Rule #3 (English-only slide content) is an ABSOLUTE RULE that sits above the interaction-language note; the note explicitly scopes itself to agent↔user conversation, not slide content.

## Out of scope (explicit)

- **Engine code / tests** — `ppt_builder.py`, `schema_extractor.py`, `repair_if_needed`, `_infer_title`, `servable_slide_types`, `get_render_contract` are all untouched. This plan only restructures how the agent *drives* already-shipped behavior.
- **New skill creation** — no skills are added or split; the three existing skills are unified under the agent's permissions/routing.
- **Headless-subagent autonomous path** — unchanged (still zero-prompt first generation; per GIT-76 it skips the Stage 5 refinement question entirely).
- **`render.json` schema changes** — Stage 5 only *reads* `templating.repair` (added in PLAN-GIT-78); no new sidecar field is introduced.
- **Historical/analysis doc wording** (README, GAP-ANALYSIS, DESIGN-*) — not updated this issue.

## Testing strategy (human testing only — no Python changes)

| ID | Scenario | Expected | AC |
|----|----------|----------|----|
| T1 | Chinese prompt, no template | zero-prompt, Chinese interaction, English slides | AC1, AC3, AC5 |
| T2 | English prompt, templated template | skip Stage -1, direct render | AC2 |
| T3 | English prompt, non-templated template | Path B (extract then render), zero-prompt | AC2, AC3 |
| T4 | English prompt, masterless template | auto-repair + repair report in Stage 5 | AC6 |
| T5 | Multiple runs of T1 | ABSOLUTE RULES prevents premature questions | AC1, AC7 |

No `pytest` regression is applicable (zero Python changes). Verification is by running the primary-agent conversation flow on each scenario above and confirming the expected interaction shape.

## References

- Requirements source: issue #86 (parent) — 10 locked design decisions + 4-phase dependency graph.
- Sub-issues: #88 (Phase 1), #87 (Phase 2), #90 (Phase 3), #89 (Phase 4).
- Reused primitives (no change): `read_embedded_schema` (US-4.3 detection); `repair_if_needed` (US-4.8, PLAN-GIT-78); `_infer_title` (US-3.2 zero-prompt title); generate-first defaults + post-generation refinement `question` (GIT-76, PLAN-GIT-76); `servable_slide_types` layout detection; `get_render_contract`.
- Current prompt structure (verified): `pptx-subagent.md` (450 lines) — frontmatter `permission.task` (L6–12, only `generate-slide-skill` allowed), `## Absolute Constraints` (L82–100), Stage 0–5 (L124–376), Speaker Notes Style Guide (L380–404), Example Interaction (L406–435), What NOT to Handle (L437–444).
- PLAN format template: `PLANS/PLAN-GIT-76.md`, `PLANS/PLAN-GIT-78.md`.
