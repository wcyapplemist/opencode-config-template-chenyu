# PLAN-GIT-76 — pptx-subagent: generate-first workflow (drop pre-generation checkpoint, offer post-generation refinements)

**Issue**: #76
**Branch**: GIT-76 (base: dev)
**Priority**: Medium
**Status**: Implemented (Phase 0–2 done; T9 manual verification pending user) — architecture-review **CHANGES-REQUIRED** incorporated (rev 2) + code-quality review (rev 3): MAJOR-1 fixed (Constraint #7 sign-off contradiction — enforcement points 7→8, AC6 reworded), MINOR-2 fixed (Stage 3 placeholder de-quoted), MINOR-4 fixed (Stage 5 sign-off inline ask). rev 2 items: MAJOR-1 (template-aware density default + 6th rubric dimension), MAJOR-2 (enforcement inventory 4→7), MAJOR-3 (density-intent detection word list; AC3 made executable), MINOR-1/2/4/6 + omission fixes.

## Goal

Reverse the `pptx-subagent` primary-agent workflow philosophy from **confirm-first, generate-after** to **generate-first, refine-after**.

Today the primary agent forces a single `question` call **before** the first render (Stage 2) asking three things — density mode, outline approval, closing-slide sign-off — and the user must answer all three before any file is produced. After this change, the **first** generation runs **zero-prompt with safe defaults** straight to render; the user receives a file immediately; then exactly **one** multi-select `question` offers optional refinements, applied in a **single** re-generation round (no loop).

This is a **prompt/docs-only change** (4 files). No engine code, no tests, no subagent path, no multi-stage pipeline change.

## Strategic Context

The "confirm-first" flow is enforced in **8 functional places** (rev 3 — code-quality review found Constraint #7 was a 8th point missed by rev 2's "7 places") that must all be updated consistently, plus several historical/analysis docs (lower priority — see Out-of-scope):

**Functional enforcement points (all must be updated — AC6):**
1. **`AGENTS.md:98`** — project-level MANDATORY checkpoint paragraph ("MUST pause... single `question` call... wait for all answers before proceeding to Stage 3").
2. **`.opencode/agents/pptx-subagent.md:201`** (Stage 2, L183–242) — the primary agent's three-question interactive branch (density / outline approval / closing sign-off), distinct from the subagent autonomous self-critique path at L235–242.
3. **`.opencode/agents/pptx-subagent.md:97`** (Constraint #7) — "Primary agent: **proactively ask** for the presenter's name and email" (a pre-generation sign-off ask). *(rev 3 — code-quality review MAJOR-1: this directly contradicts the new "no pre-generation question" rule and the rewritten Stage 2; both the implementation and the architecture review missed it. Fixed by deferring sign-off to Stage 5.)*
4. **`.opencode/agents/pptx-subagent.md:117`** — the Generation Pipeline ASCII overview line: "`(primary-agent: pick mode + approve outline in ONE question call)`". The agent reads this overview first on entering the prompt — it must not contradict the rewritten Stage 2 body.
5. **`.opencode/agents/pptx-subagent.md:129`** — multi-aspect-ratio: "When unsure which preset, **ask in the Stage 2 checkpoint (batch with the density/sign-off questions)**". The Stage 2 checkpoint no longer exists in the new flow — this must be redirected (default to native; surface as a post-generation refinement instead).
6. **`generate-slide-skill/SKILL.md:317`** — "the agent picks the mode with the user at the outline-confirmation checkpoint".
7. **`generate-slide-skill/SKILL.md:343`** — "When run as the **primary** agent, it **can pause after the outline for the user to approve/edit**". Same file as #6, different paragraph; leaving it stale would override #6's new wording when the skill file loads.
8. **`density_mode.py:7`** (module docstring) — "a mode (with the user, at the outline-confirmation checkpoint)".

**Historical/analysis docs (lower priority — disposition in Out-of-scope, not AC6-blocking):** `README.md:149` ("approve or self-critique the outline"); `docs/user-stories/GAP-ANALYSIS.md:139` ("the outline is shown for user approval"); `generate-slide-skill/docs/DESIGN-multi-stage-generation.md:37` ("the interactive checkpoint now does two things in one `question` call"); `outline_store.py:8` + `:15` (docstring: "interactive-checkpoint flow (#24)" / "checkpoint confirms it").

The **infrastructure to support this already exists** — nothing new is built:

- The subagent (headless) autonomous path already runs `standard` + self-critique (rubric: consistency / flow / coverage gaps / redundancy / length) at `pptx-subagent.md:235–242`. This plan simply **promotes that rubric to the primary agent's first generation too**.
- The density mechanism (`density_mode.py`, `schema_validator`) is unchanged — density is still a per-slide word budget; the only change is *who picks it and when* (now: agent default first; user-refined later).
- US-4.6 `target_size` + US-4.3 `auto_template` + the closing-placeholder-removal behavior (engine drops `presenter_name`/`presenter_email` when unset) all already work — the post-generation refinement options just re-invoke `generate_ppt_from_data` with different params.
- `save_outline` artifact persists in both paths; the only change is the display semantics (information-only vs. approval-wait).

**Risk profile is low** because every behavioral primitive already exists; this is purely a re-sequencing of when defaults vs. user input are applied.

## Architecture Decisions (locked)

1. **generate-first-then-refine philosophy** — the primary agent's first generation runs **zero-prompt** to render using safe defaults, returns the file, and only *then* offers optional refinements via one multi-select `question`. The pre-generation `question` (Stage 2's three-question branch) is removed for the primary-agent path.
2. **Defaults cover ONLY unstated parameters (core principle)** — if the user's first message already states a preference (page count, density, aspect ratio), that preference wins on the first generation. Defaults fill gaps; they never override an explicit user statement. "Zero-prompt" = "no *active* prompting", never "ignore what the user already said".
3. **First-generation defaults (rev 2 — MAJOR-1: density is template-aware, not blindly `standard`):**
   | Parameter | Default | Rationale |
   |---|---|---|
   | Density mode | **`standard` baseline; AUTO-DOWNSHIFT to `concise` if Stage 0 reports `content_area < ~30 in²`** | honors the MANDATORY "Template-aware content" rule (pptx-subagent.md:79 / Stage 0:127) — small-content-area templates must not render at standard density or text overflows. This downshift is part of the **default logic**, NOT a user preference (it applies even when the user said nothing). |
   | Outline | auto-generate + **self-critique** (reuse subagent rubric + new 6th dimension — Decision 11) | no longer requires user approval |
   | Closing sign-off | **none** (unset `presenter_name`/`presenter_email` → engine removes placeholder) | avoids default `Prepared by: Lecturer Name` bleed |
   | Slide count | user-stated count if given (N incl. cover+closing); else natural outline + closing by default | reuses slide-count convention |
   | Aspect ratio | native (omit `target_size`) | unless user explicitly asks 4:3 / 1:1 |
4. **Outline shown but not confirmed** — the first generation displays the outline as pure information ("Here's the outline I'll generate — proceeding with defaults") and continues without waiting for approval. The `save_outline` artifact is still written (traceability for Stage 3's density header).
5. **Post-generation refinements = one multi-select `question` (`multiple: true`)** with the **full set of 7 options** (Decision 6). Issued in Stage 5 *after* the file path is returned.
6. **The 7 refinement options and their re-generation mechanisms:**
   | Option | Effect | Re-gen mechanism |
   |---|---|---|
   | Lower text density | `standard` → `concise` (0–10 w/slide) | rewrite JSON (new budget) → re-validate → re-render |
   | Increase text density | `standard` → `text-heavy` (75–150 w/slide) | rewrite JSON → re-validate → re-render |
   | Reduce slide count | merge / cut slides | revise outline → rewrite → re-render |
   | Add / split slides | split overcrowded content or add a section | revise outline → rewrite → re-render |
   | Add presenter sign-off | closing gets `presenter_name` + `presenter_email` | re-render only |
   | Change aspect ratio | 4:3 or 1:1 | `target_size` only → re-render |
   | No adjustment (recommended) | keep current result | — |
   "No adjustment" (or no selection) ends the workflow.
7. **One round only** — after the user picks refinements and a second file is returned, the workflow ends. There is **no second refinement prompt**, no loop. If the user wants further changes they issue a fresh request.
8. **Subagent (headless) path unchanged** — it was already autonomous (`standard` + self-critique, no sign-off). This plan does not touch the subagent branch; the self-critique rubric is *promoted* to the primary path, not *moved* out of the subagent path.
9. **Multi-stage pipeline preserved** — outline → critique → detail still runs for both agents; the primary agent's first generation now uses *self-critique* instead of *interactive approval* for the critique stage.
10. **No engine code change, no tests** — `ppt_builder.py` / `density_mode.py` logic / `schema_validator` are untouched. This is a prompt/docs-only change to 7 files. Density validation, `target_size`, `auto_template`, and placeholder removal all behave exactly as before; the agent just *drives* them differently.
11. **Template-aware density is wired into the self-critique rubric (rev 2 — MAJOR-1).** The promoted self-critique rubric gains a **6th dimension — "Template fit"**: *"If Stage 0 reported a small content area (< ~30 in²), density has been downshifted to `concise`; is the body still within the concise budget?"* This makes the Stage 0 → Stage 2 density data flow (pptx-subagent.md:127) explicit in the autonomous path, so the downshift is not lost when the interactive "density choice" is removed. The rubric now reads: consistency / flow / coverage gaps / redundancy / length / **template fit**.
12. **Density-intent detection makes AC3 executable (rev 2 — MAJOR-3).** Real users rarely say literal mode names; they say "brief / detailed / 精简 / 讲义式". Stage 0 gains a **density-intent word list** that maps natural-language cues to a mode, treating a detected cue as an explicit user preference (defaults do not override it):
    - `concise` ← "简要 / 精简 / 概览 / quick / brief / minimal / overview / keynote"
    - `text-heavy` ← "详细 / 深入 / 讲义 / 详尽 / detailed / thorough / in-depth / handout / dense"
    - `standard` ← no density word (baseline; then MAJOR-1's template-aware downshift may still apply)
    This turns AC3 from a principle into an executable rule (a test like "做个简要的 5 页 PPT" → first-gen `density_mode='concise'` is now assertable).
13. **Stage 5 refinement `question` is primary-agent-only (rev 2 — MINOR-1 + omission 3).** The primary/subagent distinction **moves wholesale to Stage 5**: the primary agent issues the post-generation refinement `question`; the **subagent (headless) skips it entirely** — first generation returns and the workflow ends, no question at all. Stage 2 has no primary/subagent branch left (both run the same autonomous self-critique).

## Deliverables

**Change** `.opencode/agents/pptx-subagent.md` (core):
- Stage 2 (L183–242): rewrite — delete the primary-agent three-question branch (L201–233); promote the subagent self-critique rubric (L235–242) **plus the new 6th "Template fit" dimension (Decision 11)** to the primary-agent first-generation path; **collapse the primary/subagent Stage-2 branches into one autonomous path** (the distinction moves to Stage 5 per Decision 13); retitle to "Density Mode + Self-Critique (autonomous, no pre-generation prompt)".
- Stage 1 outline: keep `save_outline` artifact (update its purpose text L161 from "foundation for Stage 2's interactive branch" → "traceability artifact for Stage 3's density header"); change display to pure information ("proceeding with defaults — you can adjust count/density in the next step", no confirmation wait).
- Stage 5: extend the return path — after the file path is returned, the **primary agent** issues the post-generation refinement `question` (7 options, `multiple: true`) with per-option re-gen mechanism notes + a **multi-select conflict-resolution rule** (Decision 6/MINOR-2: mutually-exclusive picks like "lower"+"increase density" resolve to the last-selected, or a one-line inline confirm); **subagent skips this entirely** (Decision 13). One round only.
- Stage 0: make the defaults principle explicit (user-stated preferences win; defaults only cover the unstated) **+ add the density-intent word list (Decision 12) + the template-aware downshift note (Decision 3/11: small-content-area → concise is default logic, not user pref)**.
- **L116–117 pipeline overview**: update the ASCII line — remove "`(primary-agent: pick mode + approve outline in ONE question call)`", replace with "`(autonomous: standard-or-template-aware density + self-critique, no pre-gen prompt; refinements offered post-generation by primary agent)`".
- **L129 multi-aspect-ratio**: redirect "ask in the Stage 2 checkpoint" → "default to native; surface as a post-generation refinement option if unsure".
- Stage 3 (L276): the validator call's `density_mode=` must use the **effective first-gen density** (user-intent mode ‖ template-aware mode ‖ `standard` baseline), not a hardcoded `'standard'` (omission 1).
- Example Interaction (L384–408): update — remove "step 2 density+approval question"; show "autonomous standard + self-critique"; add a new "step 5 post-generation refinements" step; **fix the pre-existing sign-off bug at L408** (MINOR-4: "the template's default sign-off block shows" contradicts Constraint #7 — rewrite to "the engine removes the sign-off placeholder since `presenter_name` is unset").

**Change** `AGENTS.md:98` — rewrite the MANDATORY paragraph from "confirm-first" to "generate-first-then-refine": primary agent's first generation runs zero-prompt straight to render with **template-aware** defaults (standard baseline, concise for small content area) + density-intent detection; after returning the file, the primary agent issues one `question` for refinements (subagent skips); user-stated preferences honored.

**Change** `generate-slide-skill/SKILL.md:317` AND **`:343`** (rev 2 — MAJOR-2) — L317: "the agent picks the mode with the user at the outline-confirmation checkpoint" → "defaults to `standard` for the first generation (template-aware downshift to `concise` for small content areas); adjusted post-generation per user refinement (see `pptx-subagent` Stage 5)". L343: "primary agent can pause after the outline for the user to approve/edit" → "primary agent no longer pauses — first generation is autonomous with defaults; post-generation refinements are offered after file return".

**Change** `density_mode.py:7` (module docstring) — "a mode (with the user, at the outline-confirmation checkpoint)" → "defaults to `standard` for the first generation (template-aware downshift where applicable; adjusted post-generation per user refinement)".

## Acceptance Criteria

- [ ] AC1 — A primary-agent PPT request with no explicit preferences produces a `.pptx` **without any pre-generation `question`** (zero-prompt first generation; `standard` density, no sign-off, native ratio).
- [ ] AC2 — The outline is displayed before generation as information only ("proceeding with defaults"), with no wait for confirmation.
- [ ] AC3 — User-stated preferences in the first message are honored on the first generation (defaults never override them). **Density intent is detected via the word list (Decision 12)** — e.g. "做个简要的 5 页 PPT" → first-gen `density_mode='concise'`; "make a detailed handout" → `text-heavy`; page count ("5 pages") and ratio ("4:3") honored directly.
- [ ] AC4 — After the first file is returned, the **primary agent** issues exactly **one multi-select `question`** presenting the 7 refinement options; picking "No adjustment (recommended)" or none ends the workflow.
- [ ] AC5 — If the user picks refinements, exactly **one** re-generation round runs (re-render or rewrite+re-render per option; mutually-exclusive picks resolved per Decision 6/MINOR-2) and returns a new file; **no second refinement prompt** follows.
- [ ] AC6 — **All 8 functional enforcement points** are updated consistently (no stale "confirm-first" language remains in `AGENTS.md:98`, `pptx-subagent.md` Constraints #7 L97 / pipeline L117 / multi-aspect L129 / Stage 2 L201–242, `SKILL.md:317`/`:343`, or `density_mode.py:7`). (Historical/analysis docs — README, GAP-ANALYSIS, DESIGN-*, outline_store docstring — are explicitly out of scope this issue, not AC6-blocking.)
- [ ] AC7 — Subagent (headless) path is unchanged for generation (still autonomous self-critique) AND **skips the Stage 5 refinement `question` entirely** (headless — no question at all; returns after first generation).
- [ ] AC8 — **Template-aware density (MAJOR-1):** a small-content-area template (`content_area < ~30 in²`) auto-downshifts the first-gen density to `concise` even with no user input, and the Stage 3 validator receives that effective density (not a hardcoded `standard`).

## Implementation Phases

### Phase 0: Core agent prompt rewrite (pptx-subagent.md Stage 2 + Stage 5)
- [ ] T1: Stage 2 (L183–242) — delete the primary-agent three-question branch (L201–233); promote the subagent self-critique rubric (consistency/flow/coverage/redundancy/length) **+ the new 6th "Template fit" dimension (Decision 11)** to the primary-agent first-generation path; **collapse the primary/subagent branches into one autonomous path** (Decision 13 — the distinction moves to Stage 5); retitle Stage 2 to "Density Mode + Self-Critique (autonomous, no pre-generation prompt)"; keep the density-modes table as the single source of truth; record **template-aware density** (standard baseline, concise if content_area < ~30 in²) as the first-gen default; state explicitly that defaults cover only unstated params.
- [ ] T2: Stage 1 outline — keep `save_outline` artifact (incl. re-save with `mode=<effective>` header for traceability); update L161 purpose text ("traceability artifact for Stage 3's density header", not "interactive branch"); change the display language to pure information ("Here's the outline I'll generate — proceeding with defaults; you can adjust count/density in the next step") with no confirmation wait.
- [ ] T3: Stage 5 — extend the return path: after returning the absolute path, the **primary agent** issues one multi-select `question` (`multiple: true`) with the 7 refinement options; document each option's re-gen mechanism (Decision 6); add the **multi-select conflict-resolution rule** (MINOR-2: mutually-exclusive picks resolve to last-selected, or a one-line inline confirm); state "one round only" and that "No adjustment / no selection" ends the workflow; **state the subagent skips this entirely** (Decision 13).
- [ ] T4: Stage 0 — add the explicit "defaults principle" note: user-stated preferences (count/density/ratio) win on first generation; defaults only cover the unstated. **Add the density-intent word list (Decision 12).** **Add the template-aware downshift note (Decision 3/11: small content area → concise is default logic, not user pref).**

### Phase 1: Companion docs + in-file enforcement points
- [ ] T5: `AGENTS.md:98` — rewrite the MANDATORY paragraph from "confirm-first" to "generate-first-then-refine" with **template-aware defaults** + density-intent detection; primary agent first generation zero-prompt to render → return file → primary issues one multi-select `question` for refinements (subagent skips) → one re-gen round → done; user-stated preferences honored.
- [ ] T6: `generate-slide-skill/SKILL.md:317` **AND `:343`** (MAJOR-2) — L317: "picks the mode with the user at the outline-confirmation checkpoint" → "defaults to `standard` for the first generation (template-aware downshift to `concise` for small content areas); adjusted post-generation per user refinement (see pptx-subagent Stage 5)". L343: "primary agent can pause after the outline for user to approve/edit" → "primary agent no longer pauses — first generation is autonomous with defaults; post-generation refinements offered after file return".
- [ ] T7: `density_mode.py:7` module docstring — "a mode (with the user, at the outline-confirmation checkpoint)" → "defaults to `standard` for the first generation (template-aware downshift where applicable; adjusted post-generation per user refinement)".
- [ ] T7b: **`pptx-subagent.md` L116–117 pipeline overview** (MAJOR-2) — remove "pick mode + approve outline in ONE question call"; replace with the autonomous-flow summary.
- [ ] T7c: **`pptx-subagent.md` L129 multi-aspect-ratio** (MAJOR-2) — redirect "ask in the Stage 2 checkpoint" → "default to native; surface as a post-generation refinement if unsure".
- [ ] T7d: **`pptx-subagent.md` Stage 3 L276** (omission 1) — the validator call's `density_mode=` uses the **effective first-gen density** (user-intent ‖ template-aware ‖ `standard`), not hardcoded `'standard'`. *(Placeholder written as `<EFFECTIVE_DENSITY>` without quotes — code-quality MINOR-2 — to match the file's `<JSON_ARRAY>` placeholder convention and avoid literal-fill.)*
- [ ] T7e: **`pptx-subagent.md` Constraint #7 L97** (rev 3 — code-quality MAJOR-1) — rewrite the primary-agent sign-off branch from "proactively ask for name/email" to "leave unset on first generation (engine removes placeholder); ask for name/email only if the Stage 5 refinement picks 'Add presenter sign-off'". Removes the contradiction with Stage 2's "no pre-generation question" rule.

### Phase 2: Example Interaction update + verification
- [ ] T8: `pptx-subagent.md` Example Interaction (L384–408) — remove "step 2 density+approval question"; rewrite step 2 as "autonomous `standard` + self-critique (no prompt)"; add a new step showing the post-generation refinement `question` (7 options) and an example "No adjustment" / single-pick outcome. **Fix the pre-existing sign-off bug at L408 (MINOR-4)** — rewrite "the template's default sign-off block shows" → "the engine removes the sign-off placeholder since `presenter_name` is unset". Keep the `save_outline` re-save-with-mode-header demonstration.
- [ ] T9: Manual verification — run the primary-agent flow once on (a) a no-preference request, (b) a small-content-area template, (c) a density-word request ("做个简要的 5 页 PPT"); confirm (a) no pre-gen `question`, outline info-only, file returned, 7-option multi-select appears; (b) first-gen density auto-downshifted to `concise`; (c) first-gen density = `concise` from intent detection; one re-gen max after a pick, no further prompt (AC1–AC5, AC8).
- [ ] T10: Consistency sweep — grep the repo for stale phrasing across **all 7 functional points**: `outline-confirmation checkpoint`, `single \`question\` call`, `wait for all answers`, `pause after the outline`, `approve outline`, `ask in the Stage 2 checkpoint`, `interactive checkpoint`, `pick mode + approve outline`; confirm none remain in the functional files (AC6). Confirm the subagent path still reads autonomous + skips Stage 5 (AC7). *(Optional MINOR-7: add `tests/test_prompt_consistency.py` as a durable regression guard.)*

## Risks

- **Behavior shift for existing users** — users accustomed to the pre-generation checkpoint will see a different (faster) flow. Mitigated by: the post-generation refinement `question` still surfaces every option that previously appeared pre-generation, just deferred; "No adjustment (recommended)" is the explicit default. **Rollback = `git revert` this commit (prompt-only, no data migration).**
- **Default-density mismatch (rev 2 — MAJOR-1 closed):** a flat `standard` default would mismatch small-content-area templates (US-4.7 guidance says downshift to `concise`). **Mitigated by Decision 3/11**: density is template-aware (auto-downshift to `concise` when `content_area < ~30 in²`), wired into the self-critique rubric's 6th dimension + Stage 3's effective-density validator call. The post-generation "Lower/Increase text density" option remains a one-tap adjustment.
- **Consistency drift across 7 points** — if any functional enforcement point is missed, the agent may behave inconsistently. Mitigated by: AC6 (reworded to "all 7 functional points") + T10 consistency sweep with expanded grep keywords. (Optional MINOR-7: durable `test_prompt_consistency.py`.)
- **Multi-select UX ambiguity (MINOR-2)** — some options are mutually exclusive ("lower density" + "increase density"); OpenCode multi-select has no in-group exclusivity. Mitigated by Decision 6 + T3's conflict-resolution rule (last-selected wins / one-line inline confirm).
- **One-round limit frustration** — a user wanting iterative refinement must issue fresh requests after the single round. Deliberate trade-off (Decision 7).
- **Historical/analysis docs left stale (rev 2 — MAJOR-2 omission 2):** `README.md:149`, `GAP-ANALYSIS.md:139`, `DESIGN-multi-stage-generation.md:37`, `outline_store.py:8/15` carry old "checkpoint/approve" wording. **Out of scope this issue** — these record historical state; a future doc-governance pass can sweep them. Not AC6-blocking (AC6 is scoped to the 7 functional points).

## Out of scope (explicit)

- **Engine code / tests** — `ppt_builder.py`, `density_mode.py` logic, `schema_validator` untouched.
- **`render.json` `generation_pass` field** (MINOR-6) — distinguishing first-vs-second generation products would need a `ppt_builder.py` change, breaking the "prompt-only" boundary. **Deferred**; if two-pass debugging confusion arises, open a follow-up.
- **Historical/analysis doc wording** (see Risks) — not updated this issue.
- **Subagent (headless) path** — unchanged (still autonomous; additionally skips the new Stage 5 question per Decision 13).

## References

- Requirements source: issue #76 — confirmed design decisions (3 locked) + first-generation defaults table + 7 refinement options.
- Architecture review: APPROVE on philosophy + 3 locked decisions; CHANGES-REQUIRED on MAJOR-1 (template-aware density) / MAJOR-2 (enforcement inventory 4→7) / MAJOR-3 (density-intent detection) — all incorporated into this rev 2.
- Functional enforcement points (all 7 — AC6): `AGENTS.md:98`; `.opencode/agents/pptx-subagent.md:117` (overview), `:129` (multi-aspect), `:183–242` (Stage 2), `:276` (Stage 3 density mode), `:384–408` (Example); `generate-slide-skill/SKILL.md:317` + `:343`; `generate-slide-skill/scripts/density_mode.py:7`.
- Historical docs (out of scope): `README.md:149`, `docs/user-stories/GAP-ANALYSIS.md:139`, `generate-slide-skill/docs/DESIGN-multi-stage-generation.md:37`, `outline_store.py:8`+`:15`.
- Reused primitives (no change): subagent self-critique rubric (`pptx-subagent.md:235–242`) + new 6th "Template fit" dimension; density mechanism (`density_mode.py`, `schema_validator`); `save_outline` artifact (`outline_store`); US-4.6 `target_size`; US-4.3 `auto_template`; engine placeholder-removal for unset `presenter_name`/`presenter_email`; Template-aware MANDATORY rule (`pptx-subagent.md:79` + Stage 0:127).
- PLAN format template: `PLANS/PLAN-GIT-70.md`.
