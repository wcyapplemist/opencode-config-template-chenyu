# Design: Multi-Stage Generation Pipeline

> **Issue #21** (autonomous, default) + **Issue #24** (interactive, primary-agent only) — Phase 1, Track C.

## 1. Why multi-stage

The current workflow generates the **entire** `slide_data_list` JSON in a single
LLM pass (agent Step 2). As slide count grows this degrades: titles drift, flow
breaks, topics overlap, and notes thin out. presenton splits generation into
three steps (`outlines` → `structure` → `content`) for exactly this reason.

## 2. Three stages

### Stage 1 — Outline
Produce a **plain-text outline** (no JSON yet). One entry per planned slide:

```
1. [title_slide]    "AI in Construction" — subtitle 2026
2. [content_slide]  "Why now" — market pressure, labor gap, tech maturity
3. [chart_slide]    "Market growth" — bar, 2020-2026 (data_query: market size)
4. [content_image_slide] "Field example" — drone surveying (image_prompt: ...)
...
```

Each entry records: slide order, `slide_type`, a working title, and the key
points / resource placeholders. This is cheap to review and revise.

### Stage 2 — Self-critique (autonomous) / Review (interactive)
Feed the outline back to the LLM and ask it to critique and revise:
- **Consistency** — do titles tell one coherent story?
- **Flow** — does each slide set up the next?
- **Coverage gaps** — missing context, obvious omissions.
- **Redundancy** — slides that repeat each other.
- **Length** — right number of slides for the ask.

The revised outline is the input to Stage 3.

### Stage 3 — Detail + JSON
Convert the revised outline into the full `slide_data_list` JSON:
- Write the actual body text per slide.
- Write the full English speaker notes (~120–180 words, house style).
- Emit resource placeholders (`image_prompt`, `icon_query`, `data_query`) where
  the slide needs a resolved asset.

Every stage that emits JSON is **gated by schema validation** (#20): the stage
output is run through `validate_slide_data_list(strict=True)` before proceeding.

## 3. Autonomous vs interactive

OpenCode **subagents run headless**: a single prompt → an autonomous run → a
single result. They **cannot pause mid-run** to ask the user a question.

| Mode     | When                                   | Behaviour                                                       |
|----------|----------------------------------------|-----------------------------------------------------------------|
| Autonomous | default; any subagent invocation     | Stage 2 = LLM self-critique, no human pause. Portable & safe.  |
| Interactive | only when `pptx-subagent` is the **primary** conversation agent | After Stage 1, the outline is surfaced to the user; Stage 2 proceeds only after approval/edits. |

The agent detects its context: if it is the primary agent with a live user
turn-loop, it offers the checkpoint; otherwise it uses autonomous mode. When in
doubt (e.g. cannot confirm a user channel), it falls back to autonomous —
**never hangs** waiting for input.

### Two-call fallback for subagent context
If an orchestrating agent wants a human checkpoint while invoking `pptx-subagent`
as a subagent, it can use a **two-call** pattern:
1. First call → agent returns the **outline artifact only** (Stage 1).
2. Orchestrator surfaces the outline to the user, collects edits.
3. Second call → agent resumes from the edited outline through Stages 2–3.

## 4. Outline artifact

The outline is persisted to a temp file so it can be surfaced/edited later (the
foundation for #24's interactive checkpoint). A small helper,
`outline_store.py`, provides:

```python
save_outline(outline_text, deck_id=None) -> Path   # persists, returns path
load_outline(path) -> str                           # reads it back
```

Artifacts live under the OS temp dir (`opencode/pptx_outline_<id>.md`) so they
never pollute the repo.

## 5. Critique prompt (Stage 2 contract)

The self-critique call is given: the original request, the Stage 1 outline, and
a fixed rubric (the four lenses above). It returns a **revised outline** in the
same format, plus an optional short "changes" note. The agent does **not**
proceed to JSON until the revised outline is non-empty and internally consistent.

## 6. Acceptance mapping

- Generation proceeds in three explicit stages. ✔ (workflow)
- Long decks show better title/flow consistency. ✔ (self-critique gate)
- Every JSON stage is schema-validated before continuing. ✔ (#20, strict)
- Works headless as a subagent. ✔ (autonomous default)
- Outline artifact persisted for interactive use. ✔ (`outline_store.py`)
