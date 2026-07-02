---
name: template-modifier-skill
description: "Extend a PowerPoint template when content exceeds its limits. Reads the Slide Master, checks whether a requirement fits the template's layouts, and (when it doesn't) clones a new layout into a derived template_new.pptx. Works alongside the generate-slide-skill engine (Capability B). Do NOT use for normal template filling — use generate-slide-skill for that."
license: Apache-2.0
compatibility: opencode
metadata:
  audience: developers
  workflow: template-extension
---

## What I do

I am the **template-modifier-skill** (Capability B). When a deck's content exceeds what the base `template.pptx` can hold — a layout is missing, or a body is too large for its placeholder — I extend the template by **cloning a new layout** into a derived `template_new.pptx`, which the `generate-slide-skill` engine then renders against.

I do **not** fill templates myself. Normal filling is the `generate-slide-skill` skill's job. I am invoked only when the base template cannot satisfy a requirement.

## The 4 stakeholder steps

My pipeline mirrors the four steps a human designer performs:

1. **Read the template** — resolve the render contract via `ppt_builder.get_render_contract` (US-4.1: prefers the embedded JSON, falls back to the P0 introspection engine `template_introspector`) to get the full contract: layouts, placeholder fingerprints, `content_area_in2`, theme, slide size.
2. **Read the Slide Master** — `template_reader.read_master()` reads master-level placeholders + theme (on top of the contract).
3. **Understand the requirement** — `constraint_checker.evaluate_slide()` estimates the content area a slide needs (from its word count) and compares it against the layout's `content_area_in2`, yielding a **fits / over-limit** verdict. It also flags a `slide_type` whose layout is missing.
4. **Over-limit → create** — when a slide is over-limit (or its layout is missing), `state_machine.plan_resolution()` plans a clone; P4's `layout_creator` performs the actual XML/part clone into `template_new.pptx`.

## The `template_new.pptx` lifecycle (DESIGN §5)

Two file roles:

| File | Role |
|------|------|
| `template.pptx` | **Immutable base** (user-supplied, single authoritative path). |
| `template_new.pptx` | **Derived / ephemeral** — produced only when the base cannot satisfy a layout requirement. |

On **every** generation request, the state machine runs:

1. **① Delete leftover** — if `template_new.pptx` exists, delete it (force freshness; the base is re-evaluated each request).
2. **② Introspect base** — `get_render_contract` (embedded-preferred, sidecar fallback).
3. **③ Scan** — for each slide, check its fingerprint + content size against the contract; collect any over-limit / missing-layout slides into a clone plan.
4. **④ Clone** (P4) — produce `template_new.pptx` with the extended layout(s); swap the active template.
5. **⑤ Notify** — whenever `template_new.pptx` is used, emit a **mandatory** user notice naming the template + the reason (`template.pptx could not fit <reason>`).

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/template_reader.py` | `read_master(template_path)` — Slide Master placeholders + delegated contract. |
| `scripts/constraint_checker.py` | `evaluate_slide(...)` / `check_content_area(...)` — over-limit verdict. |
| `scripts/state_machine.py` | `plan_resolution(...)` — the ①→②→③→⑤ lifecycle; `resolve_and_clone(...)` — the **full loop** (plan + clone + hand-off); `build_notification(...)` — the user notice. |
| `scripts/layout_creator.py` | `clone_for_over_limit(...)` — the 7-step XML/part clone into `template_new.pptx` (DESIGN §7). |

## Usage — the full Capability B loop

`resolve_and_clone(base, slides)` runs the whole pipeline: it plans (①②③), and when a slide is over-limit it **clones an extended layout** into `template_new.pptx` (P4), then returns the active template + the layout-name pins + the mandatory notification. Hand the result to the `generate-slide-skill` engine:

```bash
python -c "
import sys
sys.path.insert(0, '.opencode/skills/template-modifier-skill/scripts')
sys.path.insert(0, '.opencode/skills/generate-slide-skill/scripts')
from state_machine import resolve_and_clone
from ppt_builder import generate_ppt_from_data, DEFAULT_OUTPUT_DIR

active, overrides, note = resolve_and_clone(
    '.opencode/skills/generate-slide-skill/scripts/templates/template.pptx',
    <SLIDE_DATA_LIST>,
)
out = generate_ppt_from_data(
    <SLIDE_DATA_LIST>,
    template_path=active,                 # base, or template_new.pptx when a clone was made
    config_overrides=overrides,           # pins the over-limit slide_types to the extended layouts
    output_path=str(DEFAULT_OUTPUT_DIR / 'deck.pptx'),
)
print(out)
if note:
    print(note)                           # MANDATORY: tell the user template_new.pptx was used + why
"
```

If cloning fails, `resolve_and_clone` **safely falls back** to the base template (no derived file produced) — the deck still renders, just without the extended layout.

### Safety guarantees

- The base `template.pptx` is **never written** — clones save only to the derived `template_new.pptx`.
- **Reload-verify** after every clone: the cloned layout must be findable by `get_by_name`.
- **Rollback**: any clone/verify failure deletes `template_new.pptx` so a broken derived file is never left behind.

## When to use me

- A slide's body overflows its placeholder (content area exceeded).
- A required `slide_type` has no matching layout in the template.
- You need a larger / differently-shaped layout than the base template provides.

Do **NOT** use me for normal filling, chart generation, or image embedding — those are `generate-slide-skill`.

## Reference

- Design: `.opencode/skills/generate-slide-skill/docs/DESIGN-template-agnostic.md` — §5 (state machine), §7 (Capability B pipeline + 7-step clone).
