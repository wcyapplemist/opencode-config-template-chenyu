---
name: ppt-template-filler
description: "Populate the PowerPoint template with structured JSON data using a python-pptx based engine. Uses template.pptx Slide Master layouts with proper placeholders. Do NOT use for creating presentations from scratch."
license: Apache-2.0
compatibility: opencode
metadata:
  audience: developers
  workflow: presentation-generation
---

## What I do

I fill the PowerPoint template (`template_tagged.pptx` preferred, `template.pptx` fallback) with structured content using `ppt_builder.py`. I am the **only approved method** for generating presentations from structured data.

- Accept a JSON array (`slide_data_list`) and render it into a `.pptx` file
- Resolve Slide Master layouts **by name** via embedded `<p:ext>` metadata (preferred) or hardcoded mapping (fallback)
- Add slides from the template's Slide Master layouts, filling placeholders by type
- Write English speaker notes to each slide's Notes pane (Presenter View only)
- Handle missing placeholders gracefully with warnings (never crash)

## When to use me

Use this skill when:
- You have structured content (JSON array) and need a `.pptx` output
- The `pptx-specialist-subagent` delegates template-filling work

Do NOT use for:
- Creating presentations from scratch
- OOXML editing or raw XML manipulation
- Thumbnail generation or visual analysis

## Template

The engine uses a single template:

| File | Description |
|------|-------------|
| `scripts/templates/template_tagged.pptx` | Slide Master template with embedded `<p:ext>` layout metadata (preferred) |
| `scripts/templates/template.pptx` | Slide Master template without metadata (fallback) |
| `scripts/templates/template.config.json` | Layout index mapping (`title_slide_layout`, `content_slide_layout`) — legacy override |
| `scripts/inspect_template.py` | Reads layout metadata and outputs a JSON manifest |

### Layout Mapping

Layouts are resolved **by name**, not by index. The engine uses a two-tier resolution:

1. **Metadata-driven (preferred)**: If `template_tagged.pptx` exists, the engine reads `<p:ext>` metadata from each layout to build the mapping dynamically. Each layout declares a `templateId`, `compatibleWith` (backward compat), and `slots`.
2. **Hardcoded fallback**: Uses `_LAYOUT_NAME_MAP` inside `ppt_builder.py`; `template.config.json` overrides the name for `title_slide` / `content_slide`.

To discover all available layouts at runtime:

```bash
python scripts/inspect_template.py scripts/templates/template_tagged.pptx
```

| Slide Type | Layout Name (template.pptx) | Placeholders Used |
|------------|-----------------------------|-------------------|
| `title_slide` | `Title Slide` | CENTER_TITLE + SUBTITLE |
| `content_slide` | `Title and Content` | TITLE + OBJECT |
| `section_header_slide` | `Section Header` | TITLE + BODY |
| `two_content_slide` | `7_Two Content` | TITLE + OBJECT×2 |
| `closing_slide` | `End` | CENTER_TITLE + SUBTITLE |

```json
{
  "title_slide_layout": "Title Slide",
  "content_slide_layout": "Title and Content"
}
```

## Input Data Format

**Language: English only.** All slide content AND speaker notes MUST be in English. Do not translate into any other language, even if the request is in Chinese or explicitly asks for a non-English deck.

```json
[
  {
    "slide_type": "title_slide",
    "title": "BETEKK 2026 Q1 Quarterly Review",
    "subtitle": "March 2026",
    "notes": "KEY MESSAGE: Open with energy — set the context in one line.\n\"Hold the slide for two seconds before you speak.\"\n\"Good [morning/afternoon], I'm [Name]. Welcome to our Q1 review — the short version is, we had a strong quarter.\"\nPause. Let it land.\n\"I'll take you through the numbers, then what's next.\"\nTRANSITION: \"Let's look at the numbers.\"\nCOACHING: Eye contact, confident. Do not read the slide."
  },
  {
    "slide_type": "content_slide",
    "title": "Key Business Metrics",
    "body": "**Revenue Growth** — 32% YoY increase\n**New Contracts** — 18 signed this quarter\n**Customer Satisfaction** — 96.5% approval rating",
    "notes": "KEY MESSAGE: Strong across every metric — revenue, pipeline, and satisfaction.\n\"Three numbers tell the story this quarter.\"\n\"Revenue is up thirty-two percent year on year — our fastest growth yet.\"\nPause. Let the number land.\n\"We signed eighteen new contracts, and customer satisfaction sits at ninety-six-point-five percent.\"\n\"Ask yourself: which of these would you lead with to your board?\"\nTRANSITION: \"Here is what actually drove these results.\"\nCOACHING: Matter-of-fact tone. Be ready for: \"Is the satisfaction score biased?\" — answer: independent survey, 200+ respondents."
  }
]
```

### Field Reference

| Field | Required | Description |
|-------|----------|-------------|
| `slide_type` | Yes | A legacy type (`title_slide`, `content_slide`, `section_header_slide`, `two_content_slide`, `closing_slide`) OR a `templateId` from `inspect_template.py` |
| `title` | Yes | Main heading text |
| `subtitle` | No | Subheading text (legacy types + some slot-based types) |
| `body` | No | Body content for single-body layouts (legacy only). `\n` = new paragraph. Format: `**Title** — Description` |
| `body_left` / `body_right` | No | Left/right column body (legacy `two_content_slide` only) |
| `slots` | No | Dict of `{role: content}` for slot-based layouts. Keys match the `role` field from metadata. Body-style slots accept `**Title** — Description` format |
| `notes` | Yes | Full English presenter script (**~100–150 words**). Written to the slide's Notes pane (Presenter View only). `\n` = new paragraph. Must be **spoken dialogue**, **stage directions**, a `TRANSITION` line, and `COACHING** — NOT bullet summaries. |

### Slot-Based Layouts

When using a `templateId` as `slide_type`, provide content via the `slots` dict:

```json
[
  {
    "slide_type": "title_slide",
    "title": "Quarterly Review",
    "subtitle": "Q3 2026",
    "notes": "..."
  },
  {
    "slide_type": "content_three_column",
    "title": "Performance Highlights",
    "slots": {
      "col_1": "**Revenue** — $4.2M (+32% YoY)",
      "col_2": "**Margin** — 38% (target: 35%)",
      "col_3": "**NPS** — 72 (industry avg: 55)"
    },
    "notes": "..."
  },
  {
    "slide_type": "content_grid_2x2",
    "title": "SWOT Analysis",
    "slots": {
      "cell_top_left": "**Strengths** — Market leader in SE Asia",
      "cell_top_right": "**Weaknesses** — Limited EU presence",
      "cell_bottom_left": "**Opportunities** — Green building mandate",
      "cell_bottom_right": "**Threats** — Low-cost competitors"
    },
    "notes": "..."
  },
  {
    "slide_type": "closing_slide",
    "title": "Thank You",
    "subtitle": "Questions?",
    "notes": "..."
  }
]
```

## Output Path

Output files saved under `D:\BETEKK\opencode-config-template\opencode_app\output\`.

## Execution

```bash
python -c "
import json, sys
sys.path.insert(0, 'scripts')
from ppt_builder import generate_ppt_from_data, DEFAULT_OUTPUT_DIR

slide_data = <JSON_ARRAY>
result = generate_ppt_from_data(
    slide_data,
    output_path=str(DEFAULT_OUTPUT_DIR / 'report.pptx'),
)
print(result)
"
```

## Error Handling

| Scenario | Behavior |
|----------|----------|
| Unknown `slide_type` | Log warning, skip, continue |
| Missing placeholder | Log warning, skip field, continue |
| Single slide fails | Log error, skip slide, continue |
| Template file missing | Raise `FileNotFoundError` (fatal) |

## Output

Returns the **absolute path** of the generated `.pptx` file.
