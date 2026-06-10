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

I fill the PowerPoint template (`template.pptx`) with structured content using `ppt_builder.py`. I am the **only approved method** for generating presentations from structured data.

- Accept a JSON array (`slide_data_list`) and render it into a `.pptx` file
- Use two slide types: `title_slide` and `content_slide`
- Add slides from the template's Slide Master layouts, filling placeholders by type
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
| `scripts/templates/template.pptx` | Slide Master template with named layouts and placeholders |
| `scripts/templates/template.config.json` | Layout index mapping (`title_slide_layout`, `content_slide_layout`) |

### Layout Mapping

| Slide Type | Layout Index | Layout Name | Placeholders Used |
|------------|-------------|-------------|-------------------|
| `title_slide` | 0 | Title Slide | TITLE + SUBTITLE |
| `content_slide` | 8 | Title and content | CENTER_TITLE + OBJECT |

## Input Data Format

```json
[
  {
    "slide_type": "title_slide",
    "title": "BETEKK 2026 Q1 Quarterly Review",
    "subtitle": "March 2026"
  },
  {
    "slide_type": "content_slide",
    "title": "Key Business Metrics",
    "body": "**Revenue Growth** — 32% YoY increase\n**New Contracts** — 18 signed this quarter\n**Customer Satisfaction** — 96.5% approval rating"
  }
]
```

### Field Reference

| Field | Required | Slide Type | Description |
|-------|----------|------------|-------------|
| `slide_type` | Yes | All | Must be `"title_slide"` or `"content_slide"` |
| `title` | Yes | All | Main heading text |
| `subtitle` | No | `title_slide` only | Subheading text |
| `body` | No | `content_slide` only | Body content. `\n` = new paragraph. Format: `**Title** — Description` |

### Body Text Parsing

Each line is parsed into a bold title run + description run:
- Split at first ` — `, ` - `, or `: `
- `**` markers stripped automatically
- No card slot limit — body is a single multi-paragraph block

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
