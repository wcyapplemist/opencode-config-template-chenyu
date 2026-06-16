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
- Resolve Slide Master layouts **by name** (robust to layout reordering)
- Add slides from the template's Slide Master layouts, filling placeholders by type
- Write English speaker notes to each slide's Notes pane (Presenter View only)
- Handle missing placeholders gracefully with warnings (never crash)

## When to use me

Use this skill when:
- You have structured content (JSON array) and need a `.pptx` output
- You want to populate a Slide Master template with pre-defined layouts

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

Layouts are resolved **by name**, not by index. The default mapping (`slide_type` → layout name) lives in `_LAYOUT_NAME_MAP` inside `ppt_builder.py`; `template.config.json` overrides the name for `title_slide` / `content_slide`.

| Slide Type | Layout Name (template.pptx) | Placeholders Used |
|------------|-----------------------------|-------------------|
| `title_slide` | `Title Slide` | CENTER_TITLE + SUBTITLE |
| `content_slide` | `Title and Content` | TITLE + OBJECT |
| `section_header_slide` | `Section Header` | TITLE + BODY |
| `two_content_slide` | `7_Two Content` | TITLE + OBJECT×2 |
| `comparison_slide` | `Comparison` | TITLE + OBJECT×2 |
| `content_image_slide` | `Picture with Caption` | TITLE + BODY |
| `chart_slide` | `Blank` | TITLE + native chart |
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

| Field | Required | Slide Type | Description |
|-------|----------|------------|-------------|
| `slide_type` | Yes | All | One of: `title_slide`, `content_slide`, `section_header_slide`, `two_content_slide`, `comparison_slide`, `content_image_slide`, `chart_slide`, `closing_slide` |
| `title` | Yes | All | Main heading text |
| `subtitle` | No | `title_slide`, `closing_slide` | Subheading text |
| `body` | No | `content_slide`, `content_image_slide` | Body content. `\n` = new paragraph. Format: `**Title** — Description` |
| `body_left` / `body_right` | No | `two_content_slide`, `comparison_slide` | Left/right column body (same body-text format) |
| `chart_type` | Yes | `chart_slide` | Chart type: `bar`, `bar_stacked`, `bar_horizontal`, `bar_horizontal_stacked`, `pie`, `pie_exploded`, `doughnut`, `line`, `line_markers` |
| `categories` | Yes | `chart_slide` | Array of category labels (X-axis or pie slice labels) |
| `series` | Yes | `chart_slide` | Array of `{name, values}` objects. Multiple series supported for bar/line. |
| `chart_options` | No | `chart_slide` | Styling options (see Chart Options below) |
| `notes` | Yes | All | Full English presenter script (**~120–180 words**). Written to the slide's Notes pane (Presenter View only). `\n` = new paragraph. Must be **spoken dialogue** (quoted, speakable sentences tied to the slide's content), **interspersed stage directions**, a `TRANSITION` line, and `COACHING` with delivery + an anticipated Q&A — NOT bullet summaries. Cover/closing use `[Name]` / `[morning/afternoon]` placeholders. |

### Body Text Parsing

Each line is parsed into a bold title run + description run:
- Split at first ` — `, ` - `, or `: `
- `**` markers stripped automatically
- No card slot limit — body is a single multi-paragraph block

## Chart Slides

The `chart_slide` type generates **native PowerPoint charts** (editable in PowerPoint, not images) using `python-pptx`'s `add_chart()` API. Charts use the `Blank` layout (which has a TITLE placeholder + free space for the chart graphic).

### Chart Type Reference

| `chart_type` | XL_CHART_TYPE | Description |
|---|---|---|
| `bar` | COLUMN_CLUSTERED | Vertical bars (default) |
| `bar_stacked` | COLUMN_STACKED | Stacked vertical bars |
| `bar_horizontal` | BAR_CLUSTERED | Horizontal bars |
| `bar_horizontal_stacked` | BAR_STACKED | Stacked horizontal bars |
| `pie` | PIE | Standard pie chart |
| `pie_exploded` | PIE_EXPLODED | Exploded pie chart |
| `doughnut` | DOUGHNUT | Doughnut chart |
| `line` | LINE | Simple line chart |
| `line_markers` | LINE_MARKERS | Line with data point markers (recommended) |

### Chart Options

All fields inside `chart_options` are optional with sensible defaults:

| Option | Type | Default | Description |
|---|---|---|---|
| `legend_position` | string | `"bottom"` | `"bottom"`, `"right"`, `"top"`, `"left"`, or `"none"` |
| `show_data_labels` | bool | `true` | Show value/percentage labels on chart |
| `value_format` | string | `"#,##0.0"` | Number format for bar/line data labels |
| `y_axis_format` | string | `"#,##0.0"` | Number format for Y-axis tick labels |
| `y_axis_min` | float | auto | Y-axis minimum scale |
| `y_axis_max` | float | auto | Y-axis maximum scale |
| `y_axis_major_unit` | float | auto | Y-axis major gridline interval |
| `y_axis_title` | string | `""` | Y-axis title text |
| `x_axis_title` | string | `""` | X-axis title text |

### Theme Styling

Charts automatically use the template's theme:
- **Colors**: 8-color palette extracted from theme accent colors (`accent1`-`accent6` + `dk2` + `accent3`)
- **Fonts**: All chart text uses `Calibri` (theme minor font)
- **Gridlines**: Major gridlines in `#E7E6E6` (theme `lt2`), 0.75pt
- **Axis text**: `#44546A` (theme `dk2`)

### Example: Bar Chart (single series)

```json
{
  "slide_type": "chart_slide",
  "title": "Global Construction Tech Market (USD Billion)",
  "chart_type": "bar",
  "categories": ["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
  "series": [
    {"name": "Market Size", "values": [8.5, 11.2, 14.8, 19.5, 25.1, 31.7, 39.4]}
  ],
  "chart_options": {
    "legend_position": "bottom",
    "show_data_labels": true,
    "y_axis_min": 0,
    "y_axis_max": 45
  },
  "notes": "KEY MESSAGE: Market growing from 8.5B to 39.4B by 2026."
}
```

### Example: Pie Chart

```json
{
  "slide_type": "chart_slide",
  "title": "Technology Adoption Rate",
  "chart_type": "pie",
  "categories": ["BIM", "IoT", "Drones", "AI & ML", "Robotics", "Cloud"],
  "series": [
    {"name": "Adoption %", "values": [68, 45, 52, 28, 15, 72]}
  ],
  "chart_options": {
    "legend_position": "right"
  },
  "notes": "KEY MESSAGE: Cloud and BIM lead adoption at 72% and 68%."
}
```

### Example: Line Chart (multi-series)

```json
{
  "slide_type": "chart_slide",
  "title": "Project Performance Improvement (%)",
  "chart_type": "line_markers",
  "categories": ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  "series": [
    {"name": "Cost Savings",       "values": [5, 8, 12, 16, 20, 25, 30]},
    {"name": "Schedule Reduction", "values": [3, 6, 10, 14, 19, 24, 28]},
    {"name": "Safety Improvement", "values": [2, 4,  8, 12, 18, 22, 27]}
  ],
  "chart_options": {
    "legend_position": "bottom",
    "y_axis_min": 0,
    "y_axis_max": 35
  },
  "notes": "KEY MESSAGE: All three metrics show consistent improvement."
}
```

## Output Path

Output files saved under `<project_root>/output/`.

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
| Unknown `chart_type` | Default to `bar`, log warning |
| Missing `categories` or `series` | Skip chart, log warning |
| Invalid `chart_options` field | Ignore, use default |

## Output

Returns the **absolute path** of the generated `.pptx` file.
