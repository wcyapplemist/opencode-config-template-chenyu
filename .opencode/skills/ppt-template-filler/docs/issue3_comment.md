## Research Complete ✅

Full research findings document: `.opencode/skills/ppt-template-filler/docs/RESEARCH-chart-generation.md`

Prototype output: `output/prototype_charts.pptx` (3 charts verified as native PowerPoint objects)

---

### Key Findings Summary

**1. Layout Choice: Layout [9] `Blank`** ⭐
- Has a TITLE placeholder (engine already handles filling it)
- Chart area: x=0.92in, y=2.0in, w=11.5in, h=4.5in
- Footer/SlideNumber/Date preserved

**2. Single `chart_slide` Type (Design Decision)**
Instead of separate `bar_chart_slide` / `pie_chart_slide` / `line_chart_slide`, use ONE `chart_slide` type with a `chart_type` field. All chart types map to the same "Blank" layout, so dispatch happens at the rendering level.

```python
# _LAYOUT_NAME_MAP addition:
"chart_slide": ["Blank"],
```

**3. JSON Schema (validated with prototype):**
```json
{
  "slide_type": "chart_slide",
  "title": "Global Construction Tech Market",
  "chart_type": "bar",
  "categories": ["2020", "2021", "2022", "2023"],
  "series": [
    {"name": "Market Size", "values": [8.5, 11.2, 14.8, 19.5]}
  ],
  "chart_options": {
    "legend_position": "bottom",
    "show_data_labels": true,
    "y_axis_min": 0,
    "y_axis_max": 45
  },
  "notes": "..."
}
```

**4. Chart Type Mapping:**
| JSON `chart_type` | XL_CHART_TYPE |
|---|---|
| `bar` | COLUMN_CLUSTERED |
| `bar_stacked` | COLUMN_STACKED |
| `bar_horizontal` | BAR_CLUSTERED |
| `pie` | PIE |
| `doughnut` | DOUGHNUT |
| `line` | LINE |
| `line_markers` | LINE_MARKERS |

**5. Prototype Results:**
| Chart | Type | Series | Data Labels | Legend | Notes | Status |
|---|---|---|---|---|---|---|
| Bar | COLUMN_CLUSTERED | 1 | Yes | Bottom | Yes | Pass |
| Pie | PIE | 1 | Yes (%) | Right | Yes | Pass |
| Line | LINE_MARKERS | 3 | Yes | Bottom | Yes | Pass |

**6. Impact on Sub-Issues:**
Since we use a single `chart_slide` type, issues #2 (bar), #5 (pie), #6 (line) can be implemented in a **single PR** — one `_add_chart_to_slide()` function with `chart_type` dispatch, rather than three separate code paths.

---

### Acceptance Criteria Checklist

- [x] python-pptx chart API investigated (version 1.0.2, `add_chart()` confirmed)
- [x] `CategoryChartData` structure documented (categories + multi-series)
- [x] Layout analysis complete (Layout [9] "Blank" recommended with exact positioning)
- [x] JSON schema designed for `chart_slide` (bar, pie, line all covered)
- [x] Prototype validated all 3 chart types as native PowerPoint objects
- [x] Implementation plan documented for sub-issues #2, #5, #6, #4, #7, #8

Closing this issue as research is complete. Implementation can proceed with #2/#5/#6.
