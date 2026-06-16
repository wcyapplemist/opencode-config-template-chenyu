# Research: Native Chart Generation in ppt_builder.py

**Issue:** [#3](https://github.com/wcyapplemist/opencode-config-template-chenyu/issues/3)
**Date:** 2026-06-15
**Status:** ✅ Complete

---

## 1. Environment

| Item               | Value                                                                  |
| ------------------ | ---------------------------------------------------------------------- |
| python-pptx        | **1.0.2** (stable release)                                             |
| Python             | 3.12                                                                   |
| Template           | `template.pptx` — 35 layouts, 16:9 widescreen (13.33in × 7.50in)      |
| Chart data class   | `CategoryChartData` (supports categories + named series with values)  |
| Chart placement    | `SlideShapes.add_chart(chart_type, x, y, cx, cy, chart_data)`         |

---

## 2. Layout Analysis

### Candidate Layouts

| Layout Index | Name      | Has TITLE? | Has DATE? | Usable Chart Area (W × H) |
| ------------ | --------- | ---------- | --------- | ------------------------- |
| 6            | `1_Blank` | ❌ No      | ❌ No     | 13.33in × 7.50in (full)   |
| 9            | `Blank`   | ✅ Yes     | ✅ Yes    | 11.5in × 4.5in (below title) |

### Recommendation: Layout [9] `Blank`

**Layout [9] is the clear winner** because:

1. **Has a TITLE placeholder** (idx=0, type=TITLE) — the engine's existing `_find_title_placeholder()` already handles filling it, so chart slides will have consistent titles with all other slide types.
2. **Title occupies** top=0.40in, height=1.45in → bottom at 1.85in.
3. **Chart area** below the title: top=2.0in to ~6.5in (above footer at 6.95in).
4. Footer + Slide Number + Date placeholders are preserved (consistent with other slides).

### Recommended Chart Positioning

```
┌─────────────────────────────────────────────────┐  ← 0.0in
│                                                 │
│    [TITLE PLACEHOLDER]                          │  ← top: 0.40in, h: 1.45in
│                                                 │
├─────────────────────────────────────────────────┤  ← 2.0in
│                                                 │
│                                                 │
│         [CHART GRAPHIC FRAME]                   │  ← top: 2.0in
│         left: 0.92in  width: 11.5in             │
│         height: 4.5in                           │
│                                                 │
│                                                 │
├─────────────────────────────────────────────────┤  ← 6.5in
│ [Date]    [Footer]        [Slide #]             │  ← 6.95in (footer area)
└─────────────────────────────────────────────────┘  ← 7.50in
```

| Parameter | Value              | EMU              |
| --------- | ------------------ | ---------------- |
| `x`       | 0.92 in            | 840,528          |
| `y`       | 2.0 in             | 1,828,800        |
| `cx`      | 11.5 in            | 10,515,600       |
| `cy`      | 4.5 in             | 4,114,800        |

---

## 3. python-pptx Chart API Findings

### 3.1 Chart Data: `CategoryChartData`

```python
from pptx.chart.data import CategoryChartData

chart_data = CategoryChartData()
chart_data.categories = ['2020', '2021', '2022']       # X-axis labels
chart_data.add_series('Market Size', (8.5, 11.2, 14.8)) # Named data series
chart_data.add_series('Forecast',     (9.0, 12.0, 15.5)) # Multiple series supported
```

- `categories` accepts a list/tuple of strings.
- `add_series(name, values, number_format=None)` — `values` is any iterable of numbers.
- Multiple series are fully supported (critical for grouped/stacked bar charts and multi-series line charts).

### 3.2 Chart Creation: `SlideShapes.add_chart()`

```python
graphic_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,   # chart type enum
    Inches(0.92), Inches(2.0),        # x, y position
    Inches(11.5), Inches(4.5),        # cx, cy size
    chart_data,                        # CategoryChartData
)
chart = graphic_frame.chart  # Access the Chart object for styling
```

- Returns a `GraphicFrame`, NOT a `Chart` directly. Access via `.chart`.
- The chart is a **native PowerPoint chart object** (editable in PowerPoint, not a static image).

### 3.3 Relevant `XL_CHART_TYPE` Values

Of the 73 available chart types, we need only these for v1:

| Chart Category | XL_CHART_TYPE Enum       | Value | Use Case                          |
| -------------- | ------------------------ | ----- | --------------------------------- |
| **Bar**        | `COLUMN_CLUSTERED`       | 51    | Vertical bars (default)           |
|                | `COLUMN_STACKED`         | 52    | Stacked vertical bars             |
|                | `BAR_CLUSTERED`          | 57    | Horizontal bars                   |
|                | `BAR_STACKED`            | 58    | Stacked horizontal bars           |
| **Pie**        | `PIE`                    | 5     | Standard pie chart                |
|                | `PIE_EXPLODED`           | 69    | Exploded pie                      |
|                | `DOUGHNUT`               | -4120 | Doughnut chart                    |
| **Line**       | `LINE`                   | 4     | Simple line                       |
|                | `LINE_MARKERS`           | 65    | Line with data point markers (recommended) |
|                | `LINE_MARKERS_STACKED`   | 66    | Stacked line with markers         |

### 3.4 Styling Capabilities

**Chart-level:**
| Property          | Type    | Example                             |
| ----------------- | ------- | ----------------------------------- |
| `chart.has_title` | bool    | `True` / `False`                    |
| `chart.has_legend`| bool    | `True`                              |
| `chart.legend.position` | enum | `XL_LEGEND_POSITION.BOTTOM`        |
| `chart.legend.include_in_layout` | bool | `False` (overlay, don't shrink plot) |
| `chart.font`      | Font    | `chart.font.size = Pt(12)`          |

**Plot-level (per series group):**
| Property                    | Type | Example                              |
| --------------------------- | ---- | ------------------------------------ |
| `plot.has_data_labels`      | bool | `True`                               |
| `plot.data_labels.font.size`| Pt   | `Pt(10)`                             |
| `plot.data_labels.position` | enum | `XL_LABEL_POSITION.OUTSIDE_END`     |
| `plot.data_labels.show_percentage` | bool | `True` (for pie charts)      |
| `plot.data_labels.show_value` | bool | `True` (for bar/line charts)       |
| `plot.gap_width` (Bar only) | int  | `150` (gap between bar groups)       |

**Axis-level (CategoryAxis / ValueAxis):**
| Property                        | Type | Example                 |
| ------------------------------- | ---- | ----------------------- |
| `value_axis.minimum_scale`      | float| `0`                     |
| `value_axis.maximum_scale`      | float| `45`                    |
| `value_axis.has_major_gridlines`| bool | `True`                  |
| `value_axis.major_unit`         | float| `5`                     |
| `category_axis.tick_label_position` | enum | `XL_TICK_LABEL_POSITION.LOW` |

**Series-level:**
| Property        | Type    | Example                          |
| --------------- | ------- | -------------------------------- |
| `series.format` | ChartFormat | Fill color, line color, etc. |

### 3.5 Limitations Discovered

1. **No Plot class export** — `from pptx.chart.plot import Plot` fails; must use `_BasePlot` or type-specific classes (`BarPlot`, `PiePlot`, `LinePlot`). Access via `chart.plots[0]`.
2. **No Series class export** — `from pptx.chart.series import Series` fails; use `_BaseSeries`. Access via `plot.series[i]`.
3. **Theme color extraction** — python-pptx does NOT expose a simple API to read theme colors from `template.pptx`. For v1, we will use a **hardcoded color palette** (can be extracted manually from the template and refined in issue #4).
4. **Pie chart data label positioning** — `OUTSIDE_END` works for pie charts but may overlap with small slices. `BEST_FIT` is more robust but not always available.

---

## 4. JSON Schema Design

### 4.1 Key Design Decision: Single `chart_slide` Type

Instead of separate `bar_chart_slide` / `pie_chart_slide` / `line_chart_slide` types, we recommend a **single `chart_slide` slide_type** with a `chart_type` field that determines the chart rendering.

**Rationale:**
- All chart types map to the **same layout** ("Blank"), so `_LAYOUT_NAME_MAP` needs only one entry: `"chart_slide": ["Blank"]`.
- The rendering logic dispatches on `chart_type` — cleaner code, less repetition.
- Users specify `chart_type` in JSON data, making it easy to switch between chart types.

```python
# _LAYOUT_NAME_MAP addition:
"chart_slide": ["Blank"],
```

### 4.2 Full JSON Schema

```json
{
  "slide_type": "chart_slide",
  "title": "Global Construction Tech Market (USD Billion)",
  "chart_type": "bar",
  "categories": ["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
  "series": [
    {
      "name": "Market Size",
      "values": [8.5, 11.2, 14.8, 19.5, 25.1, 31.7, 39.4]
    }
  ],
  "chart_options": {
    "legend_position": "bottom",
    "show_data_labels": true,
    "data_label_format": "value",
    "y_axis_min": 0,
    "y_axis_max": 45,
    "y_axis_title": "USD Billion"
  },
  "notes": "KEY MESSAGE: The construction tech market is growing exponentially..."
}
```

### 4.3 Field Reference

#### Required Fields

| Field         | Type     | Description                                         |
| ------------- | -------- | --------------------------------------------------- |
| `slide_type`  | string   | Must be `"chart_slide"`                             |
| `title`       | string   | Slide title (filled into TITLE placeholder)         |
| `chart_type`  | string   | One of: `"bar"`, `"bar_stacked"`, `"bar_horizontal"`, `"bar_horizontal_stacked"`, `"pie"`, `"pie_exploded"`, `"doughnut"`, `"line"`, `"line_markers"` |
| `categories`  | string[] | X-axis labels or pie slice labels                   |
| `series`      | array    | Data series (see below)                             |

#### `series` Object

| Field    | Type     | Description                          |
| -------- | -------- | ------------------------------------ |
| `name`   | string   | Series name (shown in legend)        |
| `values` | number[] | Data values matching categories count |

#### `chart_options` (Optional)

| Field                | Type   | Default     | Description                                    |
| -------------------- | ------ | ----------- | ---------------------------------------------- |
| `legend_position`    | string | `"bottom"`  | `"bottom"`, `"right"`, `"top"`, `"left"`, `"none"` |
| `show_data_labels`   | bool   | `true`      | Show value labels on chart                     |
| `data_label_format`  | string | `"value"`   | `"value"`, `"percentage"` (pie/doughnut only)  |
| `y_axis_min`         | float  | auto        | Y-axis minimum scale                           |
| `y_axis_max`         | float  | auto        | Y-axis maximum scale                           |
| `y_axis_title`       | string | `""`        | Y-axis title text                              |
| `x_axis_title`       | string | `""`        | X-axis title text                              |

#### `notes` (Optional)

| Field   | Type   | Description             |
| ------- | ------ | ----------------------- |
| `notes` | string | Speaker notes (English) |

### 4.4 Chart Type → XL_CHART_TYPE Mapping

```python
_CHART_TYPE_MAP = {
    "bar":                   XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar_stacked":           XL_CHART_TYPE.COLUMN_STACKED,
    "bar_horizontal":        XL_CHART_TYPE.BAR_CLUSTERED,
    "bar_horizontal_stacked":XL_CHART_TYPE.BAR_STACKED,
    "pie":                   XL_CHART_TYPE.PIE,
    "pie_exploded":          XL_CHART_TYPE.PIE_EXPLODED,
    "doughnut":              XL_CHART_TYPE.DOUGHNUT,
    "line":                  XL_CHART_TYPE.LINE,
    "line_markers":          XL_CHART_TYPE.LINE_MARKERS,
}
```

### 4.5 Example: Multi-Series Line Chart

```json
{
  "slide_type": "chart_slide",
  "title": "Project Performance Improvement Over Time (%)",
  "chart_type": "line_markers",
  "categories": ["2019", "2020", "2021", "2022", "2023", "2024", "2025"],
  "series": [
    {"name": "Cost Savings",        "values": [5, 8, 12, 16, 20, 25, 30]},
    {"name": "Schedule Reduction",  "values": [3, 6, 10, 14, 19, 24, 28]},
    {"name": "Safety Improvement",  "values": [2, 4, 8, 12, 18, 22, 27]}
  ],
  "chart_options": {
    "legend_position": "bottom",
    "show_data_labels": true,
    "y_axis_min": 0,
    "y_axis_max": 35
  },
  "notes": "KEY MESSAGE: All three metrics show consistent improvement."
}
```

---

## 5. Implementation Plan

### 5.1 Changes to `ppt_builder.py`

1. **Add to `_LAYOUT_NAME_MAP`:**
   ```python
   "chart_slide": ["Blank"],
   ```

2. **Add `_CHART_TYPE_MAP`** (section 4.4 above).

3. **Add chart constants:**
   ```python
   _CHART_X = Inches(0.92)
   _CHART_Y = Inches(2.0)
   _CHART_CX = Inches(11.5)
   _CHART_CY = Inches(4.5)
   ```

4. **Add a new set for layouts with charts:**
   ```python
   _LAYOUTS_WITH_CHART = {"chart_slide"}
   ```

5. **Add `_add_chart_to_slide()` function** that:
   - Builds `CategoryChartData` from `categories` + `series`
   - Calls `slide.shapes.add_chart()`
   - Applies styling (legend, data labels, axis scale)
   - Returns success/failure

6. **In `generate_ppt_from_data()` main loop**, add a branch:
   ```python
   if slide_type in _LAYOUTS_WITH_CHART:
       _add_chart_to_slide(slide, slide_data)
   ```

7. **Title and notes** are already handled by existing code (TITLE placeholder + `_set_notes()`).

### 5.2 Integration with Sub-Issues

| Sub-Issue | Scope                                           | Implementation Approach                          |
| --------- | ----------------------------------------------- | ------------------------------------------------ |
| #2 (Bar)  | Bar/column chart rendering                      | Implement `_CHART_TYPE_MAP` bar entries + styling |
| #5 (Pie)  | Pie/doughnut chart rendering                    | Add percentage data label support               |
| #6 (Line) | Line chart rendering                            | Multi-series + markers support                  |
| #4 (Style)| Color palette, fonts, consistent visual design  | Hardcode palette for v1, theme extraction later  |
| #7 (Test) | Unit + integration tests                        | Test schema, chart creation, backward compat    |
| #8 (Docs) | SKILL.md documentation + examples               | Document schema, add examples                   |

Since we use a single `chart_slide` type, issues #2, #5, #6 can be implemented **in a single PR** (one `_add_chart_to_slide()` function with dispatch), rather than three separate code paths.

---

## 6. Prototype Validation

A prototype was built and validated (`research_prototype.py` → `output/prototype_charts.pptx`):

| Chart Type   | XL_CHART_TYPE     | Series Count | Data Labels | Legend | Notes | Result |
| ------------ | ----------------- | ------------ | ----------- | ------ | ----- | ------ |
| Bar          | COLUMN_CLUSTERED  | 1            | ✅          | ✅ Bottom | ✅ | ✅ Pass |
| Pie          | PIE               | 1            | ✅ (%)      | ✅ Right  | ✅ | ✅ Pass |
| Line         | LINE_MARKERS      | 3            | ✅          | ✅ Bottom | ✅ | ✅ Pass |

All charts verified as **native PowerPoint chart objects** (`shape.has_chart == True`), not images. Charts are fully editable in PowerPoint.

---

## 7. Open Decisions

### Decision 1: Single `chart_slide` vs. Separate Types ✅ Resolved
**Recommendation:** Single `chart_slide` type with `chart_type` field. (See section 4.1)

### Decision 2: Color Palette
For v1, use a hardcoded palette matching the template's blue/teal theme:
```python
_CHART_COLORS = [
    RGBColor(0x1F, 0x77, 0xB4),  # Blue
    RGBColor(0xFF, 0x7F, 0x0E),  # Orange
    RGBColor(0x2C, 0xA0, 0x2C),  # Green
    RGBColor(0xD6, 0x27, 0x28),  # Red
    RGBColor(0x94, 0x67, 0xBD),  # Purple
    RGBColor(0x8C, 0x56, 0x4B),  # Brown
]
```
Refinement deferred to issue #4 (Styling & Theme Integration).

### Decision 3: Error Handling
- Missing `chart_type` → default to `"bar"`
- Missing `categories` → skip chart, log warning
- Missing `series` → skip chart, log warning
- Invalid `chart_type` string → log warning, default to `"bar"`

---

## 8. Next Steps

1. ✅ Research complete — this document
2. → Implement `chart_slide` in `ppt_builder.py` (issues #2, #5, #6 — can be single PR)
3. → Apply styling (issue #4)
4. → Write tests (issue #7)
5. → Update documentation (issue #8)
6. → Generate the "Digital Technology in Construction" PPT with chart slides
