# Proposed JSON Schema — `template_schema.json`

> Extracted from `chenyu-user-stories.md`.

The canonical structure stored at `ppt/template_schema.json` inside the PPTX zip.

```json
{
  "template_metadata": {
    "title": "Q3 Investor Pitch Deck — Dark Theme",
    "schema_version": "1.0.0",
    "generated_by": "opencode-pptx-subagent/generate-template",
    "generated_at": "2025-06-15T10:30:00Z",
    "slide_dimensions": {
      "width_emu": 12192000,
      "height_emu": 6858000,
      "width_inches": 13.333,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "missing_fonts": [
      {
        "family": "Neue Haas Grotesk",
        "is_available": false,
        "fallback": "Arial",
        "download_url": "https://fonts.adobe.com/fonts/neue-haas-grotesk"
      }
    ],
    "header_footer": {
      "has_header": true,
      "has_footer": true,
      "header_components": ["comp_010"],
      "footer_components": ["comp_011", "comp_012"],
      "suggestions": []
    }
  },
  "theme": {
    "primary_color": "#1A1A2E",
    "secondary_color": "#16213E",
    "accent_color": "#E8A838",
    "background_color": "#0F0F1A",
    "text_color": "#E8ECF4",
    "font_palette": {
      "heading": "Neue Haas Grotesk",
      "body": "Calibri",
      "accent": "Calibri Light"
    }
  },
  "slide_layouts": [
    {
      "layout_id": "title_slide",
      "layout_name": "Title Slide",
      "layout_index": 0,
      "components": [
        {
          "id": "comp_001",
          "type": "textbox",
          "name": "Main Title",
          "placeholder_type": "title",
          "polygon": [
            { "x": 0.08, "y": 0.28 },
            { "x": 0.92, "y": 0.28 },
            { "x": 0.92, "y": 0.48 },
            { "x": 0.08, "y": 0.48 }
          ],
          "font": {
            "family": "Neue Haas Grotesk",
            "size_pt": 44,
            "weight": "bold",
            "color": "#E8ECF4",
            "alignment": "left",
            "is_available": false,
            "fallback": "Arial"
          },
          "runs": [
            {
              "text": "{{slide_title}}",
              "font": { "family": "Neue Haas Grotesk", "size_pt": 44, "weight": "bold", "color": "#E8ECF4" }
            }
          ],
          "z_order": 1,
          "content_template": "{{slide_title}}"
        },
        {
          "id": "comp_002",
          "type": "textbox",
          "name": "Subtitle",
          "placeholder_type": "subtitle",
          "polygon": [
            { "x": 0.08, "y": 0.52 },
            { "x": 0.70, "y": 0.52 },
            { "x": 0.70, "y": 0.62 },
            { "x": 0.08, "y": 0.62 }
          ],
          "font": {
            "family": "Calibri",
            "size_pt": 20,
            "weight": "regular",
            "color": "#7A8399",
            "alignment": "left",
            "is_available": true,
            "fallback": null
          },
          "runs": [
            {
              "text": "{{slide_subtitle}}",
              "font": { "family": "Calibri", "size_pt": 20, "weight": "regular", "color": "#7A8399" }
            }
          ],
          "z_order": 2,
          "content_template": "{{slide_subtitle}}"
        },
        {
          "id": "comp_003",
          "type": "image",
          "name": "Background Image",
          "placeholder_type": null,
          "polygon": [
            { "x": 0.0, "y": 0.0 },
            { "x": 1.0, "y": 0.0 },
            { "x": 1.0, "y": 1.0 },
            { "x": 0.0, "y": 1.0 }
          ],
          "z_order": 0,
          "content_template": null,
          "image_properties": {
            "embed_id": "rId1",
            "is_background": true
          }
        },
        {
          "id": "comp_004",
          "type": "shape",
          "name": "Accent Line",
          "placeholder_type": null,
          "polygon": [
            { "x": 0.08, "y": 0.24 },
            { "x": 0.35, "y": 0.24 },
            { "x": 0.35, "y": 0.255 },
            { "x": 0.08, "y": 0.255 }
          ],
          "z_order": 3,
          "shape_properties": {
            "preset": "rect",
            "fill": { "type": "solid", "color": "#E8A838" },
            "line": null
          }
        }
      ]
    },
    {
      "layout_id": "content_slide",
      "layout_name": "Content Slide",
      "layout_index": 1,
      "components": [
        {
          "id": "comp_005",
          "type": "textbox",
          "name": "Slide Title",
          "placeholder_type": "title",
          "polygon": [
            { "x": 0.08, "y": 0.06 },
            { "x": 0.92, "y": 0.06 },
            { "x": 0.92, "y": 0.16 },
            { "x": 0.08, "y": 0.16 }
          ],
          "font": {
            "family": "Neue Haas Grotesk",
            "size_pt": 32,
            "weight": "bold",
            "color": "#E8ECF4",
            "alignment": "left",
            "is_available": false,
            "fallback": "Arial"
          },
          "z_order": 1,
          "content_template": "{{slide_title}}"
        },
        {
          "id": "comp_006",
          "type": "textbox",
          "name": "Body Content",
          "placeholder_type": "body",
          "polygon": [
            { "x": 0.08, "y": 0.20 },
            { "x": 0.60, "y": 0.20 },
            { "x": 0.60, "y": 0.85 },
            { "x": 0.08, "y": 0.85 }
          ],
          "font": {
            "family": "Calibri",
            "size_pt": 16,
            "weight": "regular",
            "color": "#C0C8D8",
            "alignment": "left",
            "is_available": true,
            "fallback": null
          },
          "z_order": 2,
          "content_template": "{{slide_body}}",
          "text_properties": {
            "bullet_style": "disc",
            "line_spacing_pt": 24,
            "space_after_pt": 8
          }
        },
        {
          "id": "comp_007",
          "type": "image",
          "name": "Right Panel Image",
          "placeholder_type": "picture",
          "polygon": [
            { "x": 0.65, "y": 0.20 },
            { "x": 0.92, "y": 0.20 },
            { "x": 0.92, "y": 0.85 },
            { "x": 0.65, "y": 0.85 }
          ],
          "z_order": 3,
          "content_template": "{{slide_image}}",
          "image_properties": {
            "embed_id": null,
            "is_background": false
          }
        }
      ]
    },
    {
      "layout_id": "two_column_slide",
      "layout_name": "Two Column",
      "layout_index": 2,
      "components": [
        {
          "id": "comp_008",
          "type": "textbox",
          "name": "Slide Title",
          "placeholder_type": "title",
          "polygon": [
            { "x": 0.08, "y": 0.06 },
            { "x": 0.92, "y": 0.06 },
            { "x": 0.92, "y": 0.16 },
            { "x": 0.08, "y": 0.16 }
          ],
          "font": { "family": "Neue Haas Grotesk", "size_pt": 32, "weight": "bold", "color": "#E8ECF4", "alignment": "left", "is_available": false, "fallback": "Arial" },
          "z_order": 1,
          "content_template": "{{slide_title}}"
        },
        {
          "id": "comp_009",
          "type": "textbox",
          "name": "Left Column",
          "placeholder_type": "body",
          "polygon": [
            { "x": 0.08, "y": 0.20 },
            { "x": 0.48, "y": 0.20 },
            { "x": 0.48, "y": 0.85 },
            { "x": 0.08, "y": 0.85 }
          ],
          "font": { "family": "Calibri", "size_pt": 15, "weight": "regular", "color": "#C0C8D8", "alignment": "left", "is_available": true, "fallback": null },
          "z_order": 2,
          "content_template": "{{column_left}}"
        },
        {
          "id": "comp_009b",
          "type": "textbox",
          "name": "Right Column",
          "placeholder_type": "body",
          "polygon": [
            { "x": 0.52, "y": 0.20 },
            { "x": 0.92, "y": 0.20 },
            { "x": 0.92, "y": 0.85 },
            { "x": 0.52, "y": 0.85 }
          ],
          "font": { "family": "Calibri", "size_pt": 15, "weight": "regular", "color": "#C0C8D8", "alignment": "left", "is_available": true, "fallback": null },
          "z_order": 3,
          "content_template": "{{column_right}}"
        }
      ]
    }
  ],
  "component_type_enum": [
    "textbox",
    "image",
    "table",
    "video",
    "shape",
    "chart",
    "group",
    "smartart",
    "placeholder",
    "audio"
  ],
  "placeholder_type_enum": [
    "title",
    "subtitle",
    "body",
    "picture",
    "chart",
    "table",
    "media",
    "date",
    "slide_number",
    "footer",
    "header",
    null
  ]
}
```

## Key Design Decisions

1. **Normalized coordinates (0.0–1.0):** Makes the schema resolution-independent. The slide dimensions in metadata allow denormalization to EMUs for OOXML output.

2. **Anti-clockwise polygon winding:** Consistent with mathematical convention and allows cross-product-based normal/area calculations for text-fitting heuristics.

3. **`content_template` with `{{placeholder}}` syntax:** Lets the slide generator do simple string substitution without LLM involvement for straightforward fills, while the LLM handles content generation for the placeholders.

4. **`runs` array for font runs:** PowerPoint textboxes often contain mixed formatting (bold title + regular body in one box). Capturing runs preserves this nuance.

5. **`missing_fonts` at top level:** Centralized list means the user sees all font issues in one place, rather than scattered across individual component font objects.

6. **JSON inside zip, not in `[Content_Types].xml`:** PowerPoint silently ignores files it doesn't know about. This is the safest embedding strategy — zero risk of corrupting the file.

7. **`component_type_enum` and `placeholder_type_enum` in the JSON itself:** Self-documenting. Any consumer of the JSON can enumerate valid types without external documentation.
