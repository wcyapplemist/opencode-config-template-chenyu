# PPTX Subagent Development

Development workspace for the `pptx-subagent` and `ppt-template-filler` skill used with OpenCode.

## Overview

This project contains a project-level OpenCode subagent that generates PowerPoint presentations by filling a `template.pptx` Slide Master, rather than building slides from scratch.

## Components

### pptx-subagent (`.opencode/agents/pptx-subagent.md`)

A specialized PPT Content Strategist that:

- Transforms user requests into structured presentation content
- Delegates to the `ppt-template-filler` skill for `.pptx` generation
- Is forbidden from raw `python-pptx` construction

### ppt-template-filler (`.opencode/skills/ppt-template-filler/`)

The underlying engine:

- Loads `template.pptx` with named Slide Master layouts
- Removes example slides, adds new slides from layouts
- Fills placeholders by type (TITLE, SUBTITLE, OBJECT)
- Embeds **native charts** (editable, not images) and **native pictures**
- Outputs to `output/` directory

### Phase 1: Content Intelligence & Resource Resolution

The engine includes a content-intelligence layer on top of the renderer:

- **Schema validation (#20)** — every `slide_data_list` is validated against explicit JSON schemas (8 slide types + `chart_options`) with a two-layer retry wrapper for LLM JSON. Structured errors (slide index + field path) let the agent self-correct.
- **Resource resolution pipeline (#19/#18/#23)** — emit placeholders (`image_prompt`, `icon_query`, `data_query`); an independent resolver pass replaces them with real assets (stock photos, semantic icons, sourced chart data) before rendering. All resolution is non-fatal.
- **Native image embedding (#18)** — `image_path` inserts an editable PowerPoint picture (placeholder fill or named presets).
- **Multi-stage generation (#21/#24)** — outline → critique → detail, schema-gated per stage; autonomous by default with an optional interactive outline checkpoint in primary-agent mode.

Design docs live in `.opencode/skills/ppt-template-filler/docs/`.

## Usage

In any conversation, trigger the subagent with phrases like:

- "Create a presentation about..."
- "Generate a PowerPoint deck for..."
- "Make a .pptx with..."

## Project-Level vs Global

This project defines resources scoped to this repository only:

| Resource              | Location            | Global? |
| --------------------- | ------------------- | ------- |
| `pptx-subagent`       | `.opencode/agents/` | No      |
| `ppt-template-filler` | `.opencode/skills/` | No      |

Global agents (30 subagents) and skills (54+) are managed at `C:\Users\LENOVO\.config\opencode\`.
