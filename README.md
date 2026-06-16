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
- Outputs to `output/` directory

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
