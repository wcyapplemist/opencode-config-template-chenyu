# Project-Specific Agent Instructions

## Project Overview

PPTX subagent development — iterating and testing the `pptx-subagent` and `ppt-template-filler` skill.

## Project Structure

```
pptx-subagent-development/
├── .opencode/
│   ├── agents/
│   │   └── pptx-subagent.md       # Project-level PPT subagent
│   └── skills/
│       └── ppt-template-filler/   # Template filling engine + SKILL.md
├── output/                        # Generated .pptx files
└── AGENTS.md                      # This file
```

## Project-Level Resources

| Resource              | Type  | Scope             |
| --------------------- | ----- | ----------------- |
| `pptx-subagent`       | Agent | This project only |
| `ppt-template-filler` | Skill | This project only |

Global subagents and skills are managed at `~/.config/opencode/` and are available in all projects.

## Development Notes

- The `pptx-subagent` uses `ppt_builder.py` from the `ppt-template-filler` skill to populate `template.pptx` layouts
- Generated files are saved to `output/`
- The subagent is STRICTLY FORBIDDEN from building PPTX files from scratch
