# Design: Placeholder Convention & Resource Resolver Architecture

> **Issue #19** — Phase 1, Track B. Design rationale for the placeholder/resolver
> decoupling. Implementation lives in `scripts/resolvers/`. The user-facing
> placeholder table and pipeline order are also documented in `SKILL.md` — this
> doc adds the interface contract, injectable hooks, and degradation matrix.

## 1. Problem

Today the agent (LLM) is asked to produce *concrete* assets: a real image URL,
a chart's exact numbers, an icon file. This is fragile:

- The LLM hallucinates URLs that 404 or fabricates statistics.
- The rendering layer (`ppt_builder.py`) must never touch the network or guess
  asset formats.
- Mixing "content authoring" with "asset acquisition" in one model call degrades
  both.

## 2. Core Idea — Placeholder Decoupling

Inspired by presenton's `utils/process_slides.py`, we **split authoring from
acquisition**:

```
Agent (LLM)                 Resolver (code)              Engine (ppt_builder)
─────────────────────────────────────────────────────────────────────────────
emits placeholders    →     resolves placeholders   →   renders concrete values
```

The agent **only emits placeholders** describing what it wants. An independent,
deterministic **resolver pass** walks the deck, finds placeholders, and replaces
them with concrete values (`image_path`, `icon_path`, populated
`categories`/`series`). Only the resolved JSON reaches `generate_ppt_from_data()`.
The rendering layer stays untouched and network-free.

## 3. Placeholder Convention

All placeholders are **optional** fields on a slide dict. A slide may carry any
combination. When a placeholder is present and unresolved, the resolver fills it
in place; the original placeholder key is then removed or left harmless.

### 3.1 Image

```json
{
  "image_prompt": "modern construction site with drones at sunset",
  "image_source": "auto",
  "image_query": "construction drone aerial"
}
```

| Field           | Required | Values              | Notes                                              |
|-----------------|----------|---------------------|----------------------------------------------------|
| `image_prompt`  | No       | string              | Free-text description; preferred input.            |
| `image_source`  | No       | `auto`/`stock`/`ai` | `stock` = photo API; `ai` = generative model.      |
| `image_query`   | No       | string              | Explicit search query (overrides prompt).          |

**Resolved output:** `image_path` (a local file path) consumed by `_add_image_to_slide()` (#18).

### 3.2 Icon

```json
{
  "icon_query": "growth trend up arrow"
}
```

**Resolved output:** `icon_path` (local SVG/PNG). Icons are inserted as native,
editable pictures.

### 3.3 Real-data chart

```json
{
  "slide_type": "chart_slide",
  "chart_type": "bar",
  "data_query": "global construction-tech market size 2020-2026 in USD billion",
  "data_hint": {"categories": ["years"], "series": [{"name": "Market Size"}]}
}
```

**Resolved output:** populated `categories` and `series` arrays with **sourced**
numbers, plus a `data_source` string (URL/citation) appended to the slide notes.
If the slide already carries concrete `categories`/`series`, the resolver does
**not** overwrite them.

## 4. Resolver Interface

Every resolver is a **module-level function** with a uniform signature (not a
class — the original design proposed a `Resolver` class, but it was simplified
to plain functions during implementation):

```python
from typing import Any, Dict

def resolve(slide_data: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Any]:
    """Return a NEW slide_data dict with this resolver's placeholders resolved.

    * Non-fatal: on any failure, log a warning and return the slide
      UNCHANGED. Never raise.
    * Idempotent: a slide with no placeholders (or already-resolved ones)
      is returned unchanged.
    """
```

Three resolvers are implemented: `image_resolver.resolve()`,
`icon_resolver.resolve()`, `chart_data_resolver.resolve()`. They are independent
and may be added/removed without touching the others.

### 4.1 Injectable hooks (for testability)

Each resolver accepts optional **injectable callable hooks** via the `config`
dict, so tests can substitute deterministic fakes without monkeypatching:

| Hook          | Used by            | Signature                             |
|---------------|--------------------|---------------------------------------|
| `fetch_fn`    | `image_resolver`   | `fetch_fn(query, source) -> bytes`    |
| `match_fn`    | `icon_resolver`    | `match_fn(query) -> Optional[str]`     |
| `search_fn`   | `chart_data_resolver` | `search_fn(query, hint) -> dict`   |

When a hook is absent from `config`, the resolver falls back to its real
provider (stock photo API, local icon library, web search). This makes the
resolver pipeline fully testable without network access.

## 5. Pipeline

`resolve_slide_data_list(slide_data_list, config) -> slide_data_list` orchestrates:

```
for each slide:
    for resolver in [image, icon, chart_data]:
        slide = resolver.resolve(slide, config)
return slides
```

Order is **image → icon → chart_data** (chart-data last so it can read a stable
`data_source` if needed). Each resolver mutates only its own placeholder keys.

### 5.1 End-to-end pipeline order (Phase 1)

```
1. Agent emits slide_data_list (with placeholders)
2. resolve_slide_data_list() replaces placeholders with concrete assets
3. validate_slide_data_list(strict=True) gates the resolved deck   (#20)
4. generate_ppt_from_data() renders the validated deck              (#18)
```

## 6. Configuration

A single `resolver.config.json` (gitignored) holds provider keys and selection.
A checked-in `resolver.config.example.json` documents the schema:

```json
{
  "image": {"provider": "pexels", "api_key_env": "PEXELS_API_KEY"},
  "icon":  {"library": "phosphor", "path": "assets/icons"},
  "chart_data": {"search": "webfetch"}
}
```

If a provider is unconfigured, the matching resolver logs a warning and skips
(never aborts the build).

## 7. Graceful Degradation (contract)

| Situation                       | Behavior                                              |
|---------------------------------|-------------------------------------------------------|
| Placeholder present, no config  | Warning, slide renders without that asset             |
| Provider returns no result      | Warning, slide renders without that asset             |
| Network error                   | Warning, slide renders without that asset             |
| Resolver throws                 | Caught, warning, slide renders without that asset     |

**A failed resolution NEVER crashes the build** — the deck is always produced;
only the affected asset is omitted.

## 8. What this does NOT do

- Does not change the rendering paradigm (still native `python-pptx`).
- Does not embed linked images — resolved assets are local & embedded.
- Does not require network at render time — resolution is a separate pre-pass.
