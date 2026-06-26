# Gap Analysis — chenyyu-user-stories vs Current Implementation

> **Document type:** Requirements-vs-implementation gap analysis (analysis only — no code changes)
> **Requirements source:** `docs/user-stories/chenyu-user-stories.md` (5 Epics, 19 Stories)
>
> *Note: the source document's own header states "17 Stories", but it actually contains 19 (Epic 1: 5, Epic 2: 2, Epic 3: 4, Epic 4: 5, Epic 5: 3). This report counts 19.*
> **Implementation audited:** `.opencode/skills/ppt-template-filler/`, `.opencode/skills/template-modifier-skill/`, `.opencode/agents/pptx-subagent.md`
> **Date:** June 2026 (Revision 2 — post-US-1.1)
>
> **Revision 2 (post-US-1.1, PR #49):** US-1.1 now ✅ Met (new `schema_extractor.py` emits the proposed-schema JSON). US-1.2 → 🟡 Partial (polygon field now exists + normalized, but cross-product winding check pending). US-1.3 improved (full 10-value enum on all elements). Counts updated: Met 2 / Partial 7 / Not met 9.
>
> **Revision 3 (post-US-1.2, PR #51):** US-1.2 now ✅ Met — cross-product winding check delivered (`_signed_area()` + check in `validate_template_schema()`; canonical TL→TR→BR→BL = algebraic CCW = anti-clockwise). Counts: Met 3 / Partial 6 / Not met 9.
>
> **Revision 4 (post-US-1.3, issue #52):** US-1.3 now ✅ Met — `type_confidence` is always emitted (`"high"` default, `"low"` only for `shape_type` None/unreadable and indeterminate MEDIA; no whitelist per architecture-review MAJOR-1), `"audio"` is now reachable via OOXML `<a:audioFile>`/`<a:videoFile>` split, `WEB_VIDEO`→`video/high`, and a non-fatal WARNING surfaces `shape/low` (MINOR-2). Optional `type_confidence` added to `template_schema.json`. Counts: Met 4 / Partial 5 / Not met 9.
>
> **Revision 5 (audit-path consistency):** Re-graded Epic 3/5 stories on the **same path used for Epic 1** (`schema_extractor.py`, the proposed-schema extractor) instead of the renderer contract. Corrections: **US-3.2** ❌→🟡 (`_infer_title` emits `title`; AC1 met, AC2/AC3 not), **US-3.4** 🟡→✅ (`_build_theme` maps semantic roles + `font_palette`; all 3 ACs met), **US-5.1** ❌→🟡 (`schema_extractor` is a real CLI with `--input/--output/--log-level` + exit codes 0/1/2; AC3 met, but skill names/decomposition differ), **US-5.2** 🟡 (clarified: `schema_version` IS in the proposed schema; `template_schema.json` exists but is not loaded at runtime), **US-5.3** 🟡 (clarified: `--log-level` DOES exist in `schema_extractor`; still not JSON-lines). Counts: Met 5 / Partial 6 / Not met 7.
>
> **Revision 6 (post-US-1.4, issue #54):** US-1.4 now ✅ Met — `_extract_text_fonts` populates per-textbox `font` (explicit-only, Latin-only) + nested `runs[]`, with a guarded RGB `color`; deduped `missing_fonts[]` against a curated `_BUILTIN_FONTS` allowlist (theme-aware `fallback`, AC4 → validator ERROR); `validate_template_schema` emits a non-fatal WARNING per missing font (AC3). Counts: Met 6 / Partial 6 / Not met 6.

---

## Executive Summary

This report compares `chenyu-user-stories.md` (the requirements document) against the current implementation. **Core conclusion: the architectural route has diverged.**

- The requirements document describes a **"normalized polygon JSON Schema embedded in PPTX"** architecture: each component contains a 4-point 0–1 normalized `polygon`, a `type` enum, `font`/`runs`, and JSON embedded inside the zip at `ppt/template_schema.json`.
- The current implementation follows a **"introspection contract with placeholder fingerprint matching"** architecture: placeholders use absolute inch coordinates, the contract is stored as a sidecar file `template.pptx.contract.json`, and layouts are matched by fingerprint.

Both achieve "fill any template", but they differ on **data model, skill decomposition, and artifact form**.

**Story-by-story summary** (19 stories): Met 6 / Partial 6 / Not met 6 / Architecture differs 1.

**The largest gaps** are concentrated in Epic 1 (normalized polygon component model + font detection + zip embedding), Epic 2/3 (header/footer detection + standalone template generator skill), and Epic 5 (skill decomposition into generate-template / generate-slides + CLI).

For the full comparison see §2; for remediation suggestions see §4.

---

## §1 Architecture Comparison

The requirements document and the implementation describe two different routes to the same goal ("fill any template"). They diverge on **data model, skill decomposition, artifact form, and layout-matching strategy**.

| Dimension | Requirements (chenyu-user-stories) | Current Implementation |
|---|---|---|
| **Skill decomposition** | Exactly 2 skills: `generate-template` (extraction) + `generate-slides` (generation), each a standalone CLI script | 2 skills + 1 agent: `ppt-template-filler` (filling) + `template-modifier-skill` (extension) + `pptx-subagent` (content strategy). No standalone CLI entry points. |
| **Template data model** | Normalized JSON Schema: each component has `polygon` (4 anti-clockwise 0–1 coords), `type` enum, `font`/`runs`, `content_template` with `{{placeholders}}` | Introspection contract: each placeholder has `idx/name/type/left_in/top_in/width_in/height_in` + composition `fingerprint` |
| **JSON storage** | **Embedded inside the PPTX zip** at `ppt/template_schema.json` (PowerPoint-safe) | **Sidecar file** `template.pptx.contract.json` next to the template (mtime-cached, gitignored) |
| **Layout matching** | Read embedded JSON → denormalize polygon coords back to EMU → place OOXML at exact positions | Placeholder-composition **fingerprint match** + name fallback via python-pptx `add_slide(layout)` (`ppt_builder.py:298`) |
| **Template generation output** | A downloadable PPTX with the embedded JSON (round-trip tested) | A sidecar contract; the original `.pptx` is never modified |

> **Revision 2 note (§1):** A new parallel module `schema_extractor.py` now **coexists** with `template_introspector.py` — it emits the **proposed-schema** JSON (normalized `polygon`, `type` enum, `components[]`) conforming to `schemas/template_schema.json`, separate from the renderer's fingerprint contract. The two data models are now "partially bridged"; the renderer still consumes only the contract. GAP §5 Decision 1 (Coexist) is now reality.

---

## §2 Story-by-Story Gap Analysis

Status legend: ✅ Met · 🟡 Partial · ❌ Not met · ⚪ Architecture differs (not directly applicable)

### Epic 1 — Template Extraction & JSON Schema

#### US-1.1 — Extract Slide Master to Structured JSON `[Must Have]` — ✅ Met

**Implemented (PR #49).** `schema_extractor.extract_schema()` (`schema_extractor.py`) reads any `.pptx`, parses the slide master (`prs.slide_masters[0]`) AND every layout, and emits a structured JSON conforming to `schemas/template_schema.json`. All four ACs are met: no crash on valid PPTX (`TemplateExtractionError` on bad input); master parsed + every layout enumerated; output is validated by a hand-rolled `validate_template_schema()` that mirrors `template_schema.json`'s rules (the file is a conformance-target spec, **not** loaded at runtime — see US-5.2); deterministic Python. The renderer's fingerprint contract (`template_introspector.py`) is untouched — the two modules coexist (§5 Decision 1, now reality). 49 tests pass (`test_schema_extractor.py`).

#### US-1.2 — Normalized Polygon Positioning `[Must Have]` — ✅ Met

**Met (PR #51).** All four ACs satisfied. `normalize_polygon()` emits exactly 4 normalized `{x,y}` points in `[0,1]` (AC1/AC2); slide dimensions in metadata (AC4). AC3 — the cross-product winding check — is delivered by `_signed_area()` + a check in `validate_template_schema()`: the canonical order TL→TR→BR→BL yields a **positive signed area**, which is algebraically counter-clockwise (CCW = anti-clockwise), exactly what AC3 asks a cross-product to verify. (Reversed winding → error; degenerate/zero-area → warning.) Note: in screen coords (Y-down) the trace visually appears clockwise, but the algebraic winding is CCW — documented in `template_schema.json` `$comment`. **Out of scope (Details, not ACs):** non-rectangular actual vertices (custGeom/triangle/connector) — polygon stays a 4-point rectangular bounding box; deferred (polygon is metadata-only, no consumer).

#### US-1.3 — Component Type Enumeration `[Must Have]` — ✅ Met

**Met (issue #52).** `schema_extractor._classify_shape()` (`schema_extractor.py`) applies the **full 10-value enum** to all elements (placeholders and freeform shapes). All three ACs are met (type always present; unknowns degrade to `shape`, never null/unknown; OOXML→enum mapping in source). The two previously-deferred **Details** are now delivered: (1) `type_confidence` is **always emitted** (`"high"` default; `"low"` only when `shape_type` is `None`/unreadable or MEDIA is indeterminate — no whitelist, per architecture-review MAJOR-1, so recognized-but-unmapped members like `LINKED_PICTURE`/`TEXT_EFFECT`/`CALLOUT` stay `"high"`); (2) the `"audio"` enum value is **reachable** via OOXML `<a:audioFile>`/`<a:videoFile>` split of `MSO_SHAPE_TYPE.MEDIA`, and `WEB_VIDEO`→`video/high`. A non-fatal `ValidationIssue` WARNING surfaces `shape/low` ("flagged for review"). Optional `type_confidence` added to `schemas/template_schema.json`. 64 tests in `test_schema_extractor.py` (was 49).

#### US-1.4 — Font Detection & Availability Checking `[Must Have]` — ✅ Met

**Met (issue #54).** `schema_extractor._extract_text_fonts()` (`schema_extractor.py`) populates every text-bearing component's `font` (`family`, `size_pt`, `weight`, `color`, `alignment`, `is_available`, `fallback`) — **explicit-only** (inherited values → `null`); Latin/English only (CJK `<a:ea>`/`<a:cs>` out of scope). It captures a nested `runs[]` (`{text, font:{...}}`) and a guarded RGB `color` (`color.type == MSO_COLOR_TYPE.RGB`, else `null`). The deck aggregates a deduped top-level `missing_fonts[]` (`{family, is_available:false, fallback, download_url:null}`) against a curated `_BUILTIN_FONTS` allowlist; `fallback` defaults to the theme body font (if built-in) else `Arial`, always a built-in name (AC4 → validator ERROR). `validate_template_schema` emits a non-fatal `ValidationIssue(severity="warning")` per missing font (AC3); `extract_schema` also `logger.warning`s. All 4 ACs met. 78 tests in `test_schema_extractor.py` (was 65).

#### US-1.5 — JSON Storage Inside PPTX Zip `[Must Have]` — ❌ Not met

Still not met. `schema_extractor` (PR #49) outputs JSON (CLI `--output` or stdout) and authors `schemas/template_schema.json` as the **spec**, but it does **not embed** the JSON into the PPTX zip. The renderer still reads the sidecar `<stem>.pptx.contract.json` (`template_introspector.py:271-275`). No PowerPoint-safe zip-append logic, no file-size logging.

### Epic 2 — Header, Footer & Best Practices

#### US-2.1 — Header & Footer Detection `[Must Have]` — ❌ Not met

`_CHROME_TYPES` (`template_introspector.py:44-49`) **recognizes** HEADER/FOOTER/SLIDE_NUMBER/DATE, but only to **filter them out** as layout noise (`placeholder_record` returns `None` for chrome, `template_introspector.py:168-169`). There is **no `header_footer.has_header/has_footer` metadata**, no recorded component IDs, and **no user prompt** when both are absent.

#### US-2.2 — Common Practice Suggestions `[Should Have]` — ❌ Not met

No `common_practices` object exists in the contract. There is **no checks** for slide numbers, company logo, consistent margins, section dividers, or closing slide, and no `suggestions` array.

### Epic 3 — Skill — Template Generator

#### US-3.1 — End-to-End Template Generation Pipeline `[Must Have]` — ❌ Not met

Introspection is **embedded in the render path** (runs automatically before `generate_ppt_from_data`), not a standalone `generate-template` skill. There is **no** "extract → validate → embed → return templated PPTX" CLI pipeline.

#### US-3.2 — Template Naming `[Must Have]` — 🟡 Partial

The renderer's fingerprint contract (`template_introspector.py`) has only `source_file`/`source_mtime` — no `title`. **But the proposed-schema path implements it**: `schema_extractor._infer_title()` infers `template_metadata.title` (`core.xml` title → first-slide title text → filename stem) and `_build_metadata()` emits it. AC1 (title always non-empty — the filename fallback guarantees it) is **met**; AC2 (inference order ending in a **user prompt**) and AC3 (displayed for confirmation) are **not** (the third fallback is the filename, not a prompt; nothing is surfaced for confirmation). Graded Partial on the `schema_extractor` path.

#### US-3.3 — Return Downloadable Templated PPTX `[Must Have]` — ❌ Not met

No PPTX-with-embedded-JSON is produced or returned. The original `.pptx` is **never modified** (the sidecar is the only artifact). No round-trip test exists.

#### US-3.4 — Theme & Color Extraction `[Should Have]` — ✅ Met

The renderer's `_build_theme()` (`template_introspector.py`) extracts only raw OOXML role colors. **But the proposed-schema path fully implements it**: `schema_extractor._build_theme()` maps raw colors to semantic roles (`primary_color`/`secondary_color`/`accent_color`/`background_color`/`text_color`) and builds `font_palette.{heading,body,accent}`; `_raw_theme_colors_and_fonts()` is wrapped in try/except that logs a warning and yields empty defaults on a missing/malformed theme. All three ACs (semantic colors as hex; `font_palette`; sensible defaults + warning) are met on the `schema_extractor` path.

### Epic 4 — Skill — Slide Generator

#### US-4.1 — Read Embedded JSON as Layout Reference `[Must Have]` — 🟡 Partial

The engine reads the **sidecar contract** (not zip-embedded JSON) via `get_contract`, then resolves layouts by fingerprint (`_resolve_layout_by_fingerprint`, `ppt_builder.py:298`). So it does "read a contract as layout reference" — but the contract is a sidecar, and matching is fingerprint-based rather than polygon-denormalization. The "<1% EMU accuracy" criterion does not apply (it never denormalizes polygons).

#### US-4.2 — Visually Pleasing Output with Text Fitting `[Must Have]` — 🟡 Partial

Overflow prevention is **preventive, not reactive**: `density_mode.py` enforces a per-slide word budget (`DENSITY_BUDGETS`, `density_mode.py:48`) and `constraint_checker` checks `content_area_in2`. But there is **no reactive per-textbox font auto-shrink** (−2pt steps until fit) — a grep for `font_size`/`auto.*shrink`/`adjust` in `density_mode.py` returned no matches. No `font_size_adjusted` flag is set.

#### US-4.3 — Auto-Chain Extraction When No JSON Present `[Must Have]` — ⚪ Architecture differs

Because the model is contract-based (introspection runs automatically before every render) rather than embedded-JSON-based, there is no concept of "no JSON → extract first". The two-step chain (`extract` then `generate`) does not map; introspection is always implicit. Functionally the user can hand any `.pptx` and generate from it in one step — but the *mechanism* differs from the story.

#### US-4.4 — Style Picker for Template-Less Generation `[Should Have]` — ❌ Not met

A template is **required** (`template.pptx` must exist or the engine raises `FileNotFoundError`). There are **no built-in style presets** (minimalist/corporate/creative/dark) and no style picker. A default template exists, but no preset-selection flow.

#### US-4.5 — Multi-Slide Batch Generation `[Could Have]` — ✅ Met

The multi-stage pipeline (outline → critique → detail → render) generates 2–20+ slides from one prompt, the outline is shown for user approval, and the stages act as a progress indicator. Fully satisfied.

### Epic 5 — Skill Architecture & Scripts

#### US-5.1 — Two Independent Skills with CLI Scripts `[Must Have]` — 🟡 Partial

The skills are `ppt-template-filler` (fill) and `template-modifier-skill` (extend) — **not** the `generate-template`/`generate-slides` decomposition the story names, and the manifests are `SKILL.md` (no `skill.yaml`) → AC1 not met. **But a standalone CLI does exist**: `schema_extractor.py` `main()` is a real CLI (`--input/-i`, `--output/-o`, `--log-level`) with documented exit codes 0/1/2 (success/validation/runtime) → AC3 (exit codes) met and AC2 (runnable from CLI independently of the LLM) met *for the extractor*. `ppt_builder.py` remains function-based (its `main()` is a demo, no argparse). Graded Partial: one of the two intended skills exists as a CLI with exit codes, but the skill decomposition/naming differs and the slide side is not a CLI.

#### US-5.2 — Shared JSON Schema for Validation `[Must Have]` — 🟡 Partial

`slide_schemas.py` validates **slide content** for all 8 slide types + `chart_options` (`schema_validator.py` `validate_slide_data_list` / `parse_and_validate`). The **extraction** side: `schemas/template_schema.json` (JSON Schema draft-2020-12) exists as the conformance spec and the proposed schema **does** carry `schema_version` (`SCHEMA_VERSION = "1.0.0"`, emitted in `_build_metadata`). However, the schema file is **not loaded at runtime** — `validate_template_schema()` is hand-rolled (no `jsonschema` dependency), so spec and validator are kept in sync manually (a pre-existing divergence: `additionalProperties:false` and the `id` `pattern` in the schema are not enforced). Graded Partial: the spec + version exist, but runtime validation is hand-rolled, not schema-driven.

#### US-5.3 — Structured Logging `[Should Have]` — 🟡 Partial

Python's `logging` module is used across the modules. `schema_extractor.py` **does** expose a `--log-level` flag (debug/info/warn/error) applied via `logging.basicConfig`. But logging is **not JSON-lines structured** (plain `%(asctime)s [%(levelname)s] %(message)s` format), output is not explicitly routed to stderr-only, and only `schema_extractor` exposes the flag (the engine modules don't). Graded Partial: the flag exists for the extractor, but structured JSON-lines + stderr-only routing are not implemented.

---

## §3 Gap Summary

### §3.1 Statistics

| Status | Count | Stories |
|---|---|---|
| ✅ Met | 6 | US-1.1, US-1.2, US-1.3, US-1.4, US-3.4, US-4.5 |
| 🟡 Partial | 6 | US-3.2, US-4.1, US-4.2, US-5.1, US-5.2, US-5.3 |
| ❌ Not met | 6 | US-1.5, US-2.1, US-2.2, US-3.1, US-3.3, US-4.4 |
| ⚪ Architecture differs | 1 | US-4.3 |

### §3.2 Priority × Status Matrix

| | Must Have | Should Have | Could Have |
|---|---|---|---|
| ✅ Met | US-1.1, US-1.2, US-1.3, US-1.4 | US-3.4 | US-4.5 |
| 🟡 Partial | US-3.2, US-4.1, US-4.2, US-5.1, US-5.2 | US-5.3 | — |
| ❌ Not met | **US-1.5, US-2.1, US-3.1, US-3.3** | US-2.2, US-4.4 | — |
| ⚪ Differs | US-4.3 | — | — |

**Highest-risk gaps** are the 4 unmet **Must-Have** stories (US-1.4 now Met in Rev 6), clustered in Epic 1 (zip), Epic 2 (header/footer), Epic 3 (template generator), and Epic 5 (skill decomposition).

---

## §4 Remediation Suggestions

These are **suggestions only** — no code is written here. They are grouped by priority. A decision on whether to (a) build the requirements architecture alongside the current one, (b) replace it, or (c) keep the current architecture and update the requirements, is deferred to §5.

### P0 — Core Extraction Model (Must-Have, highest impact)

These four stories are the foundation everything else builds on. They define the normalized component schema the requirements center on.

| Story | Suggestion |
|---|---|
| ~~**US-1.2 Polygon**~~ | **Done (US-1.2):** `polygon` field + 4 normalized 0–1 coords + the cross-product winding check (AC3) are all delivered — see §2 (Met). *Deferred Details (not ACs):* non-rectangular actual vertices (custGeom/triangle) — polygon stays a rectangular bounding box (metadata-only, no consumer). |
| ~~**US-1.4 Fonts**~~ | **Done (issue #54):** `_extract_text_fonts` populates per-textbox `font` (explicit-only, Latin-only) + nested `runs[]`; deduped `missing_fonts[]` against `_BUILTIN_FONTS` with theme-aware `fallback` (AC4 → validator ERROR); `validate_template_schema` warns per missing font (AC3). RGB `color` guarded by `color.type==RGB`. *Out of scope:* full-cascade inheritance, CJK (`<a:ea>`/`<a:cs>`), theme-color → hex. |
| **US-1.5 Zip embedding** | Add a `embed_schema(pptx_path, schema)` function that opens the zip, appends `ppt/template_schema.json` (minified), and writes a new zip **without touching `[Content_Types].xml` or any existing entry**. Verify PowerPoint opens it without repair. |
| **US-2.1 Header/Footer** | Stop discarding chrome: record `has_header/has_footer` booleans + component IDs in `header_footer` metadata. Emit a user prompt when both are absent. |

### P1 — Template Generator Skill & Decomposition (Must-Have)

| Story | Suggestion |
|---|---|
| **US-3.1 + US-5.1** | Introduce a dedicated `generate-template` skill with a real CLI (`--input/--output`) that runs the full P0 extraction → validation → zip-embed → return. Mirror it as `generate-slides` for the read-JSON-then-render path. Add documented exit codes (0 success / 1 validation / 2 runtime). |
| **US-3.2** | Infer `title` from `docProps/core.xml` → first-slide title → user prompt, and store it as `template_metadata.title`. |
| **US-3.3** | Produce the downloadable templated PPTX (P0 embedding) + a human-readable extraction summary. Add a round-trip test (open → re-upload → re-extract → JSON identical). |

### P2 — Polish & Best Practices (Should/Could-Have)

| Story | Suggestion |
|---|---|
| ~~**US-1.3**~~ | **Done (issue #52):** `type_confidence` now always emitted (`"high"` default; `"low"` only for `shape_type` None/unreadable or indeterminate MEDIA — no whitelist, so recognized-unmapped members stay `"high"`); `"audio"` reachable via `<a:audioFile>`/`<a:videoFile>` split; `WEB_VIDEO`→`video/high`; `shape/low` surfaces a non-fatal WARNING. Optional `type_confidence` added to `template_schema.json`. *(The original suggestion "extend extraction beyond placeholders" was already done in US-1.1/PR #49.)* |
| **US-2.2** | Add a `common_practices` checker (5+ practices: slide numbers, logo, margins, section dividers, closing slide) emitting a `suggestions` array. |
| ~~**US-3.4**~~ | **Done (Rev 5):** `schema_extractor._build_theme()` maps semantic roles (`primary/secondary/accent/background/text_color`) + builds `font_palette.{heading,body,accent}`; missing-theme → empty defaults + warning. All 3 ACs met. *(The renderer-side `_build_theme` still emits raw role colors only; that is US-5.2/migration scope.)* |
| **US-4.4** | Ship 4+ built-in style presets (minimalist/corporate/creative/dark) as JSON schemas and add a style-picker prompt when no template is provided. |
| **US-5.2** | Load `template_schema.json` at runtime (or generate the validator from it) so `additionalProperties:false`/`pattern` are enforced, not just hand-rolled; bridge the spec↔validator divergence. (`schema_version` and the spec file already exist.) |
| **US-5.3** | Switch logging to structured JSON-lines (timestamp/level/skill/action/details) routed to stderr. *(The `--log-level` flag already exists in `schema_extractor.py`; extend it to the engine modules.)* |

### Partial stories already mostly satisfied

~~US-1.1~~ (now Met), US-4.1, US-4.2, US-4.5, US-5.2 (content side), US-5.3 — these need **field-set additions or mode changes**, not greenfield work. The cheapest wins are here.

---

## §5 Open Decisions

Before any build work, these route questions need a decision:

**1. Coexist / Replace / Update-requirements**
- (a) **Coexist**: build the requirements architecture (polygon schema + zip embed + generate-template/generate-slides skills) *alongside* the current fingerprint-contract engine. Most work, lowest disruption.
- (b) **Replace**: migrate the current engine onto the normalized-schema model and deprecate the sidecar contract. Highest consistency, highest risk.
- (c) **Update requirements**: judge the current fingerprint architecture as the better-evolved path and rewrite `chenyu-user-stories.md` to match it, documenting the divergence rationale. Least work; changes the contract.

> **Status (Revision 2):** Decision 1 is now **(a) Coexist — implemented**. `schema_extractor.py` (PR #49) coexists with the fingerprint contract; the renderer still consumes only the contract. The full migration (deprecating the introspector / bridging the polygon model into rendering) remains open via Decision 2.

**2. Who consumes the polygon schema**
The requirements assume the slide generator denormalizes polygons back to EMU and places OOXML at exact positions. The current engine delegates positioning to python-pptx `add_slide(layout)` (the layout's own placeholders). If the polygon model is built, does the renderer switch to manual coordinate placement, or do polygons stay metadata-only?

**3. Priority confirmation**
Are the 8 unmet Must-Have stories still Must-Have, or has the fingerprint-contract path made some of them obsolete in practice (e.g. US-4.3 is arguably superseded by automatic introspection)?

---

*End of analysis. This document contains no code changes; it records the current state and suggests directions for a future decision.*
