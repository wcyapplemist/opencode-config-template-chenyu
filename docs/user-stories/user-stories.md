# PPTX Subagent — User Stories

> **Project:** opencode.ai / pptx-subagent
> **Date:** June 2025
> **Scope:** 4 Epics, 20 Stories
> **Author:** Founder (OpenCode User)
> **Source:** `chenyu-user requirement.html` (`.md` export) — pure requirements, stripped of HTML/CSS/JS boilerplate.

---

## Table of Contents

- [Epic 1: Template Extraction & Templating](#epic-1-template-extraction--templating)
- [Epic 2: Slide Generation](#epic-2-slide-generation)
- [Epic 3: Template Extension](#epic-3-template-extension)
- [Epic 4: Engineering Foundations](#epic-4-engineering-foundations)
- [Reference: Proposed JSON Schema](proposed-json-schema.md)

> **Epic map (revised):** the backlog was reorganized from 6 epics to 4, grouped by lifecycle role (one epic ≈ one skill). Old → new: **Epic 1 → 1** (extraction engine); **Epic 2 → merged into 1** (header/footer & best-practices are extraction-time schema enrichments — no separate skill/value stream); **Epic 3 → merged into 1** (the templating skill is the productization of extraction); **Epic 4 → 2** (slide generation); **Epic 6 → 3** (template extension); **Epic 5 → 4** (engineering foundations).
>
> *Story IDs (e.g., `US-2.1`, `US-6.1`) are immutable and retain their original prefixes as historical tags — they no longer match the containing epic number. See GAP-ANALYSIS for the same old→new mapping.*

---

## Epic 1: Template Extraction & Templating

The *ingest/understand* side of the lifecycle: read any source `.pptx`, parse its slide master/layouts/theme, and package that understanding back into the file as an embedded JSON schema (`ppt/template_schema.json`), producing a self-describing, reusable "templated" PPTX. This epic consolidates the extraction engine (US-1.x), schema enrichment such as header/footer and best-practice detection (US-2.x), and the templating skill UX (US-3.x) — one capability served by `generate-template-skill`. *(Formerly spread across old Epics 1, 2, and 3.)*

---

### US-1.1 [Epic 1] — Extract Slide Master to Structured JSON `[Must Have]`

**As a** OpenCode user working as a founder,
**I want** the subagent to read any PPTX file I upload, parse its slide master and all slide layouts, and convert them into a structured JSON representation,
**so that** I have a machine-readable blueprint of my template instead of relying on unpredictable LLM-only generation.

**Details:**
The system must use a scripted skill (not raw LLM) to open the PPTX zip, locate `ppt/slideMasters/` and `ppt/slideLayouts/` XML files, and parse the OOXML into the proposed JSON schema. This ensures deterministic, repeatable extraction.

**Acceptance Criteria:**
- [x] Subagent accepts a .pptx file input and does not crash on any valid PPTX.
- [x] Slide master XML is parsed; every layout under the master is extracted.
- [x] Output is a valid JSON object conforming to the proposed schema.
- [x] Extraction is performed by a Python/Node script — not by LLM guessing.

**Tags:** extraction, slide-master, ooxml, deterministic

---

### US-1.2 [Epic 1] — Normalized Polygon Positioning `[Must Have]`

**As a** OpenCode user,
**I want** each component in the extracted JSON to have a `polygon` field defined as an array of exactly 4 coordinate pairs in anti-clockwise order, with all values normalized to a 0.0–1.0 scale relative to the slide dimensions,
**so that** the position and size of every element is unambiguous, resolution-independent, and can be mapped back to any slide size (e.g., 16:9, 4:3).

**Details:**
- Coordinate order: top-left → top-right → bottom-right → bottom-left (anti-clockwise).
- Values are floats between 0.0 and 1.0, where (0,0) is the top-left corner of the slide and (1,1) is the bottom-right.
- Non-rectangular shapes (e.g., trapezoids, arrows) must use their actual polygon vertices, still normalized.
- The slide's native aspect ratio is recorded in `template_metadata.slide_dimensions` so the subagent can denormalize when writing back.

**Acceptance Criteria:**
- [x] Every component has a `polygon` array with exactly 4 `{x, y}` objects for rectangular shapes.
- [x] All x and y values are in [0.0, 1.0] range.
- [x] Anti-clockwise winding is verified by a simple cross-product check in the script.
- [x] Slide dimensions (EMU, inches, and aspect ratio string) are recorded in metadata.

**Tags:** polygon, normalized-coords, positioning

---

### US-1.3 [Epic 1] — Component Type Enumeration `[Must Have]`

**As a** OpenCode user,
**I want** each component to have a `type` field drawn from a fixed enum that maps directly to PowerPoint element types (e.g., `textbox`, `image`, `table`, `video`, `shape`, `chart`, `group`, `smartart`, `placeholder`),
**so that** the slide-generation skill knows exactly which OOXML element to create for each component, avoiding ambiguous or incorrect element types.

**Details:**
- The enum is defined in the schema and documented in the skill's README.
- OOXML `<p:sp>` elements with `<p:txBody>` → `textbox`.
- `<p:pic>` → `image`.
- `<p:graphicFrame>` containing a table → `table`.
- `<p:graphicFrame>` containing a chart → `chart`.
- `<p:sp>` with a preset geometry but no text → `shape`.
- `<p:grpSp>` → `group` (with nested `children` array).
- If the script encounters an unrecognized element, it falls back to `shape` and sets `type_confidence: "low"`.

**Acceptance Criteria:**
- [x] `type` field is always present and always one of the defined enum values.
- [x] No component has `type: null` or `type: "unknown"`.
- [x] A mapping table from OOXML tags to enum values is included in the skill source.

**Tags:** type-enum, ooxml-mapping, component

---

### US-1.4 [Epic 1] — Font Detection & Availability Checking `[Must Have]`

**As a** OpenCode user,
**I want** the extraction to capture font metadata for every text-bearing component — including `family`, `size_pt`, `weight`, `color`, `alignment`, and an `is_available` boolean — and populate a top-level `missing_fonts` array when a font is not among PowerPoint's built-in defaults,
**so that** I know exactly which fonts my template depends on and can install them before generating slides, preventing layout breakage from font substitution.

**Details:**
- Built-in PowerPoint fonts (Calibri, Arial, Times New Roman, etc.) are considered always available.
- Custom fonts get `is_available: false` and appear in `missing_fonts` with download suggestions if known.
- Each font entry includes a `fallback` field suggesting the closest built-in substitute.
- Multiple font runs within a single textbox (e.g., bold title + regular subtitle) are captured as a `runs` array.

**Acceptance Criteria:**
- [x] Every textbox component has a `font` object with all specified fields.
- [x] `missing_fonts` array is empty when all fonts are built-in.
- [x] When non-built-in fonts are found, the subagent prints a user-facing warning listing them.
- [x] `fallback` is always a built-in font name.

**Tags:** fonts, availability, fallback

---

### US-1.5 [Epic 1] — JSON Storage Inside PPTX Zip `[Must Have]`

**As a** OpenCode user,
**I want** the generated JSON template to be stored inside the PPTX zip archive at a path like `ppt/template_schema.json` — without breaking the PPTX file or causing PowerPoint to reject it,
**so that** the JSON travels with the file itself and the slide-generation skill can read it directly from any PPTX I provide, without needing a separate database or file system.

**Details:**
- PowerPoint ignores unknown files inside the zip as long as `[Content_Types].xml` is not modified.
- The script appends the JSON without altering any existing entry in the zip.
- Opening the file in PowerPoint before and after embedding produces identical behavior.
- The JSON is minified to keep file size impact minimal (typically < 50 KB).

**Acceptance Criteria:**
- [x] After embedding, the PPTX opens in PowerPoint without errors or repair prompts.
- [x] The JSON is retrievable by re-reading the zip at the known path.
- [x] Existing slide content, layouts, and media are untouched.
- [x] File size increase is logged to the user.

**Tags:** zip-embedding, pptx-safe, portability

---

### US-2.1 [Epic 1] — Header & Footer Detection `[Must Have]`

**As a** OpenCode user,
**I want** the subagent to detect whether my slide master contains header and footer zones — and, if not, explicitly ask me whether I want to add them before proceeding,
**so that** my generated slides don't lack standard structural elements that make a deck look unfinished or unprofessional.

**Details:**
The script checks for the presence of header (`<p:hdr>`) and footer (`<p:ftr>`) elements in the slide master XML. If absent, the subagent pauses extraction and returns a structured prompt to the user.

**Acceptance Criteria:**
- [x] `template_metadata.header_footer.has_header` and `.has_footer` are booleans reflecting actual detection.
- [x] When both are `false`, the subagent outputs a user-facing question before continuing.
- [x] If the user says "yes, add header", the script injects a default header zone into the JSON (not into the PPTX yet — only into the schema).

**Tags:** header, footer, detection, prompt

---

### US-3.1 [Epic 1] — End-to-End Template Generation Pipeline `[Must Have]`

**As a** OpenCode user,
**I want** to invoke a "generate template" skill that takes my PPTX file, runs the full extraction pipeline (Epic 1), and produces the JSON schema,
**so that** I have a standardized, reusable template definition without manually editing any JSON.

**Details:**
This skill is the entry point. The user says something like "extract the template from this PPTX" and the subagent routes to this skill. The skill orchestrates: zip read → XML parse → JSON build → validate → embed → return file.

**Acceptance Criteria:**
- [x] Skill is invocable by natural language intent detection (no special command needed).
- [x] Full pipeline runs end-to-end without manual intermediate steps.
- [x] Validation errors (e.g., no slide master found) are reported clearly to the user.

**Tags:** skill, template-generator, pipeline

---

### US-3.2 [Epic 1] — Template Naming `[Must Have]`

**As a** OpenCode user,
**I want** the generated JSON to include a `title` field at the top level that names the template (e.g., "Q3 Investor Pitch Deck — Dark Theme"),
**so that** when I have multiple templated PPTX files, I can query or reference them by name and the subagent knows which template I mean.

**Details:**
The title is inferred from the PPTX file metadata (document title property), or from the first slide's title text, or — if neither exists — the subagent prompts the user to name it.

**Acceptance Criteria:**
- [x] `template_metadata.title` is always a non-empty string.
- [x] Inference order: core.xml title → slide 1 title text → user prompt.
- [x] The title is displayed to the user for confirmation after extraction.

**Tags:** metadata, naming, queryable

---

### US-3.3 [Epic 1] — Return Downloadable Templated PPTX `[Must Have]`

**As a** OpenCode user,
**I want** the skill to return a new PPTX file (with the JSON embedded) that I can download and keep,
**so that** I have a portable, self-describing template file I can reuse across sessions or share with teammates.

**Details:**
The returned file is the original PPTX plus `ppt/template_schema.json`. The subagent provides a download link and a summary of what was extracted (number of layouts, components, fonts, etc.).

**Acceptance Criteria:**
- [x] Downloadable PPTX is provided via OpenCode's file output mechanism.
- [x] A human-readable extraction summary is printed to the user.
- [x] File passes a round-trip test: open in PowerPoint, re-upload, re-extract → identical JSON.

**Tags:** output, download, round-trip

---

### US-3.4 [Epic 1] — Theme & Color Extraction `[Should Have]`

**As a** OpenCode user,
**I want** the template JSON to also capture the color theme (primary, secondary, accent, background) and font palette from the PPTX's `ppt/theme/theme1.xml`,
**so that** when the slide-generation skill creates new content, it can use the exact brand colors and typography without guessing.

**Details:**
Theme colors are extracted from `<a:clrScheme>` and mapped to semantic roles. Font palette comes from `<a:fontScheme>`.

**Acceptance Criteria:**
- [x] `theme` object contains `primary_color`, `secondary_color`, `accent_color`, `background_color` as hex strings.
- [x] `theme.font_palette` contains `heading`, `body`, `accent` font names.
- [x] If theme1.xml is missing or malformed, sensible defaults are used and a warning is shown.

**Tags:** theme, colors, branding

---

## Epic 2: Slide Generation

The *generate* side: consume a templated PPTX and produce new, on-brand slides — reading the embedded JSON for layout selection, fitting text, batching, multi-aspect-ratio output, and auto-chaining extraction when no JSON is present. Served by `generate-slide-skill`. *(Was Epic 4.)*

---

### US-4.1 [Epic 2] — Read Embedded JSON as Layout Reference `[Must Have]`

**As a** OpenCode user,
**I want** the slide-generation skill to read the `ppt/template_schema.json` from a templated PPTX I provide and use it as the authoritative layout reference,
**so that** every generated slide places content in the correct positions with the correct styling, rather than relying on LLM hallucination of coordinates.

**Details:**
The skill reads the JSON **from the zip** (it does not re-extract or re-parse the PPTX XML), identifies which slide layout to use based on the user's intent (e.g., "title slide", "content slide", "two-column") via `layout_name` matching, and generates new slides **using the slide master's own layouts** (`add_slide(layout)`) — NOT by manually placing OOXML elements at polygon coordinates. The embedded JSON is the faithful, portable description of the template (layout names, component types, fonts, theme, normalized positions) that drives layout selection and consistency; the template's layouts themselves carry the actual positioning and inherited styling (bullets, theme, master defaults). The normalized `polygon` coordinates (US-1.2) remain a faithful geometric description and may feed an optional consistency/conformance check — they are **not** a placement data source.

**Acceptance Criteria:**
- [x] Skill reads JSON from the zip — does not re-extract or re-parse XML.
- [x] Layout selection is based on `layout_name` matching or user confirmation.
- [x] Generated slides use the template's own layouts (via `add_slide`); the embedded JSON drives layout selection, not element placement at polygon coordinates. (A polygon-fidelity consistency check is optional and non-fatal.)

**Tags:** slide-generation, layout-matching, embedded-json

---

### US-4.2 [Epic 2] — Visually Pleasing Output with Text Fitting `[Must Have]`

**As a** OpenCode user,
**I want** generated slides to be visually pleasing — with proper text fitting (no overflow), appropriate font sizing for the content length, and consistent spacing — while still adhering to the template's defined zones,
**so that** the output looks like a human-designed slide, not a raw data dump into text boxes.

**Details:**
- The script implements text-fitting logic: if text exceeds the textbox's polygon area at the template's font size, it reduces size in steps (e.g., -2pt) until it fits, with a minimum floor.
- Line spacing and paragraph spacing are derived from the template's existing text runs.
- Bullet points use the template's bullet style if detected; otherwise, a clean default is applied.

**Acceptance Criteria:**
- [ ] No generated slide has text overflowing its bounding box.
- [x] Font size is only reduced when necessary — never increased beyond the template's defined size.
- [x] A `font_size_adjusted` flag is set in the component when auto-sizing occurs.

**Tags:** text-fitting, visual-quality, auto-sizing

---

### US-4.3 [Epic 2] — Auto-Chain Extraction When No JSON Present `[Must Have]`

**As a** OpenCode user,
**I want** to be able to hand the subagent a PPTX that does NOT yet have an embedded JSON and say "extract the template from this, then generate slides using it" — all in one interaction,
**so that** I don't have to run two separate commands when I'm working with a fresh file for the first time.

**Details:**
The subagent detects the absence of `ppt/template_schema.json`, automatically chains the Template Generator skill first, then proceeds to Slide Generation. The user is informed of the two-step process but doesn't need to manually trigger each step.

**Acceptance Criteria:**
- [x] Single user prompt triggers both skills in sequence without error.
- [x] Intermediate JSON is embedded in the output PPTX.
- [x] User sees a status message like "No template found — extracting first, then generating slides...".

**Tags:** chaining, auto-detect, ux

---

### US-4.5 [Epic 2] — Multi-Slide Batch Generation `[Could Have]`

**As a** OpenCode user,
**I want** the slide generator to support multi-slide batch generation from a single prompt — e.g., "create a 10-slide investor deck with market, problem, solution, team, and financials sections",
**so that** I can generate a full presentation in one go rather than creating slides one at a time.

**Details:**
The LLM plans the slide order and content outline first (as a structured array), then the script iterates over the plan, selecting the appropriate layout for each slide and filling in the content. A progress indicator shows completion (e.g., "Slide 3/10").

**Acceptance Criteria:**
- [x] Single prompt can produce 2–20+ slides in one PPTX.
- [x] Each slide uses the correct layout for its content type.
- [x] Progress is reported to the user during generation.
- [x] The LLM outline is shown to the user before generation begins, with an option to edit.

**Tags:** batch, multi-slide, outline

---

### US-4.6 [Epic 2] — Multi-Aspect-Ratio Rendering `[Should Have]`

**As a** OpenCode user,
**I want** to generate a deck at a different slide size or aspect ratio than the template (e.g., render a 4:3 deck from a 16:9 template), with every element — textboxes, images, shapes — scaling proportionally to the new dimensions,
**so that** I can reuse one template across multiple output formats (16:9, 4:3, square) without redesigning the template or getting a misaligned layout.

**Details:**
When the target slide dimensions differ from the template's native size, the slide-generation skill deviates from the default US-4.1 path (`add_slide(layout)`, which renders at the template's native size) and switches to a **coordinate-placement path**: it reads the embedded JSON's normalized `polygon` coordinates (US-1.2, 0.0–1.0), denormalizes them against the **target** slide dimensions, and creates OOXML elements at the resulting EMU positions — yielding proportional scaling of every element to the new size. This is possible because US-1.2's normalized coordinate model is resolution-independent by design. The skill prompts the user for (or infers) the target aspect ratio. Because layout placeholders are not used on this path, styling that python-pptx would otherwise inherit (fonts, theme colors, bullets) is re-applied from the embedded JSON's `theme` and per-component `font` metadata.

**Acceptance Criteria:**
- [x] Given a 16:9 templated PPTX, the skill renders an equivalent deck at 4:3 on request (and vice versa), via the coordinate-placement path.
- [x] Every element (textboxes, images, shapes) scales proportionally to the new dimensions — no clipping or misalignment beyond the US-4.2 text-fitting tolerance.
- [x] Normalized `polygon` coordinates are denormalized against the **target** slide size; resulting positions are within 1% of the proportionally-scaled originals.
- [x] Fonts/theme/bullets are re-applied from the embedded JSON metadata so the output stays on-brand despite bypassed layout inheritance.
- [x] When the target size equals the template's native size, the default US-4.1 `add_slide(layout)` path is used (this story is a no-op in that case).

**Tags:** multi-aspect-ratio, coordinate-placement, proportional-scaling, resolution-independent

---

### US-4.7 [Epic 2] — Template Selection & Pre-Render Validation `[Must Have]`

**As a** OpenCode user,
**I want** the engine to use a default template when I don't name one, accept any `.pptx` path I give in conversation, and refuse to render (with a clear error) when the chosen template is structurally broken,
**so that** generation never silently produces a broken deck from an unusable template.

**Details:**
- The default template is `template/default.pptx` (repo root), used whenever no `template_path` is supplied; keeping it at the top level (mirroring `output/`) makes it easy to find and edit.
- A user-supplied template is passed **as a path** (`template_path` / CLI `--template`), not by overwriting the default — replacing the earlier copy-overwrite workflow.
- A **pre-flight** runs on every load (default or user-supplied): severe problems raise `TemplateError` and abort before the render loop; minor issues stay non-fatal warnings.

**Acceptance Criteria:**
- [x] With no template specified the engine renders against `template/default.pptx`; a user-supplied `.pptx` path is passed through `template_path` (the default is never overwritten).
- [x] A severe template problem — corrupt/non-PPTX, no slide master, zero layouts, or serving none of the 8 slide types — raises a clear `TemplateError` instead of producing a broken deck.
- [x] Minor issues (missing fonts, no header/footer, small content area, no embedded schema) remain non-fatal warnings; generation proceeds.

**Tags:** template, validation, error-handling

---

## Epic 3: Template Extension

The *adapt* side: at render time, when a planned slide's type has no matching layout, clone a donor layout into a derived `template_new.pptx` so the deck still renders — never mutating the user's original. Served by `template-modifier-skill`. *(Was Epic 6.)*

---

### US-6.1 [Epic 3] — Extend Template When a Layout Is Missing `[Should Have]`

**As a** OpenCode user with a minimal or specialized custom template,
**I want** the subagent to detect when my template lacks a layout needed for a planned slide, and automatically extend the template with a cloned layout into a derived file (never modifying my original),
**so that** my deck still renders with an appropriate layout instead of silently skipping the slide or crashing.

**Details:**
- Runs as a pre-render step: for each planned slide, check whether the template provides a layout whose placeholder-composition fingerprint matches the slide's type.
- When a slide type has no matching layout, clone a donor layout (closest fingerprint) into a derived `template_new.pptx` via XML/part cloning (python-pptx exposes no public API to add layouts), and pin the slide type to the cloned layout via config overrides.
- The base `template.pptx` is **immutable** — never written. Clones save only to the derived file. After cloning, reload-verify (the new layout must be findable by name); on any failure, roll back (delete the derived file) and fall back to the base so the deck still renders.
- By default, over-limit content (a body too large for its placeholder) is handled by density-mode downshift, **not** by cloning. Cloning for over-limit content is an opt-in policy.
- Whenever the derived template is used, a mandatory notification states which template was used and why.

**Acceptance Criteria:**
- [x] Before render, the engine detects any slide type whose layout is missing from the template (no fingerprint match) and flags it for extension.
- [x] A missing layout is cloned into a derived `template_new.pptx`; the original `template.pptx` is never modified.
- [x] The cloned layout is pinned to its slide type via config overrides, and the fill engine renders the entire deck against the active (base or derived) template in one pass.
- [x] Clone failure is non-fatal: the derived file is discarded and the base template is used; the deck still renders.
- [x] Whenever the derived template is used, the user is notified which template was used and the reason.

**Tags:** template-extension, layout-cloning, capability-b, graceful-degradation

---

## Epic 4: Engineering Foundations

Cross-cutting, non-functional foundations shared by all three skills: CLI architecture & exit codes, the shared JSON-schema validation contract, and structured logging. *(Was Epic 5.)*

---

### US-5.1 [Epic 4] — Two Independent Skills with CLI Scripts `[Must Have]`

**As a** OpenCode user (and as the developer maintaining the subagent),
**I want** the system to expose exactly 2 skills — "generate-template" and "generate-slides" — each backed by a dedicated, independently testable script,
**so that** each skill has a single responsibility, can be debugged in isolation, and follows OpenCode's documented skill pattern.

**Details:**
- **generate-template**: Script accepts `--input path/to/file.pptx --output path/to/output.pptx`. Reads zip, extracts, builds JSON, embeds, writes new zip.
- **generate-slides**: Script accepts `--template path/to/templated.pptx --prompt "..." --output path/to/deck.pptx`. Reads JSON, generates slides, writes new zip.
- Both scripts exit with meaningful exit codes (0 = success, 1 = validation error, 2 = runtime error).

**Acceptance Criteria:**
- [ ] Each skill has its own directory with `skill.yaml`, script, and README.
- [x] Both scripts are runnable from the CLI independently of the LLM.
- [x] Exit codes are documented and used consistently.

**Tags:** architecture, cli, testability, single-responsibility

---

### US-5.2 [Epic 4] — Shared JSON Schema for Validation `[Must Have]`

**As a** developer maintaining the subagent,
**I want** a shared JSON Schema (JSON Schema draft-07 or 2020-12) that both scripts use to validate the template JSON — with the schema file shipped alongside the skills,
**so that** the template generator's output is guaranteed to be consumable by the slide generator, catching schema drift at build time rather than at runtime.

**Details:**
The schema file (`template_schema.json`) lives in a shared `common/` directory. Both scripts run `validate(json, schema)` before reading or writing. The LLM is also given the schema in its system prompt so it understands the structure when reasoning about slide content.

**Acceptance Criteria:**
- [ ] A `.json` schema file exists and is referenced by both scripts.
- [x] Template generator validates its output before embedding.
- [ ] Slide generator validates the JSON before reading it.
- [x] Schema version is tracked in `template_metadata.schema_version`.

**Tags:** json-schema, validation, contract, versioning

---

### US-5.3 [Epic 4] — Structured Logging `[Should Have]`

**As a** OpenCode user,
**I want** the subagent to log structured output (JSON lines) for every operation — including extraction steps, layout selections, font warnings, and generation progress — to a log stream I can inspect,
**so that** when something goes wrong (e.g., a slide looks off), I can trace exactly what the script did and share the log for debugging.

**Details:**
Each log line is a JSON object with `timestamp`, `level`, `skill`, `action`, and `details`. Logs are written to stderr so they don't interfere with file output on stdout.

**Acceptance Criteria:**
- [ ] Every significant action emits a structured log line.
- [ ] Logs go to stderr; only file paths go to stdout.
- [ ] Log level can be controlled via `--log-level` flag (debug, info, warn, error).

**Tags:** logging, debugging, observability
