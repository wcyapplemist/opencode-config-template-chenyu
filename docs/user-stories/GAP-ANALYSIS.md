# Gap Analysis — chenyyu-user-stories vs Current Implementation

> **Document type:** Requirements-vs-implementation gap analysis (analysis only — no code changes)
> **Requirements source:** `docs/user-stories/user-stories.md` (4 Epics, 21 Stories)
>
> *Note: counts reflect the current source document — 21 stories across 4 Epics (after the US-4.4 descope, the US-4.7 addition, and the Epic reorganization from 6→4). Historical per-revision counts in the log below are preserved as-is.*
> **Implementation audited:** `.opencode/skills/generate-slide-skill/`, `.opencode/skills/template-modifier-skill/`, `.opencode/agents/pptx-subagent.md`
> **Date:** July 2026 (latest: Revision 20 — US-2.2 descoped & removed)

## Revision Log (compact)

| Rev | Trigger | Key change | Counts (Met / Partial / Not met / Differs) |
|---|---|---|---|
| 2 | US-1.1 (PR #49) | `schema_extractor.py` emits proposed-schema JSON | 2 / 7 / 9 / 0 |
| 3 | US-1.2 (PR #51) | cross-product winding check (AC3) | 3 / 6 / 9 / 0 |
| 4 | US-1.3 (#52) | full 10-value enum; `type_confidence` always emitted; `audio` reachable | 4 / 5 / 9 / 0 |
| 5 | audit-path consistency | re-grade Epic 3/5 on the extractor path | 5 / 6 / 7 / 0 |
| 6 | US-1.4 (#54) | font detection + missing-font allowlist | 6 / 6 / 6 / 0 |
| 7 | US-1.5 (#55) | zip embed at `ppt/template_schema.json` — **Epic 1 done** | 7 / 6 / 5 / 0 |
| 8 | US-3.1 (#56) | `generate-template-skill` pipeline; US-3.2/3.3 ✅ — **Epic 3 done** | 10 / 5 / 3 / 1 |
| 9 | US-4.1 (#58) | renderer reads embedded JSON via adapter (source-swap) | 11 / 4 / 3 / 1 |
| 10 | US-4.2 (#60/#62) | `text_fit.py` (AC2/AC3 ✅; AC1 deferred) | 11 / 4 / 3 / 1 |
| 11 | US-4.3 (#63) | auto-templated output; ⚪→✅ | 12 / 4 / 3 / 0 |
| 12 | skill rename | `ppt-template-filler` → `generate-slide-skill` | 12 / 4 / 3 / 0 |
| 13 | US-2.1 (#68) | header/footer detection — **no unmet Must-Have** | 13 / 4 / 2 / 0 |
| 14 | US-4.6 backfill (PLAN-GIT-70) | multi-aspect-ratio via coordinate-path prep | 14 / 4 / 2 / 0 |
| 15 | US-6.1 added | `template-modifier-skill` (Capability B) — 21 stories, 6 Epics | 15 / 4 / 2 / 0 |
| 16 | Epic 5 `_common/` (PLAN-GIT-72) | shared infra extracted; both engine CLIs exist | 15 / 4 / 2 / 0 |
| 17 | scope change | US-4.4 descoped | 15 / 4 / 1 / 0 |
| 18 | US-4.7 added | template selection + `TemplateError` pre-flight — 21 stories | 16 / 4 / 1 / 0 |
| 19 | US-4.2 AC1 attempt (#74) | normAutofit **falsified & reverted** — US-4.2 stays Partial | 16 / 4 / 1 / 0 |
| 20 | US-2.2 descoped | common-practice suggestions **removed** from backlog + `common_practices` schema field dropped — 20 stories | 16 / 4 / 0 / 0 |

*Full per-revision narrative preserved in Appendix A.*

---

## Executive Summary

> **Epic renumbering (current):** this analysis uses the revised 4-Epic structure (was 6). Old → new: Epic 1 → **1** (Template Extraction & Templating); Epic 2 → **merged into 1**; Epic 3 → **merged into 1**; Epic 4 → **2** (Slide Generation); Epic 6 → **3** (Template Extension); Epic 5 → **4** (Engineering Foundations). Story IDs (US-x.y) are unchanged. The revision log (Rev 1–17) retains the original Epic numbers as a historical record.

This report compares `user-stories.md` (the requirements document) against the current implementation. **Core conclusion: the architectural route has diverged.**

- The requirements document describes a **"normalized polygon JSON Schema embedded in PPTX"** architecture: each component contains a 4-point 0–1 normalized `polygon`, a `type` enum, `font`/`runs`, and JSON embedded inside the zip at `ppt/template_schema.json`.
- The current implementation follows a **"introspection contract with placeholder fingerprint matching"** architecture: placeholders use absolute inch coordinates, the contract is stored as a sidecar file `template.pptx.contract.json`, and layouts are matched by fingerprint.

Both achieve "fill any template", but they differ on **data model, skill decomposition, and artifact form**.

**Story-by-story summary** (20 stories, 4 Epics): Met 16 / Partial 4 / Not met 0.

The delivered-stories detail and the remaining-gaps breakdown live in **§3** (statistics + priority matrix) and **§4** (Open Work). **No fully-unmet Must-Have remains** — the open items are four Partials. *(US-2.2 was descoped — Rev 20.)*

For the architecture comparison see §1; for the per-story detail see §2.

---

## §1 Architecture Comparison

The requirements document and the implementation describe two different routes to the same goal ("fill any template"). They diverge on **data model, skill decomposition, artifact form, and layout-matching strategy**.

| Dimension | Requirements (chenyu-user-stories) | Current Implementation |
|---|---|---|
| **Skill decomposition** | Exactly 2 skills: `generate-template` (extraction) + `generate-slides` (generation), each a standalone CLI script | 2 skills + 1 agent: `generate-slide-skill` (filling) + `template-modifier-skill` (extension) + `pptx-subagent` (content strategy). Both engine scripts are CLI-runnable (`ppt_builder.py main()` + `schema_extractor.py`, Rev 16); manifests are `SKILL.md`, not `skill.yaml`. |
| **Template data model** | Normalized JSON Schema: each component has `polygon` (4 anti-clockwise 0–1 coords), `type` enum, `font`/`runs`, `content_template` with `{{placeholders}}` | Introspection contract: each placeholder has `idx/name/type/left_in/top_in/width_in/height_in` + composition `fingerprint` |
| **JSON storage** | **Embedded inside the PPTX zip** at `ppt/template_schema.json` (PowerPoint-safe) | **Sidecar file** `template.pptx.contract.json` next to the template (mtime-cached, gitignored) |
| **Layout matching** | Read embedded JSON → denormalize polygon coords back to EMU → place OOXML at exact positions | Placeholder-composition **fingerprint match** + name fallback via python-pptx `add_slide(layout)` (`ppt_builder.py:298`) |
| **Template generation output** | A downloadable PPTX with the embedded JSON (round-trip tested) | A sidecar contract; the original `.pptx` is never modified |

> **Revision 2 note (§1):** A new parallel module `schema_extractor.py` now **coexists** with `template_introspector.py` — it emits the **proposed-schema** JSON (normalized `polygon`, `type` enum, `components[]`) conforming to `schemas/template_schema.json`, separate from the renderer's fingerprint contract. The two data models are now "partially bridged"; the renderer still consumes only the contract. GAP §5 Decision 1 (Coexist) is now reality.

---

## §2 Story-by-Story Gap Analysis

Status legend: ✅ Met · 🟡 Partial · ❌ Not met · ⚪ Architecture differs (not directly applicable)

### Epic 1 — Template Extraction & Templating

#### US-1.1 — Extract Slide Master to Structured JSON `[Must Have]` — ✅ Met

**Implemented (PR #49).** `schema_extractor.extract_schema()` (`schema_extractor.py`) reads any `.pptx`, parses the slide master (`prs.slide_masters[0]`) AND every layout, and emits a structured JSON conforming to `schemas/template_schema.json`. All four ACs are met: no crash on valid PPTX (`TemplateExtractionError` on bad input); master parsed + every layout enumerated; output is validated by a hand-rolled `validate_template_schema()` that mirrors `template_schema.json`'s rules (the file is a conformance-target spec, **not** loaded at runtime — see US-5.2); deterministic Python. The renderer's fingerprint contract (`template_introspector.py`) is untouched — the two modules coexist (§5 Decision 1, now reality). 49 tests pass (`test_schema_extractor.py`).

#### US-1.2 — Normalized Polygon Positioning `[Must Have]` — ✅ Met

**Met (PR #51).** All four ACs satisfied. `normalize_polygon()` emits exactly 4 normalized `{x,y}` points in `[0,1]` (AC1/AC2); slide dimensions in metadata (AC4). AC3 — the cross-product winding check — is delivered by `_signed_area()` + a check in `validate_template_schema()`: the canonical order TL→TR→BR→BL yields a **positive signed area**, which is algebraically counter-clockwise (CCW = anti-clockwise), exactly what AC3 asks a cross-product to verify. (Reversed winding → error; degenerate/zero-area → warning.) Note: in screen coords (Y-down) the trace visually appears clockwise, but the algebraic winding is CCW — documented in `template_schema.json` `$comment`. **Out of scope (Details, not ACs):** non-rectangular actual vertices (custGeom/triangle/connector) — polygon stays a 4-point rectangular bounding box; deferred (polygon is metadata-only, no consumer).

#### US-1.3 — Component Type Enumeration `[Must Have]` — ✅ Met

**Met (issue #52).** `schema_extractor._classify_shape()` (`schema_extractor.py`) applies the **full 10-value enum** to all elements (placeholders and freeform shapes). All three ACs are met (type always present; unknowns degrade to `shape`, never null/unknown; OOXML→enum mapping in source). The two previously-deferred **Details** are now delivered: (1) `type_confidence` is **always emitted** (`"high"` default; `"low"` only when `shape_type` is `None`/unreadable or MEDIA is indeterminate — no whitelist, per architecture-review MAJOR-1, so recognized-but-unmapped members like `LINKED_PICTURE`/`TEXT_EFFECT`/`CALLOUT` stay `"high"`); (2) the `"audio"` enum value is **reachable** via OOXML `<a:audioFile>`/`<a:videoFile>` split of `MSO_SHAPE_TYPE.MEDIA`, and `WEB_VIDEO`→`video/high`. A non-fatal `ValidationIssue` WARNING surfaces `shape/low` ("flagged for review"). Optional `type_confidence` added to `schemas/template_schema.json`. 64 tests in `test_schema_extractor.py` (was 49).

#### US-1.4 — Font Detection & Availability Checking `[Must Have]` — ✅ Met

**Met (issue #54).** `schema_extractor._extract_text_fonts()` (`schema_extractor.py`) populates every text-bearing component's `font` (`family`, `size_pt`, `weight`, `color`, `alignment`, `is_available`, `fallback`) — **explicit-only** (inherited values → `null`); Latin/English only (CJK `<a:ea>`/`<a:cs>` out of scope). It captures a nested `runs[]` (`{text, font:{...}}`) and a guarded RGB `color` (`color.type == MSO_COLOR_TYPE.RGB`, else `null`). The deck aggregates a deduped top-level `missing_fonts[]` (`{family, is_available:false, fallback, download_url:null}`) against a curated `_BUILTIN_FONTS` allowlist; `fallback` defaults to the theme body font (if built-in) else `Arial`, always a built-in name (AC4 → validator ERROR). `validate_template_schema` emits a non-fatal `ValidationIssue(severity="warning")` per missing font (AC3); `extract_schema` also `logger.warning`s. All 4 ACs met. 78 tests in `test_schema_extractor.py` (was 65).

#### US-1.5 — JSON Storage Inside PPTX Zip `[Must Have]` — ✅ Met

**Met (issue #55).** `schema_extractor.embed_schema()` (`schema_extractor.py`) writes the schema into the PPTX zip at `ppt/template_schema.json` via an **order-preserving full rewrite**: `[Content_Types].xml` first (with a `<Default Extension="json" ContentType="application/json"/>` injected — strict-safe, architecture review MAJOR-2), every other original entry **decompressed-content-identical** in original order (MAJOR-1; AC3), then the minified schema appended. **Idempotent** (re-embed replaces, never duplicates — MAJOR-3), **atomic** (temp + `os.replace` — MINOR-6), returns an `EmbeddedSchemaResult` (AC4 size delta — MINOR-5). `read_embedded_schema()` retrieves it with a clear error contract (absent→`None`; malformed→`None`+warn; non-zip→`TemplateExtractionError` — MINOR-3/4). CLI `--embed` + `--output-pptx` (additive; without-`--embed`→exit 2). AC1 verified by proxy (`python-pptx` re-opens; `[Content_Types].xml` first + declares `json` + originals intact; all other entries hash-identical — real PowerPoint is manual). 93 tests in `test_schema_extractor.py` (was 80). → **Extraction engine complete (US-1.1–1.5)**.

#### US-2.1 — Header & Footer Detection `[Must Have]` — ✅ Met (Rev 13)

`_detect_header_footer(prs)` scans the slide master's placeholders for HEADER/FOOTER types and records `{has_header, has_footer}` booleans in `template_metadata.header_footer` (AC1). `needs_header_footer_prompt(schema)` returns True when both are absent (AC2). `inject_default_header_zone(schema)` injects a 4-point top-strip polygon + English note into the schema (AC3, schema-only). `generate-template-skill` Stage 2 prompts the user (batched with title confirmation per arch-review M2) and injects on "yes"; `pptx-subagent` Stage 0 surfaces a light note via `read_embedded_schema` for templated inputs (arch-review M1: `get_render_contract`/adapter strips `template_metadata`, so `read_embedded_schema` is the only accessor; non-templated inputs defer the note). 9 tests; full suite 405 passed.

#### US-2.2 — Common Practice Suggestions `[Should Have]` — *descoped (Rev 20)*

**Descoped and removed from the backlog (Rev 20).** The story (suggest common PowerPoint practices — slide numbers, logo, margins, section dividers, closing slide) will not be implemented; the placeholder `common_practices` schema field was removed. The requirement point ("other common practices suitable for powerpoint slide") is intentionally not addressed. See Rev 20.

#### US-3.1 — End-to-End Template Generation Pipeline `[Must Have]` — ✅ Met (Rev 8)

A standalone `generate-template-skill` (`.opencode/skills/generate-template-skill/SKILL.md`) now orchestrates the full pipeline end-to-end via the `schema_extractor` engine: `extract → validate → (title confirm) → embed → return templated PPTX + summary`. NL intent routing is via the SKILL.md `description` (extraction verbs) + a one-line "What NOT to Handle" deferral in `pptx-subagent.md` (architecture review MAJOR-1 — the agent's greedy `pptx` triggers would otherwise have misrouted extraction requests). All three ACs met.

#### US-3.2 — Template Naming `[Must Have]` — ✅ Met (Rev 8)

`_infer_title` now returns a `TitleInference(title, source)` NamedTuple and `_build_metadata` emits `template_metadata.title_source` (`core_xml` | `slide1` | `filename`); the skill prompts the user to name the template when `title_source == "filename"` and writes back `title_source = "user"` on an override (AC2 — the inference order now ends in a **user prompt**, not just a filename). The skill always displays the title for confirmation (AC3). AC1 (non-empty title via the filename fallback) was already met. All three ACs met.

#### US-3.3 — Return Downloadable Templated PPTX `[Must Have]` — ✅ Met (Rev 8)

**Corrects the Rev-7 stale rating.** US-1.5 already delivered `embed_schema` (produces the downloadable templated PPTX) and the round-trip test (`test_round_trip_deep_equal`). This issue adds the skill surface (the downloadable PPTX is returned via `output/<stem>.templated.pptx`, AC1) and `build_extraction_summary(schema) -> str` + CLI `--summary` (AC2 — a human-readable summary of layouts/components/fonts/theme). AC3 (round-trip) was already met. All three ACs met.

#### US-3.4 — Theme & Color Extraction `[Should Have]` — ✅ Met

The renderer's `_build_theme()` (`template_introspector.py`) extracts only raw OOXML role colors. **But the proposed-schema path fully implements it**: `schema_extractor._build_theme()` maps raw colors to semantic roles (`primary_color`/`secondary_color`/`accent_color`/`background_color`/`text_color`) and builds `font_palette.{heading,body,accent}`; `_raw_theme_colors_and_fonts()` is wrapped in try/except that logs a warning and yields empty defaults on a missing/malformed theme. All three ACs (semantic colors as hex; `font_palette`; sensible defaults + warning) are met on the `schema_extractor` path.

### Epic 2 — Slide Generation

#### US-4.1 — Read Embedded JSON as Layout Reference `[Must Have]` — ✅ Met (Rev 9)

The renderer now reads the **embedded** `ppt/template_schema.json` (not the sidecar) via `ppt_builder.get_render_contract` → `contract_adapter.embedded_schema_to_contract` (US-4.1, issue #58). Layout selection still uses `layout_name`/fingerprint matching (`_resolve_layout_by_fingerprint`); generation keeps `add_slide(layout)` (chenyu's #4 — "using the slide master's slide template"). AC3 ("within 1%") was clarified: coordinate placement was never required (see the US-4.1 historical note + §5 Decision 2); it is deferred to US-4.6 (multi-aspect-ratio). Architecture-review findings C1/M3/M4/M5/M6/m1-m3 all addressed. All three ACs met.

#### US-4.2 — Visually Pleasing Output with Text Fitting `[Must Have]` — 🟡 Partial (AC2/AC3 Met; AC1 deferred)

Delivered (issue #60 / #62). Text fitting is **reactive**: a pure `text_fit.py` estimator shrinks a placeholder's font in −2pt steps (8pt floor) when text would overflow its box, writing an explicit `run.font.size` **only on actual shrink** (else inheritance is preserved) — AC3 ✅ (`font_size_adjusted` in the `<output>.render.json` sidecar). The base size is template-derived (M1 chain: schema `size_pt` → layout sample-run → conservative role ceiling body **14** / title **28** / subtitle **18**), replacing the `Pt(14)`/`Pt(12)` hardcode — AC2 ✅ ("≤ resolved base"). An auto-grow guard prevents false-shrinking on short-base-height placeholders; inter-paragraph spacing is reserved in the height estimate. **AC1 is deferred by decision (architecture-review C1):** python-pptx has no layout/text-measurement engine, so the hard "no overflow" guarantee **cannot be verified** by the engine — it is delivered best-effort (conservative estimator + `word_wrap=True`, a horizontal-only backstop that does not prevent vertical overflow), and a full oracle (LibreOffice headless render-to-image) is out of scope. AC1 stays unchecked. *(Attempted closure via PowerPoint-native `normAutofit` in #74 was **empirically falsified** (Rev 19 reverted): `<a:normAutofit/>` without a pre-computed `fontScale` is inert on file open — PowerPoint does not recompute the shrink on open, only on user edit; the #74 deck visibly overflowed in PowerPoint including the control slides. AC1 remains deferred. The falsification also surfaced that `text_fit.py` systematically underestimates rendered height — likely text-frame internal margins / line-spacing — tracked as a follow-up research issue.)* *(The bundled-template mis-fingerprint defect #61 — body/picture placeholders showing only TITLE/SUBTITLE — was resolved by shipping a pre-templated replacement with real OBJECT/BODY/PICTURE placeholders; all 8 slide types servable.)*

#### US-4.3 — Auto-Chain Extraction When No JSON Present `[Must Have]` — ✅ Met (mechanism differs; function + all 3 ACs met, Rev 11)

Because the model is contract-based (introspection runs automatically before every render) rather than embedded-JSON-based, there is no literal "no JSON → extract first" two-skill chain. Instead, the **engine inlines** the behavior (Rev 11): `generate_ppt_from_data(auto_template=True)` re-embeds `ppt/template_schema.json` into the **output** after save (python-pptx otherwise strips the part), sourcing the schema from the **input template** (arch-review M1 — the title is the template's identity, not the rendered deck's cover) and skipping a stale embedded input schema (M2); the agent detects a non-templated input at Stage 0 (`read_embedded_schema`, exception-safe) and emits *"No template found — extracting first, then generating slides..."*. Every output `.pptx` is a self-describing/templated deck (AC2); the one-line status message fires for non-templated inputs (AC3); and any PPTX renders in one call via the `get_render_contract` sidecar/embedded fallback (AC1). The interactive `generate-template-skill` is **not** chained (headless-infeasible + agent `task` permission denies it) — "architecture differs, function met". The render report gains an additive `templating` field.

#### US-4.5 — Multi-Slide Batch Generation `[Could Have]` — ✅ Met

The multi-stage pipeline (outline → critique → detail → render) generates 2–20+ slides from one prompt, the outline is shown for user approval, and the stages act as a progress indicator. Fully satisfied.

#### US-4.6 — Multi-Aspect-Ratio Rendering `[Should Have]` — ✅ Met (Rev 14)

Delivered via the coordinate path. `generate_ppt_from_data(..., target_size=...)` accepts a preset (`16:9`/`4:3`/`1:1`) or explicit `{width_in,height_in}`. A **ratio** no-op gate (`geometry.aspect_ratios_match`) keeps the native US-4.1 path when the target ratio matches the template's native ratio (AC5); otherwise a coordinate-path **prep** (`_apply_target_resize` + `_scale_shapes_geometry`) resizes the canvas and proportionally rescales every master/layout shape, then the shared native `add_slide` loop fills target-geometry placeholders — so fonts/theme/bullets stay on-brand via normal layout inheritance (AC4, no manual re-application needed). AC3 ("within 1%") is mechanical (pure ratio scaling / polygon round-trip). Pure modules: `geometry.py` (polygon/EMU primitives, de-duped from 3 modules) + `coordinate_placer.py` (placement planner, reserved for a future freeform rebuild). Schema gains `text_properties.bullets` + `image_properties` (partname/rId); `schema_version` → 1.1.0. The output's embedded `slide_dimensions` is rewritten to the target size (self-describing); `<output>.render.json` gains an `aspect_ratio` field. CLI `--target-size`. All 5 ACs Met (PLAN-GIT-70, PR #71).

#### US-4.7 — Template Selection & Pre-Render Validation `[Must Have]` — ✅ Met (Rev 18)

New story. The default template is now **`template/default.pptx`** at the repo root (used when `template_path` is omitted), moved out of the deep `scripts/templates/` path for easy discovery/editing (the `template.config.json` pin file became `default.config.json` in lockstep). A user-supplied `.pptx` path is passed through `template_path`/`--template` instead of copy-overwriting the default (the earlier `cp`-overwrite workflow is retired in both `pptx-subagent.md` and `generate-slide-skill/SKILL.md`). A new `TemplateError` + `_validate_template()` pre-flight runs on every load (default or user-supplied): corrupt/non-PPTX (the `Presentation(...)` open is wrapped), no slide master, zero layouts, or serving none of the 8 slide types → clear error + abort before the render loop; minor issues (missing fonts, no header/footer, small content area, no embedded schema) stay non-fatal warnings (unchanged). The CLI maps `TemplateError` to exit 1 (input error, not a runtime crash). All 3 ACs Met (+12 tests in `test_template_validation.py`).

### Epic 3 — Template Extension

#### US-6.1 — Extend Template When a Layout Is Missing `[Should Have]` — ✅ Met (Rev 15)

`template-modifier-skill` delivers Capability B. `constraint_checker.evaluate_slide` flags a slide type whose layout is absent in the template (no fingerprint-matching layout) — the `cause="missing"` verdict (AC1). `state_machine.resolve_and_clone` runs the full lifecycle (① delete leftover → ② `get_render_contract` → ③ scan → ④ clone → ⑤ notify); on a missing layout, `layout_creator.clone_for_over_limit` performs the 7-step XML/part clone into a derived `template_new.pptx` (AC2; the base is never written), returning `config_overrides` that pin the slide type to the cloned layout so the fill engine renders the whole deck in one pass against the active template (AC3). Clone failure is caught and falls back to the base template — the deck still renders (AC4); the base stays immutable (reload-verify + rollback-delete). The mandatory `build_notification` names the derived template + reason whenever it is used (AC5). Default policy `clone_on="missing"` handles only missing layouts; over-limit content is handled by density downshift (`clone_on="any"` is the opt-in that also clones for over-limit). This capability is **beyond chenyu's original 2-skill scope** (a stakeholder requirement, `DESIGN-template-agnostic.md` §7), retroactively entered into the story set. With the bundled template (all 8 slide types servable) this path is rarely exercised — it matters mainly for incomplete user-supplied templates. Tests in `template-modifier-skill/scripts/tests/`.

### Epic 4 — Engineering Foundations

#### US-5.1 — Two Independent Skills with CLI Scripts `[Must Have]` — 🟡 Partial

The skills are `generate-slide-skill` (fill) and `template-modifier-skill` (extend) — **not** the `generate-template`/`generate-slides` decomposition the story names, and the manifests are `SKILL.md` (no `skill.yaml`) → AC1 not met. **Both engine scripts are now standalone CLIs (Rev 16):** `schema_extractor.py` (`--input/-i`, `--output/-o`, `--log-level`) and **`ppt_builder.py` `main()`** (argparse: `--template/-t`, `--data/-d`, `--output/-o`, `--target-size`, `--log-level`; US-4.6) — each with documented exit codes 0/1/2 (success/validation/runtime) → **AC2 (CLI-runnable) Met** and **AC3 (exit codes) Met**. The Epic 4 `_common/` refactor (PLAN-GIT-72) additionally made `template-modifier-skill`'s production code **zero-coupling** to `generate-slide-skill` (shared infra in `_common/`). Graded Partial: AC1 (the "exactly 2 skills" decomposition + `skill.yaml` manifest) remains unmet — both scripts-as-CLIs + exit-code consistency are delivered.

#### US-5.2 — Shared JSON Schema for Validation `[Must Have]` — 🟡 Partial

`slide_schemas.py` validates **slide content** for all 8 slide types + `chart_options` (`schema_validator.py` `validate_slide_data_list` / `parse_and_validate`). The **extraction** side: `template_schema.json` (JSON Schema draft-2020-12) now lives in the **shared `_common/scripts/schemas/`** home (PLAN-GIT-72 — the story's "shared `common/` directory" design, delivered) alongside `validate_template_schema()` (in `schema_extractor`, also in `_common`), shared by all three skills. The extract→embed CLI validates the schema before embedding (AC2 Met) and `schema_version` is tracked (`SCHEMA_VERSION = "1.1.0"` in `_build_metadata` → AC4 Met). However, the schema file is **not loaded at runtime** — `validate_template_schema()` is hand-rolled (no `jsonschema` dependency), so spec and validator are kept in sync manually (a pre-existing divergence: `additionalProperties:false` and the `id` `pattern` in the schema are not enforced). Graded Partial: shared home + version + pre-embed validation delivered (AC2/AC4); schema-driven runtime validation (AC1/AC3) still open.

#### US-5.3 — Structured Logging `[Should Have]` — 🟡 Partial

Python's `logging` module is used across the modules. `schema_extractor.py` **does** expose a `--log-level` flag (debug/info/warn/error) applied via `logging.basicConfig`. But logging is **not JSON-lines structured** (plain `%(asctime)s [%(levelname)s] %(message)s` format), output is not explicitly routed to stderr-only, and only `schema_extractor` exposes the flag (the engine modules don't). Graded Partial: the flag exists for the extractor, but structured JSON-lines + stderr-only routing are not implemented.

---

## §3 Gap Summary

### §3.1 Statistics

| Status | Count | Stories |
|---|---|---|
| ✅ Met | 16 | US-1.1, US-1.2, US-1.3, US-1.4, US-1.5, US-2.1, US-3.1, US-3.2, US-3.3, US-3.4, US-4.1, US-4.3, US-4.5, US-4.6, US-4.7, US-6.1 |
| 🟡 Partial | 4 | US-4.2, US-5.1, US-5.2, US-5.3 |
| ❌ Not met | 0 | — |
| ⚪ Architecture differs | 0 | — |

### §3.2 Priority × Status Matrix

| | Must Have | Should Have | Could Have |
|---|---|---|---|
| ✅ Met | US-1.1, US-1.2, US-1.3, US-1.4, US-1.5, US-2.1, US-3.1, US-3.2, US-3.3, US-4.1, US-4.3, US-4.7 | US-3.4, US-4.6, US-6.1 | US-4.5 |
| 🟡 Partial | US-4.2, US-5.1, US-5.2 | US-5.3 | — |
| ❌ Not met | — | — | — |
| ⚪ Differs | — | — | — |

**The extraction stories (US-1.x), the templating-skill stories (US-3.x), and US-4.1 are complete** (Rev 9); **US-4.2 delivered** (Rev 10, AC2/AC3 Met, AC1 deferred); **US-4.3 delivered** (Rev 11); **US-2.1 delivered** (Rev 13 — all Must-Have stories are now Met or Partial); **US-4.6 delivered** (Rev 14 — multi-aspect-ratio via coordinate-path prep + shared native loop; all 5 ACs Met); **US-6.1 added** (Rev 15 — Epic 3 "Template Extension", a capability beyond the original 2-skill scope, Met); **Epic 4 `_common/` refactor delivered** (Rev 16 — shared infra extracted to `_common/`, both engine CLIs now exist; US-5.1 AC2/AC3 + US-5.2 AC2/AC4 now Met, both stories still Partial pending the "exactly 2 skills" decomposition / runtime `jsonschema` loading); **US-4.7 added** (Rev 18 — default template moved to repo-root `template/default.pptx`, user paths passed through `template_path`, `TemplateError` pre-flight for severe template problems; all 3 ACs Met). The remaining gaps cluster in Epic 2 (US-4.2's deferred AC1 overflow-oracle) and Epic 4 (US-5.1 skill.yaml/2-skill decomposition, US-5.2 schema-driven runtime validation, US-5.3 JSON-lines logging). **No fully-unmet Must-Have remains.** *(Rev 19 attempted to close US-4.2 AC1 via PowerPoint-native `normAutofit` (#74) but was **empirically falsified** — `<a:normAutofit/>` is inert on file open; reverted, AC1 stays deferred. Rev 20 descoped US-2.2 — common-practice suggestions removed from the backlog.)*

---

## §4 Open Work

Only stories that are **not yet Met** (done items are documented in §2; this section no longer repeats them).

| Story | Status | Remaining work |
|---|---|---|
| **US-4.2** — Text Fitting (AC1) | 🟡 Partial (Must-Have) | AC1's hard "no overflow" guarantee needs a render oracle (LibreOffice headless render-to-image); the best-effort `text_fit.py` estimator is already shipped. *(#74's PowerPoint-`normAutofit` path was falsified — inert on open — and reverted.)* |
| **US-5.1** — Skill Decomposition | 🟡 Partial (Must-Have) | AC1 only: manifests are `SKILL.md` not `skill.yaml`, and 3 skills exist (not the story's "exactly 2"). AC2/AC3 Met — both engine CLIs exist (Rev 16). |
| **US-5.2** — Shared JSON Schema | 🟡 Partial (Must-Have) | Load `template_schema.json` at runtime (or generate the validator from it) so `additionalProperties:false`/`pattern` are enforced. AC2/AC4 Met (shared `_common/` home + `schema_version` + pre-embed validation). |
| **US-5.3** — Structured Logging | 🟡 Partial (Should-Have) | Switch to JSON-lines (timestamp/level/skill/action) routed to stderr; extend `--log-level` to the engine modules. |

**No fully-unmet Must-Have remains** — all remaining items are Partials with most ACs already satisfied. *(US-2.2 — the only Not-met — was descoped in Rev 20.)*

---

## §5 Decisions (all resolved)

| # | Decision | Resolution |
|---|---|---|
| 1 | Coexist / Replace / Update-requirements | ✅ **Coexist — implemented.** `schema_extractor.py` (proposed-schema) coexists with the fingerprint-contract engine; the renderer consumes the contract. The polygon↔contract bridge is `contract_adapter` (US-4.1). |
| 2 | Who consumes the polygon schema | ✅ **Resolved.** chenyu's #4 = "use the slide master's layouts" (`add_slide`) — coordinate placement was never requested. The polygon model (US-1.2) is the faithful self-description; layout selection is via `add_slide`. Coordinate placement is scoped to the one case that genuinely needs it — **US-4.6** multi-aspect-ratio (size mismatch). US-4.1's source-swap is faithful, not a compromise. |
| 3 | Must-Have priority confirmation | ✅ **Resolved.** No fully-unmet Must-Have remains (US-2.1 delivered, Rev 13). Open items: Partials US-4.2 / US-5.1 / US-5.2 / US-5.3 — all with most ACs already satisfied (see §4). *(US-2.2 descoped — Rev 20.)* |

The original option-by-option prose for Decisions 1–3 is preserved in Appendix A.

---

## Appendix A — Revision History (full narrative)

> The compact one-line-per-revision table is at the top of this document. The full per-revision narrative is preserved verbatim below as the audit trail.

> **Revision 2 (post-US-1.1, PR #49):** US-1.1 now ✅ Met (new `schema_extractor.py` emits the proposed-schema JSON). US-1.2 → 🟡 Partial (polygon field now exists + normalized, but cross-product winding check pending). US-1.3 improved (full 10-value enum on all elements). Counts updated: Met 2 / Partial 7 / Not met 9.
>
> **Revision 3 (post-US-1.2, PR #51):** US-1.2 now ✅ Met — cross-product winding check delivered (`_signed_area()` + check in `validate_template_schema()`; canonical TL→TR→BR→BL = algebraic CCW = anti-clockwise). Counts: Met 3 / Partial 6 / Not met 9.
>
> **Revision 4 (post-US-1.3, issue #52):** US-1.3 now ✅ Met — `type_confidence` is always emitted (`"high"` default, `"low"` only for `shape_type` None/unreadable and indeterminate MEDIA; no whitelist per architecture-review MAJOR-1), `"audio"` is now reachable via OOXML `<a:audioFile>`/`<a:videoFile>` split, `WEB_VIDEO`→`video/high`, and a non-fatal WARNING surfaces `shape/low` (MINOR-2). Optional `type_confidence` added to `template_schema.json`. Counts: Met 4 / Partial 5 / Not met 9.
>
> **Revision 5 (audit-path consistency):** Re-graded Epic 3/5 stories on the **same path used for Epic 1** (`schema_extractor.py`, the proposed-schema extractor) instead of the renderer contract. Corrections: **US-3.2** ❌→🟡 (`_infer_title` emits `title`; AC1 met, AC2/AC3 not), **US-3.4** 🟡→✅ (`_build_theme` maps semantic roles + `font_palette`; all 3 ACs met), **US-5.1** ❌→🟡 (`schema_extractor` is a real CLI with `--input/--output/--log-level` + exit codes 0/1/2; AC3 met, but skill names/decomposition differ), **US-5.2** 🟡 (clarified: `schema_version` IS in the proposed schema; `template_schema.json` exists but is not loaded at runtime), **US-5.3** 🟡 (clarified: `--log-level` DOES exist in `schema_extractor`; still not JSON-lines). Counts: Met 5 / Partial 6 / Not met 7.
>
> **Revision 6 (post-US-1.4, issue #54):** US-1.4 now ✅ Met — `_extract_text_fonts` populates per-textbox `font` (explicit-only, Latin-only) + nested `runs[]`, with a guarded RGB `color`; deduped `missing_fonts[]` against a curated `_BUILTIN_FONTS` allowlist (theme-aware `fallback`, AC4 → validator ERROR); `validate_template_schema` emits a non-fatal WARNING per missing font (AC3). Counts: Met 6 / Partial 6 / Not met 6.
>
> **Revision 7 (post-US-1.5, issue #55):** US-1.5 now ✅ Met — `embed_schema` writes the schema into the PPTX zip at `ppt/template_schema.json` via an order-preserving rewrite (`[Content_Types].xml` first + injected `json` Default; decompressed-content-identical originals; idempotent; atomic; `EmbeddedSchemaResult` for AC4); `read_embedded_schema` retrieves it; CLI `--embed` + `--output-pptx`. **Epic 1 complete** (all 5 stories Met). Counts: Met 7 / Partial 6 / Not met 5.
>
> **Revision 8 (post-US-3.1, issue #56):** US-3.1 now ✅ Met — a standalone `generate-template-skill` (`.opencode/skills/generate-template-skill/SKILL.md`) orchestrates the full `extract → validate → (title confirm) → embed → return templated PPTX + summary` pipeline; `pptx-subagent.md` gains a one-line "What NOT to Handle" deferral fixing the NL-routing collision (architecture review MAJOR-1). US-3.2 🟡→✅ — `_infer_title` now returns a `TitleInference(title, source)` NamedTuple and `_build_metadata` emits `title_source`; the skill prompts the user when `title_source == "filename"` and always displays the title for confirmation (AC2/AC3). US-3.3 ❌→✅ (corrects the Rev-7 stale rating — US-1.5 already delivered embed + the round-trip test; this issue adds the downloadable-PPTX surface via the skill + `build_extraction_summary` + CLI `--summary`). `title_source` is runtime-enforced by `validate_template_schema` (MAJOR-2), closing the US-5.2 gap for that field. **Epic 3 complete** (all 4 stories Met). Counts: Met 10 / Partial 5 / Not met 3 / Differs 1.
>
> **Revision 9 (post-US-4.1, issue #58):** US-4.1 now ✅ Met — the renderer reads the embedded `ppt/template_schema.json` via a source-swap adapter (`contract_adapter.embedded_schema_to_contract` → `ppt_builder.get_render_contract`), preferring embedded JSON and falling back to the sidecar. Generation keeps `add_slide(layout)` (chenyu's #4); the embedded JSON drives layout selection, not coordinate placement (AC3 clarified — see US-4.1 historical note; coordinate placement is deferred to US-4.6). Architecture-review findings addressed: C1 (clone path re-embeds the schema into `template_new.pptx`), M3 (`chart`→`OBJECT` canonical map + top-level placeholder filtering; parity-tested), M4 (`_source` provenance + absent/corrupt-distinction), M5 (templated default template + staleness guard), M6 (adapter as a bridge → US-4.6), m1/m2/m3 (consumer migration + grep audit). The bundled `template.pptx` now ships pre-templated. Counts: Met 11 / Partial 4 / Not met 3 / Differs 1.
>
> **Revision 10 (post-US-4.2, issue #60 — delivered & merged via #62):** US-4.2 delivered — a pure `text_fit.py` estimator (−2pt steps to an 8pt floor) + a sidecar `<output>.render.json` carrying the `font_size_adjusted` flag. **AC2 ✅ & AC3 ✅ Met; AC1 deferred (best-effort).** Architecture review (APPROVE-WITH-CHANGES) incorporated: M1 (base-size resolution chain → schema `size_pt` → **layout sample-run** → conservative role ceiling body **14** / title **28** / subtitle **18**; AC2 re-framed to "≤ resolved base"), M2 (inter-paragraph spacing reserve in the height estimate), M3 (explicit `run.font.size` written **only on actual shrink**, else inheritance preserved; body is the documented exception — always template-derived). Code review (APPROVE-WITH-CHANGES) incorporated: a Major fix to `_layout_line_spacing` (exact-point `Length`/int-subclass spacing no longer misread as a multiplier), a DRY refactor, and +7 coverage tests. **Finding C1 is deferred by design** (see the US-4.2 caveat below): AC1's hard "no overflow" guarantee is **not verifiable** without a layout engine, so it is delivered best-effort and a full oracle (LibreOffice headless render) is left to a follow-up; AC1 stays unchecked. US-4.2 stays 🟡 Partial (AC1 pending). The bundled-template mis-fingerprint defect (#61) was also resolved by shipping a pre-templated replacement template with real OBJECT/BODY/PICTURE placeholders. Full suite 381 passed. Counts: Met 11 / Partial 4 / Not met 3 / Differs 1 (unchanged — US-4.2 remains Partial pending AC1).
>
> **Revision 11 (post-US-4.3, issue #63):** US-4.3 reclassified ⚪ Architecture differs → ✅ Met. AC1 was already functionally met post-US-4.1 (sidecar fallback). The real gaps — AC2 (output carries embedded JSON) and AC3 (status message) — are closed: `generate_ppt_from_data(auto_template=True)` re-embeds `ppt/template_schema.json` into the **output** after save (python-pptx otherwise strips the part), sourcing the schema from the **input template** (arch-review M1 — the title is the template's identity, not the rendered deck's cover) and skipping a stale embedded input schema (M2); the agent detects a non-templated input at Stage 0 (`read_embedded_schema`, exception-safe per m6) and emits the one-line status message. Every output `.pptx` is now a self-describing/templated deck; the render report gains an additive `templating` field. The interactive generate-template-skill is not chained (mechanism differs, function met). Full suite 389 passed. Counts: Met 12 / Partial 4 / Not met 3 / Differs 0.
>
> **Revision 12 (skill rename):** the `ppt-template-filler` skill was renamed to **`generate-slide-skill`** — the slide-generation engine that realizes chenyu's `generate-slides` skill (Epic 4 / US-5.1). The directory was `git mv`'d and **every reference updated** (agent `permission.task`, cross-skill `sys.path` in `template-modifier-skill`, all SKILL.md/README/AGENTS/path strings). This **narrows the US-5.1 naming divergence** (the skill name now matches chenyu's `generate-slides` intent, modulo the `slide`/`slides` wording and the lack of a standalone CLI). Historical revision paragraphs above that now read `generate-slide-skill` refer to the same skill under its former name `ppt-template-filler`. Suites green: 396 + 33. Counts unchanged.
>
> **Revision 13 (post-US-2.1, issue #68):** US-2.1 reclassified ❌ Not met → ✅ Met. `_detect_header_footer(prs)` scans the slide master for HEADER/FOOTER placeholders and records `{has_header, has_footer}` in `template_metadata.header_footer` (AC1). `needs_header_footer_prompt(schema)` returns True when both absent (AC2 — `generate-template-skill` Stage 2 prompts, batched with title-confirm per arch-review M2; `pptx-subagent` Stage 0 surfaces a light note via `read_embedded_schema`, scoped to templated inputs per arch-review M1). `inject_default_header_zone(schema)` injects a 4-point top-strip polygon + English note (AC3, schema-only). **No fully-unmet Must-Have remains.** Full suite 405 passed. Counts: Met 13 / Partial 4 / Not met 2 / Differs 0.
>
> **Revision 14 (backfill — US-4.6, PLAN-GIT-70):** US-4.6 — Multi-Aspect-Ratio Rendering `[Should Have]` — ✅ Met. Delivered via a coordinate-path **prep** step (resize canvas + proportionally rescale every master/layout shape) then the shared native render loop; all 5 ACs Met; a ratio no-op gate keeps the native path when the target ratio matches; the output's embedded `slide_dimensions` is rewritten to the target size (self-describing). This entry backfills the missing revision-log record (§3.1 already counted it). Epic 4 grew to 6 stories when US-4.6 was added — header count corrected 19 → 20. Counts: Met 14 / Partial 4 / Not met 2.
>
> **Revision 15 (US-6.1 added — Epic 6 inserted, method B):** new story **US-6.1 — Extend Template When a Layout Is Missing `[Should Have]` — ✅ Met**, retroactively entering the story set. Realized by `template-modifier-skill` (Capability B; beyond chenyu's original 2-skill scope — stakeholder requirement per `DESIGN-template-agnostic.md` §7). A new **Epic 6 — Skill: Template Extender** is inserted between Epic 4 and Epic 5 (placement option B: three skill-Epics grouped 3/4/6, with Epic 5 "Architecture" retained as the concluding meta-Epic); no existing story is renumbered (stable-ID convention preserved — the `chenyu-user requirement.html` source and all prior commits/issues keep their original US-5.x references). Counts: Met 15 / Partial 4 / Not met 2 (21 stories, 6 Epics).
>
> **Revision 16 (post-US-4.6 + Epic 5 `_common/`, PLAN-GIT-70 + PLAN-GIT-72):** two Epic 5 advances (both stories still 🟡 Partial — **counts unchanged**). **US-5.1**: `ppt_builder.py` now ships a real argparse CLI (`main()` with `--template/--data/--output/--target-size/--log-level`, exit codes 0/1/2 — delivered with US-4.6) → **AC2 + AC3 now Met** (both engine scripts are CLI-runnable with documented exit codes); and `template-modifier-skill`'s production code now has **zero coupling** to `generate-slide-skill` (the shared extraction/contract/schema infra was extracted to `.opencode/skills/_common/scripts/`, killing the sibling-skill `sys.path` hack). AC1 still unmet (manifests are `SKILL.md` not `skill.yaml`; 3 skills exist, not the story's "exactly 2"). **US-5.2**: the story's "shared `common/` directory" is now real — `template_schema.json` + `validate_template_schema` (in `schema_extractor`) live in `_common/scripts/` and are shared by all three skills; the extract→embed CLI validates before embedding → AC2 Met; `schema_version` tracked (1.1.0) → AC4 Met. Still Partial: the schema file is **not loaded at runtime** (validator is hand-rolled, no `jsonschema` dependency), so AC1/AC3 (schema-driven runtime validation) remain open. Counts: Met 15 / Partial 4 / Not met 2.
>
> **Revision 17 (scope change — US-4.4 removed):** US-4.4 — Style Picker for Template-Less Generation `[Should Have]` — **descoped** (removed from the requirement set, not delivered). Current counts: **Met 15 / Partial 4 / Not met 1** (US-2.2 only). Prior revision entries (Rev 9–16) are preserved verbatim as the historical record — their "Not met 2" counts reflected US-4.4's existence at that time. Epic 4 now comprises 5 stories (US-4.1, US-4.2, US-4.3, US-4.5, US-4.6).
>
> **Revision 18 (US-4.7 added — template selection & pre-render validation):** new story **US-4.7 — Template Selection & Pre-Render Validation `[Must Have]` — ✅ Met**, retroactively entering the story set. The default template moved from the deep `scripts/templates/template.pptx` path to the repo root as **`template/default.pptx`** (used whenever `template_path` is omitted — mirrors `output/`); a user-supplied `.pptx` path is passed straight through `template_path`/`--template` (the earlier copy-overwrite workflow is retired). A new `TemplateError` + `_validate_template()` pre-flight runs on every load (default or user-supplied): corrupt/non-PPTX, no slide master, zero layouts, or serving none of the 8 slide types → clear error + abort before the render loop; minor issues (missing fonts, no header/footer, small content area, no embedded schema) stay non-fatal warnings. CLI maps `TemplateError` to exit 1. All 3 ACs Met (+12 tests in `test_template_validation.py`). Header count corrected 20 → 21. Counts: Met 16 / Partial 4 / Not met 1.
>
> **Revision 19 (US-4.2 AC1 attempt — #74 normAutofit — FALSIFIED & REVERTED):** an attempt to close US-4.2 AC1 via PowerPoint-native `normAutofit` (`<a:normAutofit/>`, no `fontScale`) was **empirically falsified**. The premise (that PowerPoint recomputes the shrink on file open — "behaviour A") was disproven by opening a verification deck in PowerPoint: text visibly overflowed the placeholder boxes on **all** content slides, **including the control slides** whose text the `text_fit` heuristic had deemed "fitting" (`fits=True`). Conclusion: `<a:normAutofit/>` without a pre-computed `fontScale` is **inert on file open** — PowerPoint only recomputes the shrink when the user edits the text frame. The #74 code changes were **fully reverted** (helper, tuple return, render-report field, 6 call sites, 5 tests); US-4.2 returns to 🟡 Partial (AC1 deferred). Counts unchanged from Rev 18: Met 16 / Partial 4 / Not met 1. **Side-finding:** the control-slide overflow reveals `text_fit.py` *systematically underestimates* rendered height (suspect: text-frame internal margins ~0.1in/side and line-spacing not modelled) — tracked as a follow-up research issue (independent of the AC1 oracle decision). Lesson: the architecture-review's MAJOR-1 linchpin concern was the correct one; the accepted 85%-confidence bet landed in the 15% failure tail.
>
> **Revision 20 (US-2.2 descoped — common-practice suggestions removed):** **US-2.2 — Common Practice Suggestions `[Should Have]` — descoped & removed from the backlog.** The story (suggest common PowerPoint practices — slide numbers, company logo, consistent margins, section dividers, closing slide) will not be implemented. Rationale: the requirements cross-check (this session) confirmed US-2.2 is requirement-grounded (req 2: *"what other common practices suitable for powerpoint slide"*), but the user chose to abandon it — the detection heuristics (especially "consistent margins" and "company logo") are weak/subjective for a Should-Have, and the value is low relative to the remaining Must-Have Partials. The placeholder `common_practices` schema field (added in US-1.1 in anticipation of US-2.2) and the `_build_metadata` emit were **removed** as dead code; the `chenyu-user requirement.html` source is left untouched (historical contract). The story entry was **fully removed** from `user-stories.md`/`.zh.md` (no marker kept — the backlog stays clean); this GAP-ANALYSIS Rev 20 entry is the only historical trace (the US-4.4 Rev 17 precedent kept a marker, but the user chose a clean removal here). Requirement point 2c ("other common practices") is now intentionally not addressed. Counts: **Met 16 / Partial 4 / Not met 0** (20 stories, 4 Epics).

---

*End of analysis. This document contains no code changes; it records the current state and suggests directions for a future decision.*
