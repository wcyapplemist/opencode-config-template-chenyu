# Gap Analysis — chenyyu-user-stories vs Current Implementation

> **Document type:** Requirements-vs-implementation gap analysis (analysis only — no code changes)
> **Requirements source:** `docs/user-stories/chenyu-user-stories.md` (5 Epics, 19 Stories)
>
> *Note: the source document's own header states "17 Stories", but it actually contains 19 (Epic 1: 5, Epic 2: 2, Epic 3: 4, Epic 4: 5, Epic 5: 3). This report counts 19.*
> **Implementation audited:** `.opencode/skills/generate-slide-skill/`, `.opencode/skills/template-modifier-skill/`, `.opencode/agents/pptx-subagent.md`
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

---

## Executive Summary

This report compares `chenyu-user-stories.md` (the requirements document) against the current implementation. **Core conclusion: the architectural route has diverged.**

- The requirements document describes a **"normalized polygon JSON Schema embedded in PPTX"** architecture: each component contains a 4-point 0–1 normalized `polygon`, a `type` enum, `font`/`runs`, and JSON embedded inside the zip at `ppt/template_schema.json`.
- The current implementation follows a **"introspection contract with placeholder fingerprint matching"** architecture: placeholders use absolute inch coordinates, the contract is stored as a sidecar file `template.pptx.contract.json`, and layouts are matched by fingerprint.

Both achieve "fill any template", but they differ on **data model, skill decomposition, and artifact form**.

**Story-by-story summary** (19 stories): Met 11 / Partial 4 / Not met 3 / Architecture differs 1.

**Epics 1, 3, and US-4.1 are now complete** (Rev 9). The remaining gaps cluster in **Epic 2** (header/footer detection + common practices), the rest of **Epic 4** (US-4.2 text-fitting, US-4.4 template-less style picker, US-4.6 multi-aspect-ratio), and **Epic 5** (skill decomposition / CLI / shared-schema runtime loading). The single unmet Must-Have is **US-2.1** (header/footer).

For the full comparison see §2; for remediation suggestions see §4.

---

## §1 Architecture Comparison

The requirements document and the implementation describe two different routes to the same goal ("fill any template"). They diverge on **data model, skill decomposition, artifact form, and layout-matching strategy**.

| Dimension | Requirements (chenyu-user-stories) | Current Implementation |
|---|---|---|
| **Skill decomposition** | Exactly 2 skills: `generate-template` (extraction) + `generate-slides` (generation), each a standalone CLI script | 2 skills + 1 agent: `generate-slide-skill` (filling) + `template-modifier-skill` (extension) + `pptx-subagent` (content strategy). No standalone CLI entry points. |
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

#### US-1.5 — JSON Storage Inside PPTX Zip `[Must Have]` — ✅ Met

**Met (issue #55).** `schema_extractor.embed_schema()` (`schema_extractor.py`) writes the schema into the PPTX zip at `ppt/template_schema.json` via an **order-preserving full rewrite**: `[Content_Types].xml` first (with a `<Default Extension="json" ContentType="application/json"/>` injected — strict-safe, architecture review MAJOR-2), every other original entry **decompressed-content-identical** in original order (MAJOR-1; AC3), then the minified schema appended. **Idempotent** (re-embed replaces, never duplicates — MAJOR-3), **atomic** (temp + `os.replace` — MINOR-6), returns an `EmbeddedSchemaResult` (AC4 size delta — MINOR-5). `read_embedded_schema()` retrieves it with a clear error contract (absent→`None`; malformed→`None`+warn; non-zip→`TemplateExtractionError` — MINOR-3/4). CLI `--embed` + `--output-pptx` (additive; without-`--embed`→exit 2). AC1 verified by proxy (`python-pptx` re-opens; `[Content_Types].xml` first + declares `json` + originals intact; all other entries hash-identical — real PowerPoint is manual). 93 tests in `test_schema_extractor.py` (was 80). → **Epic 1 complete**.

### Epic 2 — Header, Footer & Best Practices

#### US-2.1 — Header & Footer Detection `[Must Have]` — ✅ Met (Rev 13)

`_detect_header_footer(prs)` scans the slide master's placeholders for HEADER/FOOTER types and records `{has_header, has_footer}` booleans in `template_metadata.header_footer` (AC1). `needs_header_footer_prompt(schema)` returns True when both are absent (AC2). `inject_default_header_zone(schema)` injects a 4-point top-strip polygon + English note into the schema (AC3, schema-only). `generate-template-skill` Stage 2 prompts the user (batched with title confirmation per arch-review M2) and injects on "yes"; `pptx-subagent` Stage 0 surfaces a light note via `read_embedded_schema` for templated inputs (arch-review M1: `get_render_contract`/adapter strips `template_metadata`, so `read_embedded_schema` is the only accessor; non-templated inputs defer the note). 9 tests; full suite 405 passed.

#### US-2.2 — Common Practice Suggestions `[Should Have]` — ❌ Not met

No `common_practices` object exists in the contract. There is **no checks** for slide numbers, company logo, consistent margins, section dividers, or closing slide, and no `suggestions` array.

### Epic 3 — Skill — Template Generator

#### US-3.1 — End-to-End Template Generation Pipeline `[Must Have]` — ✅ Met (Rev 8)

A standalone `generate-template-skill` (`.opencode/skills/generate-template-skill/SKILL.md`) now orchestrates the full pipeline end-to-end via the `schema_extractor` engine: `extract → validate → (title confirm) → embed → return templated PPTX + summary`. NL intent routing is via the SKILL.md `description` (extraction verbs) + a one-line "What NOT to Handle" deferral in `pptx-subagent.md` (architecture review MAJOR-1 — the agent's greedy `pptx` triggers would otherwise have misrouted extraction requests). All three ACs met.

#### US-3.2 — Template Naming `[Must Have]` — ✅ Met (Rev 8)

`_infer_title` now returns a `TitleInference(title, source)` NamedTuple and `_build_metadata` emits `template_metadata.title_source` (`core_xml` | `slide1` | `filename`); the skill prompts the user to name the template when `title_source == "filename"` and writes back `title_source = "user"` on an override (AC2 — the inference order now ends in a **user prompt**, not just a filename). The skill always displays the title for confirmation (AC3). AC1 (non-empty title via the filename fallback) was already met. All three ACs met.

#### US-3.3 — Return Downloadable Templated PPTX `[Must Have]` — ✅ Met (Rev 8)

**Corrects the Rev-7 stale rating.** US-1.5 already delivered `embed_schema` (produces the downloadable templated PPTX) and the round-trip test (`test_round_trip_deep_equal`). This issue adds the skill surface (the downloadable PPTX is returned via `output/<stem>.templated.pptx`, AC1) and `build_extraction_summary(schema) -> str` + CLI `--summary` (AC2 — a human-readable summary of layouts/components/fonts/theme). AC3 (round-trip) was already met. All three ACs met.

#### US-3.4 — Theme & Color Extraction `[Should Have]` — ✅ Met

The renderer's `_build_theme()` (`template_introspector.py`) extracts only raw OOXML role colors. **But the proposed-schema path fully implements it**: `schema_extractor._build_theme()` maps raw colors to semantic roles (`primary_color`/`secondary_color`/`accent_color`/`background_color`/`text_color`) and builds `font_palette.{heading,body,accent}`; `_raw_theme_colors_and_fonts()` is wrapped in try/except that logs a warning and yields empty defaults on a missing/malformed theme. All three ACs (semantic colors as hex; `font_palette`; sensible defaults + warning) are met on the `schema_extractor` path.

### Epic 4 — Skill — Slide Generator

#### US-4.1 — Read Embedded JSON as Layout Reference `[Must Have]` — ✅ Met (Rev 9)

The renderer now reads the **embedded** `ppt/template_schema.json` (not the sidecar) via `ppt_builder.get_render_contract` → `contract_adapter.embedded_schema_to_contract` (US-4.1, issue #58). Layout selection still uses `layout_name`/fingerprint matching (`_resolve_layout_by_fingerprint`); generation keeps `add_slide(layout)` (chenyu's #4 — "using the slide master's slide template"). AC3 ("within 1%") was clarified: coordinate placement was never required (see the US-4.1 historical note + §5 Decision 2); it is deferred to US-4.6 (multi-aspect-ratio). Architecture-review findings C1/M3/M4/M5/M6/m1-m3 all addressed. All three ACs met.

#### US-4.2 — Visually Pleasing Output with Text Fitting `[Must Have]` — 🟡 Partial (AC2/AC3 delivered; AC1 verification deferred)

Overflow prevention is **preventive, not reactive**: `density_mode.py` enforces a per-slide word budget (`DENSITY_BUDGETS`, `density_mode.py:48`) and `constraint_checker` checks `content_area_in2`. But there is **no reactive per-textbox font auto-shrink** (−2pt steps until fit) — a grep for `font_size`/`auto.*shrink`/`adjust` in `density_mode.py` returned no matches. No `font_size_adjusted` flag is set.

> **Delivered (issue #60, merged via #62 / commit 2a5d30d on `dev`; plan `PLANS/GIT-60`):** the reactive layer is a pure `text_fit.py` estimator (−2pt steps to an 8pt floor) + a sidecar `<output>.render.json` carrying the `font_size_adjusted` flag (AC3 ✅). Body base size is now template-derived (M1 chain: schema `size_pt` → layout sample-run → conservative role ceiling body 14 / title 28 / subtitle 18), replacing the `Pt(14)`/`Pt(12)` hardcode (AC2 ✅, re-framed to "≤ resolved base"). An explicit `run.font.size` is written only on actual shrink (M3; body always template-derived); an auto-grow guard prevents false-shrinking on short-base-height placeholders (titles). Inter-paragraph spacing is reserved in the height estimate (M2). Code-review Major fix: `_layout_line_spacing` now treats exact-point `Length` (int-subclass) spacing as a fallback-to-default rather than a 228600× multiplier. Full suite 381 passed. **AC1 caveat (architecture-review C1, deferred by decision):** python-pptx has **no layout/text-measurement engine**, so AC1's hard guarantee — "no generated slide has text overflowing its bounding box" — **cannot be verified** by the engine alone. It is delivered **best-effort** this round (conservative estimator + `word_wrap=True`, which is a **horizontal-only** backstop and does **not** prevent vertical overflow). A full oracle (e.g. a LibreOffice headless render-to-image check) is **out of scope** and left to a follow-up; consequently **AC1 stays unchecked** until that verification path exists.

> **Template defect resolved (issue #61, closed):** the previously-bundled `template.pptx` (a Google-Slides export) mis-fingerprinted body/picture placeholders — its layout fingerprints showed only `TITLE`/`SUBTITLE`, never `OBJECT`/`PICTURE` — so `content_slide` / `two_content_slide` / `comparison_slide` / `content_image_slide` did not resolve. This was a pre-existing US-4.1/Epic-1 template-data defect, **not** engine code. It is now **resolved** by shipping a replacement pre-templated `template.pptx` with real OBJECT/BODY/PICTURE placeholders (all 8 slide types servable; the suite went 26 pre-existing failures → 381 passed).

#### US-4.3 — Auto-Chain Extraction When No JSON Present `[Must Have]` — ✅ Met (mechanism differs; function + all 3 ACs met, Rev 11)

Because the model is contract-based (introspection runs automatically before every render) rather than embedded-JSON-based, there is no concept of "no JSON → extract first". The two-step chain (`extract` then `generate`) does not map; introspection is always implicit. Functionally the user can hand any `.pptx` and generate from it in one step — but the *mechanism* differs from the story.

> **Delivered (issue #63, `PLANS/GIT-63`):** reclassified ⚪→✅. **AC1** was already functionally met post-US-4.1 (`get_render_contract` sidecar fallback → any PPTX renders in one call). The real gaps — **AC2** (output carries embedded JSON) and **AC3** (status message) — are now closed: `generate_ppt_from_data(auto_template=True)` re-embeds `ppt/template_schema.json` into the **output** after save (python-pptx otherwise strips the part), sourcing the schema from the **input template** (arch-review M1 — `extract_schema(template)`, so the title is the template's identity, not the rendered deck's cover) and skipping a stale embedded input schema (M2); the agent detects a non-templated input at Stage 0 (`read_embedded_schema`, exception-safe per m6) and emits *"No template found — extracting first, then generating slides..."*. The interactive generate-template-skill is **not** chained (headless-infeasible + agent `task` permission denies it); the engine does inline extract+embed instead — the "architecture differs, function met" framing. The render report gains an additive `templating` field (`input_template_embedded`, `output_templated`, `schema_source`, `message`). Every output `.pptx` is now a self-describing/reusable templated deck.

#### US-4.4 — Style Picker for Template-Less Generation `[Should Have]` — ❌ Not met

A template is **required** (`template.pptx` must exist or the engine raises `FileNotFoundError`). There are **no built-in style presets** (minimalist/corporate/creative/dark) and no style picker. A default template exists, but no preset-selection flow.

#### US-4.5 — Multi-Slide Batch Generation `[Could Have]` — ✅ Met

The multi-stage pipeline (outline → critique → detail → render) generates 2–20+ slides from one prompt, the outline is shown for user approval, and the stages act as a progress indicator. Fully satisfied.

### Epic 5 — Skill Architecture & Scripts

#### US-5.1 — Two Independent Skills with CLI Scripts `[Must Have]` — 🟡 Partial

The skills are `generate-slide-skill` (fill) and `template-modifier-skill` (extend) — **not** the `generate-template`/`generate-slides` decomposition the story names, and the manifests are `SKILL.md` (no `skill.yaml`) → AC1 not met. **But a standalone CLI does exist**: `schema_extractor.py` `main()` is a real CLI (`--input/-i`, `--output/-o`, `--log-level`) with documented exit codes 0/1/2 (success/validation/runtime) → AC3 (exit codes) met and AC2 (runnable from CLI independently of the LLM) met *for the extractor*. `ppt_builder.py` remains function-based (its `main()` is a demo, no argparse). Graded Partial: one of the two intended skills exists as a CLI with exit codes, but the skill decomposition/naming differs and the slide side is not a CLI.

#### US-5.2 — Shared JSON Schema for Validation `[Must Have]` — 🟡 Partial

`slide_schemas.py` validates **slide content** for all 8 slide types + `chart_options` (`schema_validator.py` `validate_slide_data_list` / `parse_and_validate`). The **extraction** side: `schemas/template_schema.json` (JSON Schema draft-2020-12) exists as the conformance spec and the proposed schema **does** carry `schema_version` (`SCHEMA_VERSION = "1.0.0"`, emitted in `_build_metadata`). However, the schema file is **not loaded at runtime** — `validate_template_schema()` is hand-rolled (no `jsonschema` dependency), so spec and validator are kept in sync manually (a pre-existing divergence: `additionalProperties:false` and the `id` `pattern` in the schema are not enforced). Graded Partial: the spec + version exist, but runtime validation is hand-rolled, not schema-driven.

#### US-5.3 — Structured Logging `[Should Have]` — 🟡 Partial

Python's `logging` module is used across the modules. `schema_extractor.py` **does** expose a `--log-level` flag (debug/info/warn/error) applied via `logging.basicConfig`. But logging is **not JSON-lines structured** (plain `%(asctime)s [%(levelname)s] %(message)s` format), output is not explicitly routed to stderr-only, and only `schema_extractor` exposes the flag (the engine modules don't). Graded Partial: the flag exists for the extractor, but structured JSON-lines + stderr-only routing are not implemented.

---

## §3 Gap Summary

### §3.1 Statistics

| Status | Count | Stories |
|---|---|---|
| ✅ Met | 13 | US-1.1, US-1.2, US-1.3, US-1.4, US-1.5, US-2.1, US-3.1, US-3.2, US-3.3, US-3.4, US-4.1, US-4.3, US-4.5 |
| 🟡 Partial | 4 | US-4.2, US-5.1, US-5.2, US-5.3 |
| ❌ Not met | 2 | US-2.2, US-4.4 |
| ⚪ Architecture differs | 0 | — |

### §3.2 Priority × Status Matrix

| | Must Have | Should Have | Could Have |
|---|---|---|---|
| ✅ Met | US-1.1, US-1.2, US-1.3, US-1.4, US-1.5, US-2.1, US-3.1, US-3.2, US-3.3, US-4.1, US-4.3 | US-3.4 | US-4.5 |
| 🟡 Partial | US-4.2, US-5.1, US-5.2 | US-5.3 | — |
| ❌ Not met | — | US-2.2, US-4.4 | — |
| ⚪ Differs | — | — | — |

**Epics 1, 3, and US-4.1 are complete** (Rev 9); **US-4.2 delivered** (Rev 10, AC2/AC3 Met, AC1 deferred); **US-4.3 delivered** (Rev 11); **US-2.1 delivered** (Rev 13 — all Must-Have stories are now Met or Partial). The remaining gaps cluster in the rest of Epic 2 (US-2.2 common-practice suggestions), Epic 4 (US-4.4 template-less style picker, US-4.6 multi-aspect-ratio, US-4.2's deferred AC1 overflow-oracle), and Epic 5 (skill decomposition / CLI / shared-schema loading). **No fully-unmet Must-Have remains.**

---

## §4 Remediation Suggestions

These are **suggestions only** — no code is written here. They are grouped by priority. A decision on whether to (a) build the requirements architecture alongside the current one, (b) replace it, or (c) keep the current architecture and update the requirements, is deferred to §5.

### P0 — Core Extraction Model (Must-Have, highest impact)

These four stories are the foundation everything else builds on. They define the normalized component schema the requirements center on.

| Story | Suggestion |
|---|---|
| ~~**US-1.2 Polygon**~~ | **Done (US-1.2):** `polygon` field + 4 normalized 0–1 coords + the cross-product winding check (AC3) are all delivered — see §2 (Met). *Deferred Details (not ACs):* non-rectangular actual vertices (custGeom/triangle) — polygon stays a rectangular bounding box (metadata-only, no consumer). |
| ~~**US-1.4 Fonts**~~ | **Done (issue #54):** `_extract_text_fonts` populates per-textbox `font` (explicit-only, Latin-only) + nested `runs[]`; deduped `missing_fonts[]` against `_BUILTIN_FONTS` with theme-aware `fallback` (AC4 → validator ERROR); `validate_template_schema` warns per missing font (AC3). RGB `color` guarded by `color.type==RGB`. *Out of scope:* full-cascade inheritance, CJK (`<a:ea>`/`<a:cs>`), theme-color → hex. |
| ~~**US-1.5 Zip embedding**~~ | **Done (issue #55):** `embed_schema` writes `ppt/template_schema.json` via an order-preserving rewrite (`[Content_Types].xml` first + injected `json` Default; decompressed-content-identical originals; idempotent; atomic); `read_embedded_schema` retrieves it; `EmbeddedSchemaResult` for AC4; CLI `--embed`. *(The original "without touching `[Content_Types].xml`" was revised to "inject the `json` Default" per architecture review MAJOR-2 — strict-safe; the template declares no `json` Default.)* |
| **US-2.1 Header/Footer** | Stop discarding chrome: record `has_header/has_footer` booleans + component IDs in `header_footer` metadata. Emit a user prompt when both are absent. |

### P1 — Template Generator Skill & Decomposition (Must-Have)

| Story | Suggestion |
|---|---|
| **US-3.1 + US-5.1** | ~~US-3.1 Done (issue #56):~~ a dedicated `generate-template-skill` now runs the full P0 extraction → validation → zip-embed → return, with documented exit codes (0/1/2) inherited from the `schema_extractor` CLI. *US-5.1 still Partial:* only the `generate-template` half exists — `generate-slides` (the read-JSON-then-render CLI) is not yet built. |
| ~~**US-3.2**~~ | **Done (issue #56):** `_infer_title` returns `TitleInference(title, source)`; the skill prompts the user when `title_source == "filename"` and stores the result as `template_metadata.title` / `title_source`. |
| ~~**US-3.3**~~ | **Done (issue #56):** the downloadable templated PPTX (US-1.5 embedding) is surfaced via the skill; `build_extraction_summary` + CLI `--summary` deliver the human-readable summary. The round-trip test already exists (`test_round_trip_deep_equal`). |

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

**2. Who consumes the polygon schema — CLARIFIED (2026-06-29)**
Re-confirmed against chenyu's original requirement: chenyu's #4 states generation "uses the slide master's slide template" (`add_slide(layout)`) and **never asked for coordinate placement**. The earlier reading ("the slide generator denormalizes polygons back to EMU and places OOXML at exact positions") was an over-elaboration in US-4.1's original Details/AC3 — since corrected (see the US-4.1 historical note). **Resolution:** the polygon model (US-1.2) is the faithful, portable self-description of the template that chenyu explicitly requested; the slide generator consumes the embedded JSON for layout selection/consistency and generates via the template's own layouts. Coordinate placement is **not a requirement** — neither "locked out" nor "deferred," it was simply never asked for. This makes US-4.1's source-swap approach (PLAN-GIT-58) faithful to the requirement (not a compromise), and dissolves the architecture-review M1/M2 premise (those assumed a chenyu "coordinate-driven vision" that the source text does not support).

> **Update (2026-06-29, US-4.6 scheduled):** chenyu subsequently confirmed he wants **multi-aspect-ratio output** (e.g., render 4:3 from a 16:9 template, elements scaling proportionally) — the one use case where coordinate placement is genuinely needed, because US-4.1's `add_slide(layout)` path renders only at the template's native size. This is now scheduled as **US-4.6 — Multi-Aspect-Ratio Rendering `[Should Have]`**, which uses the coordinate-placement mechanism scoped to the size-mismatch case. Coordinate placement is therefore **deferred to US-4.6** (after the Epic 4 base), not permanently foreclosed — resolving architecture-review M2. US-4.1's native-size source-swap path is unaffected. The "within 1%" accuracy bar (originally US-4.1's AC3) is relocated to US-4.6's AC, where it legitimately applies.

**3. Priority confirmation**
With US-4.1/4.3 delivered, the remaining Must-Have gaps are **US-2.1** (header/footer — the only fully-unmet Must-Have) and the partials **US-4.2** (AC1 overflow-oracle deferred), **US-5.1/US-5.2**. Are these still Must-Have, or has the embedded-schema + auto-introspection path made any of them obsolete in practice? (US-4.3, once cited here as "arguably superseded by automatic introspection", is now **Met** — delivered as engine-inline auto-templating, Rev 11.)

---

*End of analysis. This document contains no code changes; it records the current state and suggests directions for a future decision.*
