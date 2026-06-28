# PLAN-GIT-55 — US-1.5: JSON Storage Inside PPTX Zip

**Issue**: #55
**Branch**: GIT-55 (base: dev)
**Priority**: Must Have (P0)
**Status**: Implemented (v2 + code-review follow-ups; all phases complete, 95 tests green)

## Goal

Bring US-1.5 from **❌ Not met** to ✅ Met by embedding the proposed template
schema (`schema_extractor` output) **inside the PPTX zip archive** at
`ppt/template_schema.json`, so the JSON "travels with the file" and the
slide-generation skill can read it directly from any PPTX — no sidecar file,
database, or external file system needed. Today the extractor writes JSON only to
the CLI `--output` file (or stdout); there is **no `zipfile` usage anywhere** in
the scripts, and the renderer still reads the sidecar `<stem>.pptx.contract.json`
(`template_introspector.py`). This plan adds PowerPoint-safe embedding via a
full, order-preserving zip rewrite that keeps `[Content_Types].xml` first, plus a
retrieval helper and an additive, backward-compatible `--embed` CLI flag.

## Strategic Context

Closes Epic 1 (the final `[Must Have]` story after US-1.1–US-1.4). Builds on
US-1.1 (issue #48, PR #49), US-1.2 (issue #50, PR #51), US-1.3 (issue #52, PR
#53), and US-1.4 (issue #54). `schema_extractor.py` already produces a complete
proposed-schema dict (theme, slides, components, `missing_fonts`); US-1.5 wires
that dict into the PPTX itself. The embedding is **renderer-agnostic** — the
renderer's sidecar read path (`template_introspector.py`) is **unchanged**; this
issue only adds the producer side (embed) + a read helper. GAP-ANALYSIS §2 grades
this ❌ Not met (the last P0 story in §4). On completion: Met 6→7, Not met 6→5,
Epic 1 complete.

## Zip Rewrite Strategy — Order-Preserving + `[Content_Types].xml` json-Default (MAJOR-1/2)

OOXML is a ZIP package whose consumers (PowerPoint, Keynote, `python-pptx`) are
sensitive to **entry order**: `[Content_Types].xml` must be the **first** entry.
Architecture review (MAJOR-1/MAJOR-2) established two realities:

- **MAJOR-1 — "decompressed-content-identical", not byte-identical.** Python's
  `zipfile` has no public API to copy a compressed entry's raw stream; any
  read→rewrite re-compresses (different compressed bytes, identical decompressed
  content). PowerPoint reads decompressed content, so **decompressed-content-
  identity** is what satisfies AC3 ("slide content, layouts, and media are
  untouched"). The plan therefore states "decompressed-content-identical" and
  hashes `ZipFile.read(name)` output in the integrity test — never the raw
  compressed stream.
- **MAJOR-2 — the bundled template declares no `json` Default.** Direct
  inspection: `[Content_Types].xml` declares 5 Defaults (`png/rels/xml/fntdata/
  jpg`) and no `json`. An undeclared `ppt/template_schema.json` risks a strict-
  PowerPoint "repair" prompt — the #1 AC1 risk that CI proxy tests (python-pptx
  is lenient) cannot catch.

**Decision — full rewrite, order-preserving, with a `json` Default:**
1. Open the original zip read-only.
2. Write a **NEW** zip to a temp path (atomic — MINOR-6):
   - First: `[Content_Types].xml` with a **`<Default Extension="json"
     ContentType="application/json"/>` injected** (idempotent — only if not
     already present; `Default` elements go before any `Override`), parsed/
     re-serialized via `lxml.etree`. The existing 5 Defaults + all `Override`s
     are preserved. *(This deliberately modifies `[Content_Types].xml` — deviating
     from the requirement Details' "as long as `[Content_Types].xml` is not
     modified" — because strict OOXML declares every part, and AC1 "opens without
     repair" is the binding criterion. The Details' leniency claim does not hold
     for this template.)*
   - Then: every **other** original entry in **original order**, **decompressed-
     content-identical** (read via `ZipFile.read(name)`, write via `writestr`
     against a copied `ZipInfo` preserving `compress_type`/`date_time`).
     **Idempotency (MAJOR-3):** skip any original entry whose name equals
     `_EMBEDDED_SCHEMA_PATH`, so re-embedding replaces (never duplicates).
   - Last: append `ppt/template_schema.json` = the minified schema.
3. `os.replace(temp, output_pptx_path)` (atomic — no partial/corrupt output).
4. Never modify the input in place (AC3) — always output a copy.

This keeps `[Content_Types].xml` first and valid (now declaring `json`), leaves
every other part decompressed-content-identical, and makes embedding idempotent.

## Minified JSON & Size Budget

`json.dumps(schema, separators=(",",":"), ensure_ascii=False)` — compact, UTF-8,
non-ASCII preserved (no `\uXXXX` bloat). Target: **< 50 KB** for the bundled
`template.pptx` schema (a Details hint, not an AC). Written as the
`ppt/template_schema.json` zip entry.

## CLI — Additive & Backward-Compatible

`--embed` is **opt-in**; existing flags and default behavior are unchanged:

- `--embed` (store_true): when set, embed the schema into a PPTX copy **in
  addition to** (not instead of) the `--output` JSON.
- `--output-pptx <path>` (default: `<input_stem>.templated.pptx`): destination
  for the embedded PPTX.
- `--output-pptx` without `--embed` → **argparse error (exit 2)** (MINOR-2 —
  fail loudly on a likely typo, never silently ignore).
- The existing `--output` JSON behavior is **unchanged** (backward compatible).
  *(Relaxing `--output` to optional for embed-only use is deferred to US-3.x —
  MINOR-1, see Out of Scope.)*

`embed_schema` **returns** an `EmbeddedSchemaResult` (`output_path`,
`original_bytes`, `new_bytes`, `delta_bytes`) (MINOR-5); the CLI formats the
size log from it (so AC4 is testable without coupling to log text).

The renderer's sidecar read path is **not** modified here (producer-only).

## AC1 Verification — Proxy Strategy (real PowerPoint is manual)

Real PowerPoint "opens without repair prompt" cannot be asserted in CI. AC1 is
proven by **proxy**:

- (a) `python-pptx` re-opens the embedded PPTX (`Presentation(path)`) without
  error and reports the **same slide count + layout count** as the original.
- (b) `[Content_Types].xml` is still the **first** entry, is valid XML, **now
  contains** the `json` Default, and **preserves** the original 5 Defaults + all
  `Override`s.
- (c) All **other** original zip entries are **decompressed-content-identical**
  (name-set equality + per-entry hash of `ZipFile.read(name)` output); the only
  other new entry is `ppt/template_schema.json`.

The "no repair prompt in PowerPoint" guarantee is documented as a **manual**
verification step (Risks) — the proxy is what makes the rewrite trustworthy.

## Architecture Decisions (locked, v2)

1. **Zip strategy — full rewrite, order-preserving, json-Default (MAJOR-1/2).**
   `[Content_Types].xml` first (with the `json` Default injected), then every
   other original entry **decompressed-content-identical** in original order
   (preserving `compress_type` via `ZipInfo`), then `ppt/template_schema.json`.
   **Idempotent** (skip an existing schema entry). No in-place mutation.
2. **Embed path constant.** `_EMBEDDED_SCHEMA_PATH = "ppt/template_schema.json"`.
3. **Minified JSON.** `json.dumps(..., separators=(",",":"), ensure_ascii=False)`.
4. **Output a COPY — never in place.** Write to `output_pptx_path`.
5. **CLI — additive `--embed` + `--output-pptx`.** `--output-pptx` without
   `--embed` → exit 2 (MINOR-2). `--output` JSON unchanged.
6. **Retrieval helper + error contract (MINOR-3/4).**
   `read_embedded_schema(pptx_path) -> dict | None`: valid zip + present → dict;
   valid zip + absent → `None`; **malformed JSON → `logger.warning` + `None`**;
   **corrupt/non-zip input → raise `TemplateExtractionError`** (mirrors
   `extract_schema`). Also `isinstance(result, dict)`-check (non-object JSON →
   `None`).
7. **Result + size logging (AC4, MINOR-5).** `embed_schema` returns
   `EmbeddedSchemaResult(output_path, original_bytes, new_bytes, delta_bytes)`;
   the CLI logs original → new → delta from it.
8. **Atomic write (MINOR-6).** Temp file in the output dir + `os.replace` — no
   partial/corrupt PPTX on failure.
9. **AC1 proxy (real PowerPoint manual).** (a)+(b)+(c) above.
10. **Renderer unchanged.** `template_introspector.py` sidecar read is not
    modified; renderer migration to `read_embedded_schema` is deferred (Out of
    Scope, MINOR-7).

## Deliverables (all incremental edits, zero new files)

- `scripts/schema_extractor.py`: `import zipfile`, `import os`/`tempfile`;
  `_EMBEDDED_SCHEMA_PATH`; `EmbeddedSchemaResult`; `_inject_json_default(xml_bytes)
  -> bytes` (lxml); `embed_schema(pptx_path, schema, output_pptx_path) ->
  EmbeddedSchemaResult` (rewrite + json-Default + idempotent skip + minify +
  atomic write); `read_embedded_schema(pptx_path) -> dict | None` (error contract
  MINOR-3/4); CLI `--embed` + `--output-pptx` (--output-pptx-without-embed →
  exit 2).
- `scripts/tests/test_schema_extractor.py`: round-trip; `[Content_Types].xml` +
  entry-integrity (AC1 proxy); `Presentation(embedded)` re-open; **idempotent**
  (embed twice → one entry = second schema); minified/size via the returned
  result (AC4); `read_embedded_schema` None-on-absent / None+warn-on-malformed /
  raise-on-non-zip; **non-ASCII round-trip** (NIT-1). Use a small **synthetic
  deck** for zip-mechanics tests (NIT-4) + one bundled smoke test.
- `docs/user-stories/GAP-ANALYSIS.md`: US-1.5 → ✅ Met (Met 6→7, Not met 6→5).
- `docs/user-stories/chenyu-user-stories.md`: US-1.5 ACs → `[x]`.

## Acceptance Criteria (US-1.5) — to deliver

- [x] After embedding, the PPTX opens in PowerPoint without errors or repair
      prompts (proxy-verified: `python-pptx` re-opens; `[Content_Types].xml`
      first + valid + declares `json`; all other entries decompressed-content-
      identical; "no repair prompt" itself is manual).
- [x] The JSON is retrievable by re-reading the zip at the known path
      (`read_embedded_schema` returns the dict; absent → `None`).
- [x] Existing slide content, layouts, and media are untouched (all non-
      `[Content_Types].xml`/non-schema entries decompressed-content-identical;
      only `ppt/template_schema.json` is new; `[Content_Types].xml` gains only
      the `json` Default).
- [x] File size increase is logged to the user (`embed_schema` returns
      `EmbeddedSchemaResult`; CLI logs original → new → delta).

## Implementation Phases

### Phase 1: Embedding core (schema_extractor.py)

- [x] Task 1: `import zipfile`, `import os`, `import tempfile`; constant
      `_EMBEDDED_SCHEMA_PATH = "ppt/template_schema.json"`; `EmbeddedSchemaResult`
      dataclass/NamedTuple (`output_path`, `original_bytes`, `new_bytes`,
      `delta_bytes`).
- [x] Task 2: `_inject_json_default(xml_bytes) -> bytes` — parse with
      `lxml.etree`; if no `<Default Extension="json">` exists, insert it (before
      any `<Override>`); re-serialize. Idempotent.
- [x] Task 3: `embed_schema(pptx_path, schema, output_pptx_path) ->
      EmbeddedSchemaResult` — read original; write a temp zip:
      `[Content_Types].xml` (= `_inject_json_default` of the original) **first**;
      then every other original entry in original order, **decompressed-content-
      identical** via `writestr` with copied `ZipInfo` (skip any entry named
      `_EMBEDDED_SCHEMA_PATH` — **idempotent**, MAJOR-3); then
      `ppt/template_schema.json` = minified schema; `os.replace(temp,
      output_pptx_path)` (atomic, MINOR-6). Compute sizes from the result.
- [x] Task 4: `read_embedded_schema(pptx_path) -> dict | None` — open zip; if
      the path is absent → `None`; present → `json.loads` (catch
      `json.JSONDecodeError` → `logger.warning` + `None`, MINOR-3; non-dict →
      `None`); **corrupt/non-zip → raise `TemplateExtractionError`** (MINOR-4).

### Phase 2: CLI integration (schema_extractor.py)

- [x] Task 5: Add `--embed` (store_true) + `--output-pptx <path>` (default
      `<input_stem>.templated.pptx`).
- [x] Task 6: If `--output-pptx` is set without `--embed` → argparse error
      (exit 2, MINOR-2).
- [x] Task 7: When `--embed`, call `embed_schema` after writing `--output` JSON;
      log `result.original_bytes → result.new_bytes (+result.delta_bytes)`
      (AC4, from the returned struct — MINOR-5); log the output path.

### Phase 3: Tests (test_schema_extractor.py)

- [x] Task 8: **Round-trip** on a small synthetic deck — `extract_schema` →
      `embed_schema` → `read_embedded_schema` → deep-equal to the extracted dict.
- [x] Task 9: **`[Content_Types].xml` + entry integrity (AC1 proxy)** — still the
      first entry; valid XML; **contains the `json` Default**; preserves the
      original Defaults + Overrides; every **other** original entry
      decompressed-content-identical (hash of `ZipFile.read(name)`); only
      `ppt/template_schema.json` is otherwise new.
- [x] Task 10: `Presentation(embedded_pptx)` re-opens; same slide + layout count.
- [x] Task 11: **Idempotent (MAJOR-3)** — embed the embedded PPTX again → exactly
      one `ppt/template_schema.json` entry, content = the second schema.
- [x] Task 12: **Minified/size (AC4, MINOR-5)** — embedded JSON compact (no
      `", "`/`": "`); `result.delta_bytes > 0`; assert on the returned
      `EmbeddedSchemaResult`, not log text.
- [x] Task 13: **`read_embedded_schema` error contract (MINOR-3/4)** — plain
      PPTX → `None`; malformed JSON at the path → `None` (+ warning); non-zip
      input → raises `TemplateExtractionError`.
- [x] Task 14: **Non-ASCII round-trip (NIT-1)** — inject a non-ASCII title
      (e.g. `模板测试`); survive embed → read byte-for-byte.
- [x] Task 15: One bundled-`template.pptx` smoke test (realism).

### Phase 4: Docs

- [x] Task 16: Update `GAP-ANALYSIS.md` US-1.5 → ✅ Met (Met 6→7, Not met 6→5);
      `chenyu-user-stories.md` US-1.5 ACs → `[x]`. → Epic 1 complete (all 5
      stories Met).

## Test matrix

| Case | Expected |
| --- | --- |
| extract → embed → `read_embedded_schema` | deep-equal to the extracted dict |
| `[Content_Types].xml` after embed | still first; valid; **declares `json` Default**; preserves original Defaults + Overrides |
| other zip entries after embed | decompressed-content-identical (hash of `read()`); only `ppt/template_schema.json` is otherwise new |
| `Presentation(embedded_pptx)` | re-opens without error; same slide/layout count |
| embed the embedded PPTX again (idempotent) | exactly one `ppt/template_schema.json` = the second schema |
| `embed_schema` return value | `delta_bytes > 0`; JSON minified (no insignificant whitespace) |
| `read_embedded_schema`: plain PPTX | `None` |
| `read_embedded_schema`: malformed JSON | `None` (+ warning) |
| `read_embedded_schema`: non-zip input | raises `TemplateExtractionError` |
| non-ASCII title round-trip | survives embed → read unchanged |

## Verification

```bash
# from .opencode/skills/ppt-template-filler/scripts
python -m pytest tests/test_schema_extractor.py -q
python -m pytest tests/ -q
# (on Windows use %TEMP%\s.json instead of /tmp/s.json)
python schema_extractor.py --input templates/template.pptx --output ./s.json --embed
```

## Out of Scope (deferred) / OPEN QUESTIONS

- **Renderer migration to embedded schema** — `template_introspector.py` still
  reads the sidecar; this issue is producer-only. A future issue makes the
  renderer prefer `read_embedded_schema` and deprecate the sidecar.
- **`schema_version` read-gate (MINOR-7)** — `read_embedded_schema` returns the
  dict as-is; rejecting/migrating an incompatible embed is deferred to the
  renderer-migration issue.
- **`--output` relaxed to optional for embed-only use (MINOR-1)** — deferred to
  US-3.x ("downloadable templated PPTX").
- **Real PowerPoint "no repair prompt"** — not auto-testable; verified by proxy
  (Decision 9) and documented as manual.
- **Embedding other parts** (thumbnails, preview images, resolved assets) — out
  of scope; only `ppt/template_schema.json` here.
- **Compression-level tuning** for the new entry — uses the zip default.
- **Relationship part / `ppt/_rels` registration** for the new part — PowerPoint
  does not require it for an unknown part; not added.

## Risks

- **Zip rewrite correctness** — main risk; mitigated by the
  `[Content_Types].xml`-first + per-`read()`-hash integrity test (Task 9) and
  `ZipInfo` preservation.
- **`json`-Default injection (MAJOR-2)** — modifying `[Content_Types].xml` is
  deliberate (strict-safe); mitigated by the "first + valid + json Default +
  preserves originals" assertion (Task 9). Deviates from the requirement Details'
  "do not modify `[Content_Types].xml`" because AC1 is binding and the template
  declares no `json` Default.
- **Idempotency (MAJOR-3)** — re-embed no longer duplicates (skip-existing);
  covered by Task 11.
- **Atomic write (MINOR-6)** — temp + `os.replace` prevents partial/corrupt
  output.
- **Real PowerPoint "no repair prompt"** — not auto-testable; proxy + manual.
- **Backward compatibility — low** — `--embed` opt-in; default behavior and the
  renderer (sidecar) unchanged.

## Code-review follow-ups

Post-implementation code review (verdict: Approve with revisions — 0 Critical/0
Major). All locked architecture decisions (MAJOR-1/2/3, MINOR-2/3/4/5/6) verified
correctly implemented. The applied fixes (behavior-preserving + additive tests):

- [x] **MINOR-1 — log sign rendering.** Both size-delta log lines
  (`schema_extractor.py` `embed_schema` + CLI) changed `"+%d"` → `"%+d"` so a
  negative delta (the bundled template produces one due to re-compression)
  renders `-N` instead of `"+-N"`.
- [x] **MINOR-2 — `read_embedded_schema` contract test gaps closed.** The
  malformed-JSON test now asserts the warning fires; added
  `test_non_dict_json_returns_none` (array at the path → None + warning) and
  `test_missing_file_raises_domain_error` (FileNotFoundError →
  `TemplateExtractionError`). All four contract branches now covered.
- Not applied (deferred): **MINOR-3** (widen/document `embed_schema`'s exception
  surface — unreachable via normal CLI flow since `extract_schema` opens first),
  **MINOR-4** (zip-metadata fidelity is intentionally not preserved per MAJOR-1),
  and the NITs (cosmetic).

Tests: `test_schema_extractor.py` → **95 passed** (was 93). No production-logic
change — only log formatting (MINOR-1) and additive tests (MINOR-2).

## References

- Requirements: `docs/user-stories/chenyu-user-stories.md` → Epic 1, US-1.5.
- Gap analysis: `docs/user-stories/GAP-ANALYSIS.md` → §2 US-1.5 (❌ Not met), §4
  P0.
- GitHub issue: #55 (`[US-1.5] JSON Storage Inside PPTX Zip`).
- Predecessors: US-1.1 (issue #48, PR #49), US-1.2 (issue #50, PR #51), US-1.3
  (issue #52, PR #53), US-1.4 (issue #54).
- Format template: `PLANS/PLAN-GIT-54.md`.
- Architecture review: findings MAJOR-1 (decompressed-content-identity),
  MAJOR-2 (inject `json` Default), MAJOR-3 (idempotent), MINOR-2/3/4/5/6
  incorporated above; MINOR-1/7 + NIT-1/4 noted in §"Out of Scope"/Tasks; NIT-2
  (this line) + NIT-3 (`/tmp` note) + NIT-5 (status → v2) applied.
- After implementation: update `GAP-ANALYSIS.md` US-1.5 → ✅ Met (Met 6→7, Not met
  6→5); `chenyu-user-stories.md` US-1.5 ACs → `[x]`. → Epic 1 complete (all 5
  stories Met).
