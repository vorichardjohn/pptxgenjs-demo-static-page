# Table Builder Workflow Opportunities Backlog

This backlog translates the workflow analysis into concrete implementation steps for the current static GitHub Pages demo.

## Goals

- Reduce repeated setup work for recurring exports.
- Improve large-dataset reliability for table generation.
- Add a practical data-prep layer before export.
- Keep the UI scannable and avoid horizontal-scroll-first interactions.

## Prioritized Roadmap

## Phase 1 (quick wins, high impact)

### 1) Save and Load Presets

**Problem**
Users currently re-enter styling and export settings every session.

**Scope**
- Save current options (table options + column config) to LocalStorage under a named preset.
- Load preset from LocalStorage into UI controls.
- Export/import presets as JSON file.

**Acceptance criteria**
- User can click **Save preset**, provide a name, and see it appear in a preset dropdown.
- User can click **Load preset** and all related controls update immediately.
- Imported preset JSON validates shape and fails with clear status text if invalid.
- Preview updates after loading preset.

**Files likely to change**
- `table-builder.html` (preset UI controls)
- `table-builder.js` (serialize/deserialize + LocalStorage + import/export)
- `styles.css` (compact layout for preset controls)

---

### 2) Multi-slide Pagination for Large Tables

**Problem**
Single-slide rendering is likely to overflow with real-world datasets.

**Scope**
- Add `rowsPerSlide` input.
- Split body rows into chunks; render one table per slide with repeated header.
- Include “Slide X of Y” label in title or footer.

**Acceptance criteria**
- Given 100 rows and rowsPerSlide=20, export creates 5 slides.
- Header row appears on each slide.
- Column widths, inclusion, alignment, and style options apply consistently on every slide.

**Files likely to change**
- `table-builder.html` (rows per slide option)
- `table-builder.js` (chunking logic + looped slide generation)

---

### 3) Bulk Column Operations

**Problem**
Include/exclude is currently per-column only.

**Scope**
- Add **Select all** / **Deselect all** buttons.
- Add quick “Apply width/alignment to included columns” action.

**Acceptance criteria**
- One click toggles include state for all columns.
- Bulk width/alignment updates all currently included columns.
- Preview rerenders after each bulk action.

**Files likely to change**
- `table-builder.html` (bulk action controls)
- `table-builder.js` (batch operations)
- `styles.css` (button group spacing)

## Phase 2 (data quality)

### 4) Column Data Types and Formatting

**Problem**
Values are exported as plain strings, limiting professional report quality.

**Scope**
- Per-column type selector: `text | number | currency | percent | date`.
- Per-type formatting inputs (e.g. decimals, date format).
- Optional “negative values in red” style toggle.

**Acceptance criteria**
- Typed columns render consistently in preview and PPTX output.
- Invalid numbers/dates are handled gracefully with fallback text.
- Formatting options are persisted in presets.

**Files likely to change**
- `table-builder.html` (type/format controls)
- `table-builder.js` (formatters + typed preview rendering)

---

### 5) Data Prep Step

**Problem**
No pre-export filtering/sorting/renaming workflow exists.

**Scope**
- Add optional transform panel before table options:
  - sort by column + direction
  - filter by simple conditions
  - rename display headers
  - optional derived column expression (safe subset)

**Acceptance criteria**
- Transform operations only affect derived working dataset, not original upload payload.
- Preview reflects transforms immediately.
- Export uses transformed dataset.

**Files likely to change**
- `table-builder.html` (transform controls)
- `table-builder.js` (working dataset pipeline)

## Phase 3 (operational polish)

### 6) Export Strategy Modes

**Problem**
Large outputs may be slow/heavy and hard to validate quickly.

**Scope**
- Modes: `preview-only` (first N slides), `single deck`, `split decks`.
- Split decks by max slides-per-file.

**Acceptance criteria**
- Preview mode produces only configured initial slides.
- Split mode writes sequentially named files (`name-part-1.pptx`, etc.).
- Status text reports progress across slide creation and file writing.

**Files likely to change**
- `table-builder.html` (export mode controls)
- `table-builder.js` (mode orchestration + progress reporting)

---

### 7) Theme Packs (Brand Presets)

**Problem**
Manual color/font configuration is repetitive for teams.

**Scope**
- Add curated theme packs (e.g., Corporate Blue, Minimal Gray, Dark Contrast).
- Applying a theme updates default colors/font settings, with user override retained.

**Acceptance criteria**
- Theme selection updates relevant controls in one action.
- Manual edits after theme apply are preserved until next explicit theme apply.

**Files likely to change**
- `table-builder.html` (theme selector)
- `table-builder.js` (theme map + apply behavior)

## Implementation Notes

- Preserve no-horizontal-scroll-first principle in controls layout.
- Keep keyboard accessibility for all interactive controls.
- Ensure all state-changing controls trigger preview refresh.
- Validate numeric/color inputs defensively before export.

## Suggested Order of Execution

1. Presets
2. Pagination
3. Bulk include/exclude and bulk formatting
4. Typed formatting
5. Data transforms
6. Export strategies
7. Themes

This order minimizes regressions by stabilizing state persistence and slide generation first, then layering richer data and UI features.
