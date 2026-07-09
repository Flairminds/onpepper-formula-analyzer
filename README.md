# Excel Formula Analyzer

Compares Excel formulas between two versions of the same workbook — cell by cell —
and classifies every change as **added**, **removed**, **modified** (further split
into *semantic change* vs *reference shift*), or **unchanged**. Built for
borrowing-base / compliance workbooks where the same template is re-filled every
period and a human needs to know which formula changes are genuine logic edits
versus artifacts of columns/rows shifting position.

A web UI (`server.py` + `formula-viewer 2.html`) lets you upload two `.xlsx` files,
runs the comparison, and renders an interactive, filterable diff. A CLI
(`compare_versions.py`) and a Postgres-backed ingestion script (`tracker.py`) are
also available.

## Stack & versions

Developed and tested against:

| Component | Version |
|---|---|
| Python | 3.11 |
| Flask | 3.1.3 |
| openpyxl | 3.1.5 |
| psycopg2-binary | 2.9.11 |
| python-dotenv | 1.2.2 |

No other runtime dependencies. PostgreSQL is only required for `tracker.py`
(single-workbook ingestion) — the comparison web UI and CLI need no database.

## Project structure

```
.
├── README.md
├── requirements.txt          # openpyxl, psycopg2-binary, python-dotenv, flask
├── .gitignore
├── init_schema.sql           # Postgres schema (only needed for tracker.py)
├── server.py                 # Flask backend for the web UI
├── compare_versions.py       # CLI: diff two workbooks, print + optionally save JSON
├── tracker.py                # CLI: ingest ONE workbook's formulas into Postgres
├── formula-viewer 2.html     # Web UI — upload workbooks, browse the diff (current)
├── formula-viewer.html       # Earlier standalone viewer (open-a-JSON-file only)
├── compare_versions.py tests:
│   ├── test_compare_formulas.py
│   ├── test_dependency_extractor.py
│   ├── test_formula_normalizer.py
│   └── test_mapping.py
├── reports/                  # Versioned comparison_report_{N}.json (gitignored)
├── config/
│   └── .env                  # DB credentials for tracker.py (gitignored, not committed)
└── src/
    ├── reader.py              # Load .xlsx with openpyxl, list sheets/metadata
    ├── mapping.py              # Column-letter → business-header-name mapping
    ├── formula_extractor.py    # Pull every formula string out of a workbook
    ├── formula_normalizer.py   # Replace cell refs with [BusinessName] placeholders
    ├── formula_comparator.py   # The core diff engine (added/removed/modified + reasons)
    ├── ast_builder.py          # Recursive-descent parser: formula string → AST
    ├── semantic_comparator.py  # AST-based "business meaning" diff (fine-grained change_type)
    ├── comparison_validator.py # Sanity-checks the semantic_comparator's own output
    ├── dependency_extractor.py # Cross-cell dependency graph (Phase 3, used by tracker.py)
    ├── compare_service.py      # Orchestrates the pipeline; shared by CLI + Flask
    └── persist.py              # Postgres insert/read helpers (tracker.py only)
```

## Setup

```bash
git clone <repo-url>
cd onpepper-formula-analyzer
python3 -m venv venv && source venv/bin/activate   # optional but recommended
pip3 install -r requirements.txt
```

Requires **Python 3.11+**. On macOS, use `python3`/`pip3` — the `python` command
generally isn't aliased.

No `.env` or database setup is needed to run the comparison UI or CLI. Only
`tracker.py` (single-workbook Postgres ingestion) needs the `config/.env` file
described in [Database setup](#database-setup-tracker-only) below.

## Running the web UI (recommended)

```bash
python3 server.py
```

Starts a Flask dev server on **http://127.0.0.1:5000/**. Open that URL in a
browser. From there you can:

- Pick two `.xlsx` files and click **Compare & Save Report** — runs the full
  pipeline and renders the result immediately.
- Use the **History** dropdown (top-right) to reload any past comparison.
- Browse by **sheet**, **column**, **change type**, and **reason** (see
  [Reason attribution](#reason-attribution) below) — the Cell view lists every
  changed formula with old/new formula text and normalized (business-name) form.

Stop the server with `Ctrl+C`, or `pkill -f "python3 server.py"` if it's running
in the background. Logs go to stdout; nothing is logged to a file by default.

If port 5000 is already in use (macOS's AirPlay Receiver commonly squats on it),
either disable AirPlay Receiver in System Settings → General → AirDrop & Handoff,
or edit the `app.run(port=5000)` call at the bottom of `server.py`.

### Web UI endpoints (for scripting / integration)

| Method & path | Purpose |
|---|---|
| `GET /` | Serves the viewer HTML |
| `POST /api/compare` | Multipart form fields `old_file`, `new_file` (`.xlsx`) → runs the comparison, saves `reports/comparison_report_{N}.json` (auto-incrementing), returns `{version, file, report}` |
| `GET /api/reports` | Lists saved reports: `[{version, file, generated_at, old_file, new_file}, ...]`, newest first |
| `GET /api/reports/<version>` | Returns one saved report's full JSON |

## Running the CLI

```bash
python3 compare_versions.py old.xlsx new.xlsx
python3 compare_versions.py old.xlsx new.xlsx --save report.json
python3 compare_versions.py old.xlsx new.xlsx --save report.json --max-diff 0   # show all, not just 50/sheet
```

Prints a sheet-by-sheet summary and per-cell diff to stdout (each ADDED/REMOVED/
MODIFIED line includes its `reason` in brackets), and optionally writes the same
structured report the web UI produces.

## Database setup (`tracker.py` only)

`tracker.py` ingests a *single* workbook's formulas + metadata into Postgres —
it does not compare two files. It is unrelated to the comparison pipeline above.

1. Create a Postgres database and run the schema:
   ```bash
   psql -d your_db -f init_schema.sql
   ```
2. Create `config/.env` (not committed — see `.gitignore`) with:
   ```
   DB_HOST=localhost
   DB_PORT=5432
   DB_NAME=your_db
   DB_USER=your_user
   DB_PASSWORD=your_password
   ```
3. Run:
   ```bash
   python3 tracker.py path/to/workbook.xlsx
   ```
   Each run inserts one row into `workbook_records` (a single JSONB column
   holding the column mapping, formulas, dependency graph, normalized formulas,
   and formula ASTs for that workbook).

## Architecture: how a comparison is built

`src/compare_service.py::run_comparison(old_path, new_path)` is the single entry
point both the CLI and the Flask backend call. Pipeline, per workbook:

1. **`reader.load_excel_workbook`** — load with `data_only=False` so formula
   strings (not cached values) are retained.
2. **`formula_extractor.extract_formulas`** — `{sheet: {cell_address: formula}}`
   for every cell whose value starts with `=`.
3. **`mapping.get_business_column_mapping`** — reads row 9 (configurable via
   `header_row_index`) as the header row, `{sheet: {col_letter: header_name}}`.
   Duplicate or blank header text is disambiguated with a **1-based occurrence
   count** — `Excess (1)`, `Excess (2)`, `Unnamed (1)`, … — not the column's own
   letter (see [Why not the column letter](#why-occurrence-count-not-column-letter) below).
4. **`formula_normalizer.normalize_all_formulas`** — rewrites every formula's
   cell references as `[BusinessName]` placeholders (row numbers dropped), so
   `B10` and `B99` in the same column both read as `[Revenue]`. This is what
   lets the tool tell "the row shifted" apart from "the logic changed".
5. **`formula_comparator.compare_workbooks`** → **`compare_sheet`** — the core
   diff, described below.
6. **`semantic_comparator.batch_compare`** — a second, independent, AST-based
   comparison producing a fine-grained `change_type`
   (`operator_changed`, `function_changed`, `argument_count_changed`,
   `column_changed`, `structural_change`, `constant_changed`, `sheet_changed`)
   and a plain-English `description` per cell. Stored in the report under
   `semantic_diff` / `semantic_summary`; **not currently surfaced in the viewer
   UI** (the UI's Semantic/Ref Shift badge comes from step 5, not this step).

### `compare_sheet`: added / removed / modified / unchanged

Cells are compared by **identical address** (`old["B10"]` vs `new["B10"]`), not
by row identity — this tool has no concept of "this loan moved from row 50 to
row 84because 34 rows were inserted above it"; it only sees that address B10's
formula text is or isn't the same in both files. For a MODIFIED cell,
`change_type` is:

- `reference_shift` — normalized forms match (row/address moved, business logic
  unchanged)
- `semantic_change` — normalized forms differ (logic genuinely changed)

**Column-shift alignment**: if a column's header name moved to a different
letter between versions (because some other column was inserted/removed
earlier in the row), comparing old-letter-X vs new-letter-X compares two
*different* business columns and would falsely look like a semantic change. So
before computing `change_type`, the comparator re-aligns: it compares the
column's *previous* letter in the old file against its *current* letter in the
new file. The report exposes the raw alignment (`aligned_old_cell`,
`aligned_old_formula`, `aligned_old_normalized`) so this claim is independently
verifiable in the UI's expand panel, not just asserted.

### Reason attribution

Every ADDED, REMOVED, and MODIFIED cell also gets a `reason` — a best-effort
structural explanation, computed in `formula_comparator._classify_columns` /
`_row_added` / `_row_removed`:

| reason type | Meaning |
|---|---|
| `row_added` | This row's formula density jumped from near-empty to near-full (a new record, e.g. a new loan) — calibrated per sheet against that sheet's own 75th-percentile "loaded row" size, since a Prices row (1 formula) and a Loan List row (150+ formulas) mean different things by "loaded" |
| `row_removed` | Mirror of the above — a row collapsed from full to near-empty |
| `column_added` | This column's header name is new — didn't exist in the old file at all |
| `column_removed` | This column's header name no longer exists in the new file |
| `column_shifted` | Header name exists in both files at a different letter; the aligned comparison confirms the logic is unchanged |
| `column_shifted_and_logic_changed` | Header name shifted letters **and** the aligned comparison still shows a real logic difference |
| `null` / "no reason" | No structural explanation — a targeted, standalone change (or genuinely ambiguous, see caveat below) |

**Caveat — duplicate column names**: many spreadsheets repeat the same header
text in a recurring block (e.g. "Excess" and "Net Loan Balance" once per
concentration test, ~70+ times in one real workbook). These are disambiguated
as `Excess (1)`, `Excess (2)`, … by left-to-right occurrence order *within each
file independently*. Shift/reason detection for these columns is therefore only
as correct as the assumption that occurrence *N* in the old file is the same
conceptual column as occurrence *N* in the new file (true as long as no repeat
block itself was inserted/removed/reordered — true for the workbooks this was
built against, but worth checking if a new sheet's structure differs).

#### Why occurrence-count, not column letter

An earlier version of this disambiguation appended the column's own letter
(`Excess (GB)`), which broke every single repeated column: since the whole
point of comparing by header *name* is to survive the column *letter* shifting
between versions, tying the disambiguated name to the letter guarantees
`Excess (GB)` in the old file never equals `Excess (GC)` in the new file even
when it's the exact same conceptual column — every repeated column then looks
"added" in the new file and "removed" in the old one. The occurrence count is
stable across a pure letter-shift, so it doesn't have this problem.

### Known limitations

- **No row realignment.** If actual data rows are inserted/deleted/reordered
  (e.g. loans added or paid off, causing everything below to shift down a row),
  formulas that are simply "copy the row above, adjusted for this row" look
  identical regardless of which loan occupies the row — so this kind of change
  is invisible to a formula-only diff. Row `added`/`removed` reasons detect
  *whole rows* going from sparse-skeleton to fully-loaded (or vice versa), not
  arbitrary row reordering within an already-loaded range.
- **`semantic_comparator`'s fine-grained change_type is not wired into the
  UI.** It's computed and saved in every report (`semantic_diff` /
  `semantic_summary`) but the viewer currently only reads the coarser
  `change_type` from `formula_comparator`. Wiring the AST-based description
  into the expand panel is a reasonable next step if you want per-cell
  descriptions like "Function changed: SUM → AVERAGE" instead of just a
  Semantic/Ref Shift badge.
- **`ast_builder.py`'s parser depends on exact openpyxl `Tokenizer` type
  strings** (`OPERATOR-INFIX`, `OPERATOR-PREFIX`, `OPERATOR-POSTFIX`, `PAREN`).
  These were previously mismatched (fixed in this codebase already) and could
  silently break again if openpyxl changes its tokenizer's type vocabulary in a
  future version — there's no test asserting the parser round-trips a formula
  with every operator kind, which would catch that regression early.
- **`header_row_index` is hardcoded to `9`** as a default across `reader.py`
  and `mapping.py`. If you point this at a workbook with headers on a different
  row, pass `header_row_index=` explicitly wherever these functions are called
  — there's no auto-detection.
- **The test scripts** (`test_compare_formulas.py`, `test_dependency_extractor.py`,
  `test_formula_normalizer.py`, `test_mapping.py`) are standalone scripts with
  hardcoded Windows paths (`E:\Excel_Formula_Analyser\...`), not a real pytest
  suite — `pytest` collection fails on `test_dependency_extractor.py` for this
  reason. Run them individually with `python3 <file> <path-to-workbook.xlsx>`,
  or treat `src/compare_service.run_comparison()` run against your real
  workbooks (as done throughout development) as the actual regression check.

## Generated / gitignored files

`.gitignore` excludes `__pycache__/`, `.env`, `.DS_Store`, and the generated
analysis outputs (`dependency_test_output.json`, `formula_comparison.json`,
`full_report.json`, `normalized_formulas_preview.json`, `reports/`) — all of
these are regenerable by running `tracker.py` / `compare_versions.py` / the web
UI, so none of them belong in version control. The two `.xlsx` sample workbooks
in the repo root **are** tracked, intentionally, as reference/test data.

`reports/comparison_report_{N}.json` accumulates one file per comparison run
(via the web UI or `compare_versions.py --save`) — periodically clear out
`reports/` if it grows large; each file can be several MB to tens of MB for a
large workbook.
