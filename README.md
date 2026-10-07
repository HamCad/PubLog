# PubLog Dashboard

PowerShell 5.1 + SQLite (Windows' built-in `winsqlite3.dll`) warehouse and
single-file HTML dashboard for FLIS sustainment data. Zero external
dependencies: nothing to install, no modules, no vendor executables.

Two independent pipelines, in separate files:

| Pipeline | Entry point | Input | Output |
|---|---|---|---|
| **FLIS** (drives the dashboard) | `Invoke-FlisPipeline.ps1` | the raw FLIS bulk-download `.txt` files, dropped in one folder | `sustainment.db` |
| **PUB LOG / FED LOG** (reference data) | `Update-PubLogDatabase.ps1` | `PublogDVD.zip` or its extracted `.TAB` files | `publog.db` (+ `publog_alt.db` index) |

Plus a local tool on top of PUB LOG: the **Alternate Part Finder**
(`Start-AltPartFinder.ps1`), a drill-down page for finding candidate
alternates by locking some characteristics and widening others
([docs/ALTERNATE_PARTS.md](docs/ALTERNATE_PARTS.md)).

## Monthly use

```powershell
# 1. Drop every file from the FLIS bulk download into one folder, then:
.\Build-SustainmentDashboard.ps1 -DropFolder D:\FLIS\drop      # pipeline + dashboard
#    or just the database:
.\Invoke-FlisPipeline.ps1 -DropFolder D:\FLIS\drop

# 2. Separately, when a new PUB LOG cut arrives (also rebuilds the finder's index):
.\Update-PubLogDatabase.ps1 -Source D:\Downloads\PublogDVD.zip

# 3. Any time: the Alternate Part Finder (opens http://127.0.0.1:8765/, Ctrl+C to stop)
.\Start-AltPartFinder.ps1
```

`Invoke-FlisPipeline.ps1` identifies every file by name, checks its header
against the source of truth, lands it verbatim (`fb_*` tables, stamped with
the batch number from `Batch{N}Details.txt`), converts it into the typed
sustainment tables (ISO dates, numbers, normalized contract/part numbers),
applies each table's update rule, and runs coverage checks. Anything wrong
is reported: grouped and capped on the console, every occurrence in
`fb_issue_log`, and a text report per run in `logs\`. A run with errors
ends by throwing, after reporting everything; the dashboard build refuses to
render over a failed run unless `-ContinueOnPipelineErrors`.

Useful switches: `-AsOfMonth yyyy-MM` (current month for the rolling stock and
forecast files when file timestamps were reset by copying),
`-UpdatePriorMonths N` (default 1), `-StoreFutureStockMonths`,
`-ConsoleDetail Summary|Files|All`, `-RequireAllFiles`.

## Update rules (flis_model.psd1)

| Rule | Tables | Behavior |
|---|---|---|
| ItemMaster | `item_niin` (SegmentA) | upsert; `first_seen_batch` / `last_seen_batch` / `in_latest_batch`; never deleted |
| Upsert | `vendor_cage` | upsert on CAGE; never deleted; unknown CAGEs referenced elsewhere get `is_stub = 1` rows |
| Snapshot | part numbers, management, MOE, characteristics, freight, packaging, weapon systems, ... | rows for NIINs in the batch are replaced; NIINs that dropped out keep their last rows |
| History | `contract_award_line` (ProcurementHistory + Archive), `solicitation_line` | append-only, deduplicated; re-delivered rows get `last_seen_batch`; changed values are applied and the old value logged in `history_value_change` |
| Rolling | `stock_on_hand_monthly`, `dla_forecast_monthly` | month columns unpivoted; current month (+N prior) take the latest pull, older months only filled if missing, blanks never overwrite, future stock months skipped |

## Files

| File | Purpose |
|---|---|
| `Invoke-FlisPipeline.ps1` | FLIS entry point (schema migrations -> raw landing -> typed tables -> coverage) |
| `FlisIngest.ps1` | raw landing: file accounting, header checks, strict CSV parsing, `fb_*` tables, issue reporting |
| `FlisTransform.ps1` | migrations runner, typed conversion engine, update rules, coverage checks |
| `flis_registry.psd1` | expected headers of the 33 FLIS files (generated from `docs/flis/`) |
| `flis_model.psd1` | raw -> typed column mapping, types and update rule per table (generated) |
| `migrations/NNN_*.sql` | schema, applied in order automatically (`schema_migrations`) |
| `Build-SustainmentDashboard.ps1` | optional FLIS pipeline run + dashboard HTML |
| `Update-PubLogDatabase.ps1` | PUB LOG entry point |
| `PubLogReader.ps1`, `PubLogEtl.ps1` | `.TAB` decoder and PUB LOG ETL (`Update-SustainmentFromPubLog` retired) |
| `Test-PubLogDecoder.ps1` | verifies the `.TAB` decoder on a cut |
| `Start-AltPartFinder.ps1`, `PubLogAlternates.ps1`, `altparts.html` | Alternate Part Finder: characteristic parser, `publog_alt.db` index, query engine, localhost page |
| `SQLiteInterop.ps1` | winsqlite3.dll driver |
| `tests/` | synthetic-data generator and self-tests (`Test-FlisIngest.ps1`, `Test-FlisPipeline.ps1`, `Test-AltPartFinder.ps1`, `Test-ScriptSyntax.ps1`); no real data |
| `tools/generate_flis_model.py` | developer tool: regenerates `flis_model.psd1` + `migrations/001` from `docs/flis/` |
| `docs/flis/` | owner's FLIS source-of-truth headers and the schema-alignment handoff |
| `docs/PUBLOG_*.md` | PUB LOG `.TAB` format, schema, and ideas for using more of it (`PUBLOG_LEVERAGE.md`) |
| `docs/ALTERNATE_PARTS.md` | Alternate Part Finder: how characteristics are parsed and matched, consistency findings |
| `legacy/` | the original guessed schema (reference only; not used) |
