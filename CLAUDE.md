# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository layout

The actual R project lives in `labour-market-stats-brief/`. All R code, `Dockerfile`, and the project `README.md` are inside that subdirectory — commands below assume it as the working directory. The top-level `curly-waddle/` folder also holds QA artifacts (`A01 is ok.txt`, `LM Stats Briefing QA.xlsx`) and a source zip; treat them as non-code assets.

## Running the app

This is a Shiny (R) application, not a typical build/test project — there is no package manager lockfile, no test suite, and no linter wired up.

```r
# From R, with working directory = labour-market-stats-brief/
shiny::runApp(".")

# Or (what the Dockerfile runs):
R -e "shiny::runApp('/app', host = '0.0.0.0', port = 8888)"
```

Docker image is based on an internal ECR base (`jupyterhub-visualisation-base:rv4`). Required R packages are listed in `labour-market-stats-brief/Dockerfile` — add any new dependency there, there is no `DESCRIPTION` / `renv.lock`.

Generating outputs **without the Shiny UI** (useful for debugging a single run):

```r
# Working directory must be the project root (labour-market-stats-brief/)
setwd("/path/to/labour-market-stats-brief")
source("utils/word_output.R")
generate_word_output()   # writes utils/DBoutput.docx (gitignored)
```

Note: `utils/word_output.R` is loaded via `source(..., local = FALSE)` from `app.R`, and it sources `utils/calculations.R`, which in turn sources every file in `sheets/`. All relative paths are resolved against the project root, so **never `cd` into `utils/`** before sourcing.

## Architecture

### Two data paths, one renderer

The app has two mutually-exclusive input modes selected via tabs:

- **Automatic tab** — connects to Postgres (`DBI` + `RPostgres`, DSN from `DATABASE_DSN__datasets_1`) and reads from `"ons"."*"` and `"oecd"."*"` schemas. Each `sheets/<metric>.R` file exposes a `fetch_<metric>()` that queries one table plus a `calculate_<metric>(manual_mm, ...)` that returns a list of current value + deltas. Entry point: `utils/calculations.R` sources them all, then `utils/word_output.R` / `sheets/excel_audit_workbook.R` render the Word/Excel outputs.
- **Manual tab** — user uploads ONS Excel files (A01, HR1, X09, RTISA, optional CLA01/X02/OECD). `utils/calculations_from_excel.R` parses those files and writes the **same variable names** into the target environment that `utils/calculations.R` writes globally. `utils/manual_word_output.R` then fills the Word template from those vars.

The renderers (`utils/word_output.R`, `utils/manual_word_output.R`, `sheets/excel_audit_workbook.R`, narrative generators `sheets/summary.R` and `sheets/top_ten_stats.R`) are shared/parallel across both paths. When changing a metric, **keep the DB path and the Excel path in sync** — any new variable produced by `utils/calculations.R` must also be produced by `utils/calculations_from_excel.R`, or the manual mode Word output silently shows em-dashes.

### The `manual_month` contract

`manual_month` is a lowercase 7-char string like `"dec2025"` and is the pivot for every date-based lookup. Resolution order:

1. Auto-detected from Postgres (`auto_detect_manual_month()` in `utils/helpers.R`) — takes the latest `time_period` in `ons.labour_market__age_group` and adds 2 months (release anchor).
2. In manual mode, detected from A01 sheet `"1"` (`.detect_manual_month_from_a01` in `utils/calculations_from_excel.R`).
3. User override typed into the reference-month input on either tab.
4. Fallback to `Sys.Date()` if all of the above fail (`utils/config.R`).

`parse_manual_month()` converts it to the 1st of the month; `anchor_m = manual_month %m-% 2 months` is the LFS quarter-end used everywhere. Period labels are produced by `make_lfs_label()` (e.g. `"Oct-Dec 2025"`) and `lfs_label_narrative()` (e.g. `"October 2025 to December 2025"`). Baseline dates for "change since" columns live in `utils/config.R` (`COVID_DATE`, `ELEC24_DATE`, `COVID_LFS_LABEL`, `COVID_VAC_LABEL`, `ELECTION_LABEL`).

### "Latest" vs "aligned" modes

`vacancies_mode` and `payroll_mode` (both `"latest"` or `"aligned"`) let the user choose between the newest available period and the period matching the LFS quarter end. They are read by `calculate_vacancies()` / `calculate_payroll()` and can be overridden via `vacancies_mode_override` / `payroll_mode_override` (or the single `vac_payroll_mode_override`) passed to `generate_word_output()`. `app.R` exposes UI toggles that store these in `reactiveVal`s (`selected_vac_period`, `auto_selected_vac_period`, etc.).

### Error handling convention

`utils/calculations.R` wraps every `calculate_*` call in `tryCatch` with the `.err()` helper, so one failing sheet never aborts the whole run — missing values surface as `NA` and then as `"\u2014"` (em dash) in the output via the `fmt_*` / `format_*` helpers in `utils/helpers.R`. When adding a metric, return `NA_real_` / `NA_character_` on failure rather than throwing, and guard list access with the local `.safe()` / `.safe_chr()` / `.safe_date()` helpers in `utils/calculations.R`.

### Formatting helpers

`utils/helpers.R` has two formatter families — **unsigned** (`fmt_*`: for the dashboard table) and **signed with +/- prefix** (`format_*` and `fmt_signed_*`: for narrative text and the Word briefing). Picking the wrong one is the most common source of regressions in narrative output. `utils/word_helpers.R` re-declares fallback versions that delegate to `helpers.R` when it has been sourced — don't duplicate logic, extend `helpers.R`.

### OECD data resolution

Uploaded OECD files always win; otherwise the app pre-fetches the three OECD rates from Postgres on session start via `fetch_oecd_from_db()` (shared by both tabs, same function). `.oecd_source(metric)` in `app.R` is the resolver — it returns a file path **or** an already-parsed data.frame (`country`/`period`/`value`), and every downstream consumer must accept either.

## Conventions specific to this codebase

- ONS suppression markers (`[x]`, `[c]`, `[z]`, `..`, `-`, `~`, `:`, `[e]`, `**`) appear in source spreadsheets; `.safe_read()` in `sheets/excel_audit_workbook.R` coerces columns to numeric while preserving the raw text as an attribute for later display.
- Dataset codes (e.g. `MGRZ`, `LF24`, `AP2Y`, `KAC9`) are ONS identifiers and must match the upstream source — don't rename them. Each `sheets/*.R` file has a `<NAME>_CODES` list at the top; lookups use `val_by_code()`.
- Template is `utils/ManualDB.docx` (tracked); generated output goes to `utils/DBoutput.docx` (gitignored). Don't rename either without updating the defaults in `generate_word_output()`.
- File-detection in `app.R` (`.detect_file_type`) relies on filename substrings (`a01`, `hr1`, `x09`, `rtisa`, `cla01`, `x02`, `oecd*`) and falls back to sheet-name sniffing. OECD sub-types are resolved from filename *or* content (`.detect_oecd_metric_from_content`).
- Indentation is 2 spaces (see `labour-market-stats-brief.Rproj`).

## Git workflow notes

The default branch is `main`. Feature work happens on `claude/*` branches. The repo has a `.gitignore` that excludes `utils/DBoutput.docx` — keep all regenerated outputs out of commits.
