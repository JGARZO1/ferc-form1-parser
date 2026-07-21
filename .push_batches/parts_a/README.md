# ferc-form1-parser

**Robust, reusable pipeline to extract standardized, R-friendly Excel files from FERC Form 1 HTML (iXBRL) and XBRL filings.**

Processes five key schedules for multiple utilities and years, writing clean outputs into company folders (CLECO, SWEPCO, ENTERGY LA, ENTERGY NOLA, and extensible).

---

## Schedules

| Schedule | Template | Output prefix |
|----------|----------|---------------|
| About / Contact Info | `about.xlsx` | `about_<utilityId>_<year>.xlsx` |
| Sales of Electricity by Rate Schedules | `sales_by_rate.xlsx` | `sales_by_rate_<utilityId>_<year>.xlsx` |
| Sales for Resale (Account 447) | `sales_for_resale.xlsx` | `sales_for_resale_<utilityId>_<year>.xlsx` |
| Electric Plant in Service | `plant_in_service.xlsx` | `plant_in_service_<utilityId>_<year>.xlsx` |
| Electric Operation & Maintenance Expenses | `om_expenses.xlsx` | `om_expenses_<utilityId>_<year>.xlsx` |

Templates use **snake_case** headers and metadata rows (`utility_id`, `report_year`) so they load cleanly in R/tidyverse (`readxl`, `dplyr`, etc.).

---

## Features

- **Modular fill scripts** for HTML and XBRL (parallel `fill_*.js` / `*_xbrl.js` APIs)
- **Master runner** `run_all_ferc.js` — auto-detects `.html` / `.htm` vs `.xbrl` / `.xml`
- **Company folder reuse** (case-insensitive; no duplicate folders)
- **Shared XBRL utilities** in `xbrl_common.js` (contexts, facts, segments)
- **Cross-platform Excel packing** (macOS/Linux `zip`/`unzip`; Windows PowerShell fallback)
- **Example outputs** for CLECO, SWEPCO, ENTERGY LA, ENTERGY NOLA (HTML pipeline samples, ~2021–2024)

---

## Requirements

- **Node.js** 16+ (no npm dependencies required)
- **zip** and **unzip** on PATH (preinstalled on macOS; install via package manager on Linux)
- On Windows: PowerShell (built-in) *or* zip/unzip

```bash
node --version   # should print v16+
```

---

## Installation / setup

```bash
git clone https://github.com/JGARZO1/ferc-form1-parser.git
cd ferc-form1-parser
# No npm install required
```

Place FERC Form 1 source files in the repo root or under `filings/`:

- **HTML / iXBRL renders** (common for recent years, e.g. 2021+)
- **XBRL instances** (`.xbrl` / `.xml`, common for earlier years)

A sample XBRL instance is included:

```
filings/ClecoPowerLlc2020.xbrl
```

Additional filings are available from [FERC eCollection](https://ecollection.ferc.gov/).

---

## Usage

### Process an entire filing (recommended)

```bash
# HTML (iXBRL)
node run_all_ferc.js ClecoPowerLlc2023.html
node run_all_ferc.js path/to/EntergyLouisianaLlc2022.html

# XBRL
node run_all_ferc.js filings/ClecoPowerLlc2020.xbrl
```

The runner:

1. Detects format (HTML vs XBRL)
2. Reads `utility_id`, `report_year`, and respondent name
3. Maps the company to a short folder name (e.g. CLECO)
4. Reuses or creates that folder
5. Runs all five schedule extractors
6. Writes `schedule_utilityId_year.xlsx` files into the company folder

### Run a single schedule

```bash
# HTML
node fill_about_info.js ClecoPowerLlc2023.html
node fill_sales_by_rate_totals.js ClecoPowerLlc2023.html
node fill_sales_for_resale.js ClecoPowerLlc2023.html
node fill_plant_in_service.js ClecoPowerLlc2023.html
node fill_om_expenses.js ClecoPowerLlc2023.html

# XBRL
node fill_about_info_xbrl.js filings/ClecoPowerLlc2020.xbrl
node fill_sales_by_rate_totals_xbrl.js filings/ClecoPowerLlc2020.xbrl
node fill_sales_for_resale_xbrl.js filings/ClecoPowerLlc2020.xbrl
node fill_plant_in_service_xbrl.js filings/ClecoPowerLlc2020.xbrl
node fill_om_expenses_xbrl.js filings/ClecoPowerLlc2020.xbrl
```

Optional args: `[source] [template] [outputFileName]`

---

## Folder structure

```
ferc-form1-parser/
├── run_all_ferc.js              # Master runner (HTML + XBRL)
├── fill_*.js                    # HTML schedule parsers
├── fill_*_xbrl.js               # XBRL schedule parsers
├── xbrl_common.js               # Shared XBRL parse utilities
├── lib/
│   ├── company.js               # Company aliases & folder resolution
│   ├── xlsx_utils.js            # Cross-platform xlsx pack/unpack
│   ├── plant_xbrl_map.js        # Plant line_no → concept map
│   └── om_xbrl_map.js           # O&M line_no → concept map
├── templates/                   # R-ready Excel templates (source of truth)
│   ├── about.xlsx
│   ├── sales_by_rate.xlsx
│   ├── sales_for_resale.xlsx
│   ├── plant_in_service.xlsx
│   └── om_expenses.xlsx
├── about.xlsx …                 # Root copies of templates (CLI defaults)
├── filings/                     # Drop source HTML/XBRL here
│   ├── README.md
│   └── ClecoPowerLlc2020.xbrl   # Sample XBRL
├── CLECO/                       # Example / live outputs
├── SWEPCO/
├── ENTERGY LA/
├── ENTERGY NOLA/
├── examples/                    # Snapshot copies of sample outputs
├── package.json
├── LICENSE                      # MIT
└── README.md
```

### Output naming

```
<COMPANY>/<schedule>_<utilityId>_<year>.xlsx
```

Examples:

- `CLECO/about_C000447_2023.xlsx`
- `SWEPCO/plant_in_service_C000537_2021.xlsx`
- `ENTERGY LA/om_expenses_C004995_2024.xlsx`

---

## Companies

Built-in aliases (first match wins):

| Pattern | Folder |
|---------|--------|
| Entergy New Orleans | `ENTERGY NOLA` |
| Entergy Louisiana | `ENTERGY LA` |
| Cleco | `CLECO` |
| SWEPCO / Southwestern Electric Power | `SWEPCO` |
| Entergy (generic) | `ENTERGY` |

Unknown respondents fall back to a cleaned brand token from the legal name or filename.

### Adding a new company

1. Drop a filing whose respondent name or filename contains a unique token.
2. Optionally add an alias in `lib/company.js` → `COMPANY_ALIASES` **and** in `run_all_ferc.js` / per-script aliases if you still use older local copies of the helper.
3. Re-run `node run_all_ferc.js <filing>` — the folder is created once and reused.

---

## Adding years

No code changes required for a new year of the same company:

```bash
node run_all_ferc.js EntergyNewOrleansInc2024.html
# → ENTERGY NOLA/about_C007667_2024.xlsx, …
```

Report year is taken from the filing (`ferc:ReportYear`) with a filename fallback (`…2024.html`).

---

## Extending for a new schedule

1. **Design a template** under `templates/` with:
   - Row 1–2 metadata: `utility_id`, `report_year` in column A; values in B
   - Header row with snake_case column names
   - Stable `line_no` keys in column A for row matching (when applicable)
2. **HTML parser** — scrape the schedule region from the iXBRL HTML table (`fill_<name>.js`).
3. **XBRL parser** — map FERC taxonomy concepts (and axes) in `fill_<name>_xbrl.js`; add concept maps under `lib/` if line-based.
4. **Register both** in `run_all_ferc.js` (`PIPELINE_HTML` and `PIPELINE_XBRL`).
5. Document the template columns in this README.

---

## Reading outputs in R

```r
library(readxl)
library(dplyr)

about <- read_excel("CLECO/about_C000447_2023.xlsx")
plant <- read_excel("CLECO/plant_in_service_C000447_2023.xlsx", skip = 2)
# or use the snake_case header row as appropriate for each template
```

Metadata lives in the first two rows of each workbook; data tables start at row 3–4 depending on the schedule.

---

## HTML vs XBRL

| | HTML (iXBRL) | XBRL instance |
|--|--------------|---------------|
| Typical years | ~2021+ rendered filings | Often 2011–2020 (and raw instances) |
| Parsing approach | Table row scrape + `ix:` facts | `ferc:` facts + contexts/segments |
| Strengths | Matches paper form layout; good for multi-column tables | Structured concepts; multi-year machine data |
| Scripts | `fill_*.js` | `fill_*_xbrl.js` + `xbrl_common.js` |

Both write the **same template layout** so R code can stay format-agnostic.

---

## Known limitations

- **HTML parsers** depend on FERC’s rendered table structure and section anchors; major form redesigns may need selector updates.
- **XBRL plant / O&M** coverage depends on concept maps in `lib/plant_xbrl_map.js` and `lib/om_xbrl_map.js`. Taxonomy renames or company-specific extensions may leave some lines blank.
- **Sales for resale (XBRL)** detail rows come from dimensional segments; RQ / non-RQ subtotals are filled only when the taxonomy provides clear aggregates.
- **Sales by rate (HTML)** currently extracts **All-Accounts totals** (lines 41–43), not every rate schedule line.
- Source HTML/XBRL files can be **very large** (10–20+ MB each) and are gitignored by default (except the sample Cleco 2020 XBRL).
- Example company folders in the repo are primarily from the **HTML** pipeline (roughly 2021–2024). Generate earlier years by running the XBRL path on your local instances.

---

## Future improvements

- Full rate-schedule line detail (not only All-Accounts totals)
- Depreciation & amortization schedule (template stub exists in some workspaces)
- Batch mode: `node run_all_ferc.js filings/*.xbrl`
- Unit tests comparing HTML vs XBRL extracts for overlapping years
- Richer Entergy / multi-state alias tables
- Optional pure-JS zip (no system `zip`/`unzip`)

---

## License

MIT — see [LICENSE](LICENSE).

FERC Form 1 data is public regulatory information. This project is not affiliated with FERC.
