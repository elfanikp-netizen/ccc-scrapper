# CCC Workfile Crawler

Reads `VIN` and `OEM` columns from an Excel workbook, drives CCC in Chromium, and incrementally writes vendor price results to an Excel workbook. Progress and errors are printed to the terminal. On a failed UI step, a screenshot and visible page text are saved under `artifacts/`.

## Setup

Use a virtual environment and install the requirements:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
Copy-Item .env.example .env
```

Set `CCC_USERNAME` and `CCC_PASSWORD` in `.env` on your machine. Credentials are not included in this repository. Put the input workbook at `input.xlsx` with header columns named `VIN` and `OEM` (case-insensitive), or set `CCC_INPUT` in `.env`.

## Run

Run with the browser visible:

```powershell
python crawler.py --headed
```

Run without opening a browser window:

```powershell
python crawler.py
```

Choose different input and result workbooks for a run with `--input` and `--output`:

```powershell
python crawler.py --headed --input .\data\parts.xlsx --output .\results\run-01.xlsx
```

The input workbook must have `VIN` and `OEM` header columns. The default paths are `input.xlsx` and `ccc-results.xlsx`; set `CCC_INPUT` and `CCC_OUTPUT` in `.env` to change those defaults. Command-line paths override the `.env` values. Each result row includes `date_time` in local `YYYY-MM-DD HH:MM:SS` format, recorded when that row's crawl processing finishes. Authentication cookies/local storage are saved to `playwright/.auth/ccc-state.json` and reused on later runs; delete that file to force a fresh login.

## Notes

- Use only an account and data you are authorized to access, and confirm the site's automation terms before running at scale.
- CCC's interface may change or use controls that do not expose accessible labels. The crawler reports the failed step and saves page evidence so selectors can be adapted without losing completed rows.
- “Next Lowest” is the second-lowest A/M: CAPA offer across all vendors. “Lowest CCC” and its supplier exclude iPart Automotive Inc. If no matching offer exists, the price/supplier cells are left blank and `Status` explains the result.