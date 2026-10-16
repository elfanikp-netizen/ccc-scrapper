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

```powershell
python crawler.py --headed
```

The first run is headed so the login and any unexpected prompts can be observed. `--headed` can be omitted for a headless run. Authentication cookies/local storage are saved to `playwright/.auth/ccc-state.json` and reused on later runs; delete that file to force a fresh login. Results are saved incrementally to `ccc-results.xlsx` (or `CCC_OUTPUT`).

## Notes

- Use only an account and data you are authorized to access, and confirm the site's automation terms before running at scale.
- CCC's interface may change or use controls that do not expose accessible labels. The crawler reports the failed step and saves page evidence so selectors can be adapted without losing completed rows.
- “Next Lowest” is the second-lowest A/M: CAPA offer across all vendors. “Lowest CCC” and its supplier exclude iPart Automotive Inc. If no matching offer exists, the price/supplier cells are left blank and `Status` explains the result.