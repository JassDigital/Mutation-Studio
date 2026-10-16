# Mutation Studio

A local-first Windows desktop dashboard for reviewing mutation pendency reports. Built with Python and PySide6.

## Features

- Import Excel workbooks (`.xlsx` and `.xlsm`) by detecting the report header row.
- Paste tab-separated rows copied directly from Excel.
- Dashboard summary cards for total records, incorporated mutations, entry-level records, and missing registry details.
- Search across mutation number, dates, status, and registry fields.
- Filter by mutation status and sort by date or mutation number.
- Export the currently filtered records to CSV.
- Runs locally; no account or cloud service required.

## Quick start (Windows)

Install Python 3.11 or newer, then open PowerShell in this folder and run:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

If script activation is restricted, run `py -m pip install -r requirements.txt` followed by `py main.py`.

## Import data

Use **Import Excel** to load an Excel workbook, or **Paste data** to paste cells copied from Excel. The expected columns are:

`S.No. | Mutation No | Mutation Date | Mutation Status | Registry Number | Registry Date`

Dates are normalized where possible to `DD/MM/YYYY`. Use **Export CSV** to save the filtered table.

## Demo data and privacy

The included `data/sample_report.json` contains synthetic demonstration records only. Do not commit real mutation, registry, landowner, or other sensitive government data to a public repository. Import operational data locally when using the app.

## Repository publishing

Run `./publish-github.ps1` in PowerShell after installing Git and GitHub CLI (`gh`) and authenticating with `gh auth login`. The script shows the planned repository and requires you to type `YES` before creating and pushing a public repository.

## License

MIT. See `LICENSE`.


GitHub repository: https://github.com/JassDigital/Mutation-Studio
