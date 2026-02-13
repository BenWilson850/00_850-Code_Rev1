# Matrix Pipeline

Pipeline for the persona workbook activity matrix: read the limits matrix, classifications matrix, and client workbook, apply the DOCX aggregation rules, and write a consolidated Excel report (one sheet per client).

## Requirements

- Python 3.10+
- Dependencies: `pandas`, `openpyxl`, `PyYAML`; optional: `python-docx` for DOCX input

## Setup

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

## Running (Windows, py launcher)

From the repo root:

```powershell
py -m src.pipeline.cli --clients path\to\client_workbook.xlsx --limits path\to\limits.xlsx --classifications path\to\classifications.xlsx -o out
```

Optional:

- `-c configs/default.yaml` - config file (default)
- `-d path\to\spec.docx` - DOCX logic matrix (rules)
- `-v` - verbose (DEBUG) logging

Example with included sample inputs:

```powershell
py -m src.pipeline.cli --clients "Sample Persona Database.xlsx" --limits "Exercise limits optimised.xlsx" --classifications "Exercise Threshold Classifications.xlsx" -d "Activities Threshold Logic Matrix.docx" -o out -v
```

## Layout

- `src/pipeline/cli.py` - entrypoint
- `src/pipeline/core/` - persona assessment engine
- `src/pipeline/io/` - Excel and DOCX readers
- `src/pipeline/report/` - Excel report writer
- `src/pipeline/utils/` - logging and helpers
- `configs/default.yaml` - default configuration
- `out/` - report output (ignored)

## Config

Edit `configs/default.yaml` to set default paths and sheet/header settings. Override at runtime via CLI flags.
