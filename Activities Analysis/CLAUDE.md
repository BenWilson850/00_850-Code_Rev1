# Persona Matrix Pipeline (Claude Code)

When the user asks to run a "Current vs. Future State" analysis:

1. Ask for the Client Workbook Excel file path (persona database). The repo already contains the reference matrices by default.
2. Run the pipeline from the repo root:

```powershell
.\.venv\Scripts\python.exe -m src.pipeline.cli --clients "<CLIENT_WORKBOOK.xlsx>" -o out -v
```

3. Return the generated report path: `out/persona_activity_report.xlsx`.

Notes:
- Limits matrix: `Exercise limits optimised.xlsx` (default via `configs/default.yaml`).
- Classifications matrix (Critical/Supporting): `Exercise Threshold Classifications.xlsx` (default).
- Aggregation rules: `Activities Threshold Logic Matrix.docx` (default).
