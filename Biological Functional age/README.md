# Biological Functional Age (BFA) Calculation Pipeline

An automated pipeline for calculating **Biological Functional Age** (BFA) and **Healthspan Index** (300-850) from 16 gold-standard health tests across 5 physiological pillars.

## Overview

This pipeline implements the Basis 850 framework for assessing functional health:
- Compares client test results to age/gender-specific normative data
- Calculates functional ages for 5 pillars (Vitality, Strength, Metabolic, Mobility, Cognitive)
- Aggregates into a single Biological Functional Age
- Generates Healthspan Index score and category (Critical → Elite)

## Features

- **16 Health Tests** across 5 pillars
- **Normative data interpolation** for precise age matching
- **Risk-based scoring** for metabolic markers
- **Weighted aggregation** with evidence-based weights
- **Batch processing** - single or multi-client workbooks
- **CSV/Excel output** with full detail (29 columns)

## Installation

```bash
# Create virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python -m src.bfa_pipeline.cli "Sample Persona Database (BFA).xlsx"
```

### Advanced Options

```bash
python -m src.bfa_pipeline.cli \
    "Sample Persona Database (BFA).xlsx" \
    --normative-db "Normative Database.xlsx" \
    --config-dir "configs" \
    --output "output/results.xlsx" \
    --verbose
```

### Arguments

- `client_file` - Path to client test data Excel workbook (required)
- `-n, --normative-db` - Path to normative database (default: `Normative Database.xlsx`)
- `-c, --config-dir` - Path to configuration directory (default: `configs`)
- `-o, --output` - Output file path (.csv or .xlsx) (default: `output/bfa_results.csv`)
- `-v, --verbose` - Enable verbose output

## Input Format

Client data should be in Excel format with one sheet per client:

**Template Structure:**
- Row 1: Gender (Male/Female)
- Row 2: Name
- Row 4: Age
- Rows 5-22: 16 test results (see below)
- Column B: Current test values

**Required Tests (16):**
1. VO2 max (ml/kg/min)
2. FEV1 (% predicted)
3. Grip strength (kg)
4. STS Power (W/kg)
5. Vertical jump (cm)
6. Body fat (%)
7. Waist-to-height ratio
8. Fasting glucose (mmol/L)
9. HbA1c (%)
10. HOMA IR
11. ApoB (g/L)
12. hsCRP (mg/L)
13. Gait speed (m/s)
14. Timed Up and Go (sec)
15. Single leg stance (sec)
16. Sit and reach (cm)
17. Processing speed (SD)
18. Working memory (SD)

## Output Format

Results are written to CSV or Excel with 29 columns:

**Client Metadata:**
- Name, Age, Gender

**Raw Test Results (16 columns):**
- All test values as input

**Pillar Functional Ages (5 columns):**
- Vitality_Functional_Age
- Strength_Functional_Age
- Metabolic_Functional_Age
- Mobility_Functional_Age
- Cognitive_Functional_Age

**Final Outputs (3 columns):**
- Biological_Functional_Age
- Healthspan_Index
- Healthspan_Category

**Healthspan Categories:**
- Critical (300-459): Critical functional decline
- Poor (460-549): Significant limitations
- Fair (550-599): Functional but fragile
- Average (600-649): Typical for age
- Good (650-709): Strong foundation
- Excellent (710-799): High functional capacity
- Elite (800-850): Exceptional resilience

## Configuration

### Pillar Weights (`configs/weights.yaml`)

- Metabolic: 25%
- Vitality: 20%
- Strength: 20%
- Mobility: 20%
- Cognitive: 15%

### Sub-Test Weights

**Vitality:**
- VO2 Max: 70%, FEV1: 30%

**Strength:**
- Grip: 50%, STS: 25%, Vertical Jump: 25%

**Metabolic:**
- ApoB: 30%, HOMA-IR: 25%, HbA1c: 15%, WHtR: 15%, hsCRP: 10%, Body Fat: 5%

**Mobility:**
- Single Leg Stance: 40%, Gait Speed: 30%, TUG: 20%, Sit & Reach: 10%

**Cognitive:**
- Processing Speed: 60%, Working Memory: 40%

## Methodology

### Calculation Pipeline

1. **Individual Test Scoring**
   - Physical tests: Find functional age via normative data interpolation
   - Vitality: Special formula using age/gender-specific VO2 norm
   - Metabolic: Risk-based 0-100 scoring, then convert to functional age
   - Cognitive: Convert SD to functional age (1 SD = 25 years)

2. **Pillar Aggregation**
   - Weight individual test functional ages within each pillar
   - Metabolic: First calculate weighted risk index, then convert to functional age

3. **BFA Calculation**
   - Weighted average of 5 pillar functional ages
   - If any test missing → mark as INCOMPLETE

4. **Healthspan Index**
   - Formula: `670 + (6.5 × (Chronological_Age - BFA))`
   - Clamped: 300-850
   - Categorized: Critical → Elite

### Normative Data

- Age/gender-specific reference values from validated studies
- Linear interpolation for precise age matching
- Range-based data interpolates between midpoints

## Examples

### Example 1: Single Client

```bash
python -m src.bfa_pipeline.cli "client_data.xlsx" -o "output/client_results.csv"
```

### Example 2: Batch Processing

```bash
python -m src.bfa_pipeline.cli \
    "Sample Persona Database (BFA).xlsx" \
    -o "output/batch_results.xlsx" \
    --verbose
```

## Error Handling

- **Missing tests:** BFA marked as "INCOMPLETE"
- **Invalid data:** Client skipped with warning
- **Missing normative data:** Error raised

## Project Structure

```
Biological Functional age/
├── configs/
│   ├── weights.yaml
│   └── healthspan_categories.yaml
├── src/
│   └── bfa_pipeline/
│       ├── io/
│       │   ├── client_reader.py
│       │   └── normative_reader.py
│       ├── core/
│       │   ├── interpolation.py
│       │   ├── test_scoring.py
│       │   ├── metabolic_scoring.py
│       │   ├── pillar_scoring.py
│       │   ├── bfa_calculation.py
│       │   └── healthspan_index.py
│       ├── report/
│       │   └── output_writer.py
│       └── cli.py
├── Normative Database.xlsx
├── Sample Persona Database (BFA).xlsx
├── requirements.txt
└── README.md
```

## License

Proprietary - Basis 850 Framework

## Support

For questions or issues, contact the Basis 850 team.
