"""Monte Carlo Validation Agent

Validates Monte Carlo forecast outputs for:
1. Directional consistency (decline worsens, improvement improves)
2. Monotonic progression (baseline → year 5 → year 10 should be consistent)
3. Physiological plausibility (values within realistic ranges)
"""

import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass
from typing import Any
import pandas as pd


@dataclass
class ValidationIssue:
    """Record of a validation issue."""
    client: str
    test: str
    scenario: str  # "Decline" or "Improvement"
    issue_type: str
    description: str
    baseline: float | None
    year5: float | None
    year10: float | None


class MonteCarloValidator:
    """Validator for Monte Carlo outputs."""

    def __init__(self, config_path: str | Path):
        self.config_path = Path(config_path)
        self._load_config()
        self.issues: list[ValidationIssue] = []

    def _load_config(self):
        """Load Monte Carlo configuration."""
        with open(self.config_path, 'r', encoding='utf-8-sig') as f:
            self.cfg = json.load(f)

        self.higher_is_better = set(self.cfg['higher_is_better'])
        self.lower_is_better = set(self.cfg['lower_is_better'])
        self.all_tests = self.cfg['tests_order']

        # Define physiological limits (reasonable ranges for each test)
        self.physiological_limits = {
            'VO2 max': (5.0, 90.0),  # ml/kg/min
            'FEV1': (20.0, 150.0),  # % predicted
            'Grip Strength': (5.0, 100.0),  # kg
            'STS Power': (0.5, 10.0),  # w/kg
            'Vertical Jump': (5.0, 100.0),  # cm
            'Body Fat %': (3.0, 60.0),  # %
            'Waist to Height Ratio': (0.3, 0.9),
            'HbA1c': (4.0, 15.0),  # %
            'HOMA-IR': (0.2, 20.0),
            'ApoB': (0.3, 3.0),  # g/L
            'hsCRP': (0.1, 50.0),  # mg/L
            'Gait Speed': (0.2, 2.5),  # m/s
            'Timed Up and Go': (3.0, 60.0),  # sec
            'Single Leg Stance': (1.0, 120.0),  # sec
            'Sit and Reach': (-30.0, 60.0),  # cm (can be negative)
            'Processing Speed': (-5.0, 5.0),  # z-score
            'Working Memory': (-5.0, 5.0),  # z-score
        }

    def read_excel_output(self, excel_path: Path) -> dict[str, dict[str, float]]:
        """Read Monte Carlo output Excel file.

        Returns:
            dict mapping client_name to dict of test values
        """
        z = zipfile.ZipFile(excel_path)

        # Read shared strings
        try:
            sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
            ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            shared_strings = []
            for si in sst.findall(f'.//{ns}si'):
                text = ''.join(t.text or '' for t in si.findall(f'.//{ns}t'))
                shared_strings.append(text)
        except KeyError:
            shared_strings = []

        # Read workbook to get sheets
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        ns_wb = {'w': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel_ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        rel_map = {rel.get('Id'): rel.get('Target') for rel in rels.findall('r:Relationship', rel_ns)}

        clients = {}
        for sheet in wb.findall('.//w:sheets/w:sheet', ns_wb):
            name = sheet.get('name')
            rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            path_sheet = 'xl/' + rel_map[rid]

            sheet_xml = ET.fromstring(z.read(path_sheet))
            rows = sheet_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')

            # Parse sheet as key-value pairs (column A = label, column B = value)
            data = {}
            for r in rows:
                cells = r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
                cell_dict = {}
                for c in cells:
                    ref = c.get('r')
                    col = ''.join([ch for ch in ref if ch.isalpha()])
                    t = c.get('t')
                    v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    if v is not None:
                        if t == 's':
                            try:
                                cell_dict[col] = shared_strings[int(v.text)]
                            except:
                                cell_dict[col] = v.text
                        else:
                            cell_dict[col] = v.text

                label = cell_dict.get('A', '').strip()
                value = cell_dict.get('B', '')
                if label:
                    try:
                        data[label] = float(value)
                    except:
                        data[label] = None

            clients[name] = data

        return clients

    def validate_scenario(self, scenario_name: str, excel_path: Path):
        """Validate one scenario (Decline or Improvement)."""
        print(f"\nValidating {scenario_name} scenario: {excel_path.name}")

        if not excel_path.exists():
            print(f"  [SKIP] File not found: {excel_path}")
            return

        clients = self.read_excel_output(excel_path)
        print(f"  Found {len(clients)} clients")

        for client_name, data in clients.items():
            self._validate_client(client_name, data, scenario_name)

    def _validate_client(self, client_name: str, data: dict[str, float], scenario: str):
        """Validate one client's predictions."""
        for test in self.all_tests:
            baseline = data.get(test)
            year5 = data.get(f"{test} (Year 5)")
            year10 = data.get(f"{test} (Year 10)")

            if baseline is None:
                continue  # Test not present for this client

            # Check 1: Physiological plausibility
            if test in self.physiological_limits:
                min_val, max_val = self.physiological_limits[test]
                for time_point, value in [("Baseline", baseline), ("Year 5", year5), ("Year 10", year10)]:
                    if value is not None and (value < min_val or value > max_val):
                        self.issues.append(ValidationIssue(
                            client=client_name,
                            test=test,
                            scenario=scenario,
                            issue_type="Physiological Implausibility",
                            description=f"{time_point} value {value:.2f} outside plausible range [{min_val}, {max_val}]",
                            baseline=baseline,
                            year5=year5,
                            year10=year10,
                        ))

            # Check 2: Directional consistency
            if year5 is not None:
                self._check_direction(client_name, test, scenario, baseline, year5, "Year 5")
            if year10 is not None:
                self._check_direction(client_name, test, scenario, baseline, year10, "Year 10")

            # Check 3: Monotonic progression
            if year5 is not None and year10 is not None:
                self._check_monotonic(client_name, test, scenario, baseline, year5, year10)

    def _check_direction(self, client: str, test: str, scenario: str, baseline: float, future: float, timepoint: str):
        """Check if change direction matches scenario expectation."""
        if baseline == 0:
            return  # Can't check direction from zero

        delta = future - baseline
        rel_change = delta / abs(baseline)

        # Tolerance: allow small deviations due to noise
        tolerance = 0.02  # 2% tolerance

        if test in self.higher_is_better:
            # Higher is better
            if scenario == "Decline" and rel_change > tolerance:
                self.issues.append(ValidationIssue(
                    client=client,
                    test=test,
                    scenario=scenario,
                    issue_type="Wrong Direction",
                    description=f"{timepoint}: Improved ({rel_change:+.1%}) on decline trajectory",
                    baseline=baseline,
                    year5=future if timepoint == "Year 5" else None,
                    year10=future if timepoint == "Year 10" else None,
                ))
            elif scenario == "Improvement" and rel_change < -tolerance:
                self.issues.append(ValidationIssue(
                    client=client,
                    test=test,
                    scenario=scenario,
                    issue_type="Wrong Direction",
                    description=f"{timepoint}: Declined ({rel_change:+.1%}) on improvement trajectory",
                    baseline=baseline,
                    year5=future if timepoint == "Year 5" else None,
                    year10=future if timepoint == "Year 10" else None,
                ))

        elif test in self.lower_is_better:
            # Lower is better
            if scenario == "Decline" and rel_change < -tolerance:
                self.issues.append(ValidationIssue(
                    client=client,
                    test=test,
                    scenario=scenario,
                    issue_type="Wrong Direction",
                    description=f"{timepoint}: Improved ({rel_change:+.1%}) on decline trajectory",
                    baseline=baseline,
                    year5=future if timepoint == "Year 5" else None,
                    year10=future if timepoint == "Year 10" else None,
                ))
            elif scenario == "Improvement" and rel_change > tolerance:
                self.issues.append(ValidationIssue(
                    client=client,
                    test=test,
                    scenario=scenario,
                    issue_type="Wrong Direction",
                    description=f"{timepoint}: Declined ({rel_change:+.1%}) on improvement trajectory",
                    baseline=baseline,
                    year5=future if timepoint == "Year 5" else None,
                    year10=future if timepoint == "Year 10" else None,
                ))

    def _check_monotonic(self, client: str, test: str, scenario: str, baseline: float, year5: float, year10: float):
        """Check if progression is monotonic (no reversals)."""
        if test in self.higher_is_better:
            if scenario == "Decline":
                # Should go: baseline >= year5 >= year10 (declining)
                if year5 > baseline or year10 > year5:
                    self.issues.append(ValidationIssue(
                        client=client,
                        test=test,
                        scenario=scenario,
                        issue_type="Non-Monotonic",
                        description=f"Values don't decline monotonically: {baseline:.2f} → {year5:.2f} → {year10:.2f}",
                        baseline=baseline,
                        year5=year5,
                        year10=year10,
                    ))
            else:  # Improvement
                # Should go: baseline <= year5 <= year10 (improving)
                if year5 < baseline or year10 < year5:
                    self.issues.append(ValidationIssue(
                        client=client,
                        test=test,
                        scenario=scenario,
                        issue_type="Non-Monotonic",
                        description=f"Values don't improve monotonically: {baseline:.2f} → {year5:.2f} → {year10:.2f}",
                        baseline=baseline,
                        year5=year5,
                        year10=year10,
                    ))

        elif test in self.lower_is_better:
            if scenario == "Decline":
                # Should go: baseline <= year5 <= year10 (worsening = increasing)
                if year5 < baseline or year10 < year5:
                    self.issues.append(ValidationIssue(
                        client=client,
                        test=test,
                        scenario=scenario,
                        issue_type="Non-Monotonic",
                        description=f"Values don't worsen monotonically: {baseline:.2f} → {year5:.2f} → {year10:.2f}",
                        baseline=baseline,
                        year5=year5,
                        year10=year10,
                    ))
            else:  # Improvement
                # Should go: baseline >= year5 >= year10 (improving = decreasing)
                if year5 > baseline or year10 > year5:
                    self.issues.append(ValidationIssue(
                        client=client,
                        test=test,
                        scenario=scenario,
                        issue_type="Non-Monotonic",
                        description=f"Values don't improve monotonically: {baseline:.2f} → {year5:.2f} → {year10:.2f}",
                        baseline=baseline,
                        year5=year5,
                        year10=year10,
                    ))

    def generate_report(self, output_path: Path):
        """Generate validation report."""
        if not self.issues:
            print("\n*** VALIDATION PASSED ***")
            print("No issues found! All outputs look correct.")
            return

        print(f"\n*** VALIDATION FOUND {len(self.issues)} ISSUES ***")

        # Create DataFrame
        issues_data = []
        for issue in self.issues:
            issues_data.append({
                'Client': issue.client,
                'Test': issue.test,
                'Scenario': issue.scenario,
                'Issue Type': issue.issue_type,
                'Description': issue.description,
                'Baseline': issue.baseline,
                'Year 5': issue.year5,
                'Year 10': issue.year10,
            })

        df = pd.DataFrame(issues_data)

        # Write to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Issues', index=False)

            # Summary sheet
            summary_data = {
                'Metric': [
                    'Total Issues',
                    'Clients Affected',
                    'Tests Affected',
                    'Wrong Direction',
                    'Non-Monotonic',
                    'Physiological Implausibility',
                ],
                'Count': [
                    len(self.issues),
                    len(set(i.client for i in self.issues)),
                    len(set(i.test for i in self.issues)),
                    len([i for i in self.issues if i.issue_type == 'Wrong Direction']),
                    len([i for i in self.issues if i.issue_type == 'Non-Monotonic']),
                    len([i for i in self.issues if i.issue_type == 'Physiological Implausibility']),
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        print(f"\nValidation report written to: {output_path}")

        # Print summary to console
        print("\n=== ISSUE TYPES ===")
        for issue_type in ['Wrong Direction', 'Non-Monotonic', 'Physiological Implausibility']:
            count = len([i for i in self.issues if i.issue_type == issue_type])
            if count > 0:
                print(f"  {issue_type}: {count}")

        # Show a few examples
        print("\n=== SAMPLE ISSUES (first 5) ===")
        for i, issue in enumerate(self.issues[:5], 1):
            print(f"\n{i}. {issue.client} - {issue.test} ({issue.scenario})")
            print(f"   Type: {issue.issue_type}")
            print(f"   {issue.description}")
            print(f"   Values: {issue.baseline:.2f} → {issue.year5 if issue.year5 else 'N/A'} → {issue.year10 if issue.year10 else 'N/A'}")


def main():
    """Main validation workflow."""
    base_path = Path(__file__).parent

    config_path = base_path / "mc_assumptions.json"
    decline_path = base_path / "outputs" / "MonteCarlo_Decline.xlsx"
    improvement_path = base_path / "outputs" / "MonteCarlo_Improvement.xlsx"
    report_path = base_path / "outputs" / "monte_carlo_validation_report.xlsx"

    print("=== MONTE CARLO VALIDATION ===\n")

    validator = MonteCarloValidator(config_path)

    # Validate both scenarios
    validator.validate_scenario("Decline", decline_path)
    validator.validate_scenario("Improvement", improvement_path)

    # Generate report
    validator.generate_report(report_path)


if __name__ == "__main__":
    main()
