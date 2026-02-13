"""Validation Agent for Activities Analysis Output

This agent validates persona_activity_report.xlsx by:
1. Checking that Critical/Supporting test classifications are correct
2. Verifying final status calculations follow the threshold logic rules
3. Auto-correcting errors and generating a corrections report
"""

import pandas as pd
import re
from pathlib import Path
from datetime import datetime
from typing import Any
from dataclasses import dataclass

# Import the logic functions from the pipeline
import sys
sys.path.insert(0, str(Path(__file__).parent))
from src.pipeline.core.persona_logic import (
    infer_importance,
    apply_activity_rules,
    AggregationRules,
    DEFAULT_RULES,
    Zone
)
from src.pipeline.io.docx_readers import read_docx_spec


@dataclass
class Correction:
    """Record of a correction made."""
    client: str
    activity: str
    horizon: str  # "5" or "10"
    error_type: str
    description: str
    old_value: Any
    new_value: Any


class ValidationAgent:
    """Agent to validate and correct Activities Analysis output."""

    def __init__(
        self,
        limits_path: str | Path,
        classifications_path: str | Path,
        docx_path: str | Path | None = None,
    ):
        self.limits_path = Path(limits_path)
        self.classifications_path = Path(classifications_path)
        self.docx_path = Path(docx_path) if docx_path else None

        # Load reference matrices
        self._load_matrices()

        # Load aggregation rules (use defaults - DOCX parsing is unreliable)
        self.rules = DEFAULT_RULES
        print(f"Using aggregation rules: >>{self.rules.supporting_red_for_red} supporting RED = RED, "
              f"{self.rules.supporting_red_for_yellow} supporting RED = YELLOW")

        # Track corrections
        self.corrections: list[Correction] = []

    def _load_matrices(self):
        """Load the limits and classifications matrices."""
        print(f"Loading limits matrix from {self.limits_path}...")
        self.limits_df = pd.read_excel(
            self.limits_path,
            sheet_name='Activities Thresholds_Rev2',
            header=0
        )
        self.limits_df['Activity'] = self.limits_df['Activity'].astype(str).str.strip()
        self.limits_df = self.limits_df.set_index('Activity')

        print(f"Loading classifications matrix from {self.classifications_path}...")
        self.class_df = pd.read_excel(
            self.classifications_path,
            sheet_name='Activities Thresholds_Rev2',
            header=0
        )
        self.class_df['Activity'] = self.class_df['Activity'].astype(str).str.strip()
        self.class_df = self.class_df.set_index('Activity')

        # Get test columns (exclude metadata columns)
        meta_cols = ['Unnamed: 0', 'Evidence Quality', 'Key References']
        self.test_cols = [c for c in self.limits_df.columns if c not in meta_cols]

        print(f"Found {len(self.test_cols)} test columns")

    def _load_rules(self) -> AggregationRules:
        """Load aggregation rules from DOCX if available."""
        if not self.docx_path or not self.docx_path.exists():
            print("Using default aggregation rules")
            return DEFAULT_RULES

        print(f"Loading aggregation rules from {self.docx_path}...")
        spec = read_docx_spec(self.docx_path)
        if not spec or not spec.get("tables"):
            return DEFAULT_RULES

        table0 = spec["tables"][0] if spec["tables"] else []
        supporting_red_for_red = DEFAULT_RULES.supporting_red_for_red
        supporting_red_for_yellow = DEFAULT_RULES.supporting_red_for_yellow

        for row in table0[1:]:
            if len(row) < 2:
                continue
            trigger = str(row[1]).lower()
            if "supporting" in trigger and ">" in trigger and "red" in trigger:
                nums = [int(n) for n in re.findall(r"(\d+)", trigger)]
                if nums:
                    supporting_red_for_red = max(nums)
            if "supporting" in trigger and " are" in trigger and "red" in trigger and ">" not in trigger:
                nums = [int(n) for n in re.findall(r"(\d+)", trigger)]
                if nums:
                    supporting_red_for_yellow = max(nums)

        rules = AggregationRules(
            supporting_red_for_red=supporting_red_for_red,
            supporting_red_for_yellow=supporting_red_for_yellow,
        )
        print(f"Loaded rules: >>{rules.supporting_red_for_red} supporting RED = RED, "
              f"{rules.supporting_red_for_yellow} supporting RED = YELLOW")
        return rules

    def _parse_test_name(self, text: str) -> str:
        """Extract test name from a failure string like 'VO2 Max (ml/kg/min) (RED)'."""
        # Remove the zone indicator (RED), (YELLOW), (MISSING)
        text = re.sub(r'\s*\((RED|YELLOW|GREEN|MISSING)\)\s*$', '', text)
        return text.strip()

    def _parse_failures(self, failures_str: str) -> list[tuple[str, str]]:
        """Parse a failures string into list of (test_name, zone) tuples."""
        if pd.isna(failures_str) or not str(failures_str).strip():
            return []

        # Split by commas and parse each
        parts = str(failures_str).split(',')
        results = []
        for part in parts:
            part = part.strip()
            if not part:
                continue
            # Extract zone from end
            zone_match = re.search(r'\((RED|YELLOW|GREEN|MISSING)\)$', part)
            if zone_match:
                zone = zone_match.group(1)
                test_name = self._parse_test_name(part)
                results.append((test_name, zone))

        return results

    def _normalize_test_name(self, test_name: str) -> str:
        """Normalize test name for matching (remove extra whitespace, newlines)."""
        return ' '.join(test_name.split())

    def _find_matching_column(self, test_name: str) -> str | None:
        """Find the matching column name in the matrices for a test name."""
        test_norm = self._normalize_test_name(test_name)

        for col in self.test_cols:
            col_norm = self._normalize_test_name(col)
            if test_norm == col_norm or test_norm in col_norm or col_norm in test_norm:
                return col
        return None

    def _get_true_importance(self, activity: str, test_name: str) -> str | None:
        """Get the true importance (Critical/Supporting) for a test from the classifications matrix."""
        col = self._find_matching_column(test_name)
        if not col:
            return None

        if activity not in self.class_df.index:
            return None

        raw_importance = self.class_df.at[activity, col]
        return infer_importance(raw_importance)

    def validate_activity(
        self,
        client: str,
        activity: str,
        critical_failures: str,
        supporting_failures: str,
        final_status: str,
        horizon: str,
    ) -> dict[str, Any]:
        """Validate one activity for one client at one time horizon.

        Returns:
            dict with 'corrected' flag and corrected values if needed
        """
        # Parse the failures
        crit_tests = self._parse_failures(critical_failures)
        supp_tests = self._parse_failures(supporting_failures)

        # Check each test classification
        misclassified_crit_to_supp = []  # Critical tests marked as supporting
        misclassified_supp_to_crit = []  # Supporting tests marked as critical

        for test_name, zone in crit_tests:
            true_importance = self._get_true_importance(activity, test_name)
            if true_importance == "Supporting":
                misclassified_supp_to_crit.append((test_name, zone))

        for test_name, zone in supp_tests:
            true_importance = self._get_true_importance(activity, test_name)
            if true_importance == "Critical":
                misclassified_crit_to_supp.append((test_name, zone))

        # Build corrected lists
        corrected_crit = [t for t in crit_tests if t not in misclassified_supp_to_crit]
        corrected_crit.extend(misclassified_crit_to_supp)

        corrected_supp = [t for t in supp_tests if t not in misclassified_crit_to_supp]
        corrected_supp.extend(misclassified_supp_to_crit)

        # Recalculate status
        critical_zones = [z for _, z in corrected_crit]
        supporting_zones = [z for _, z in corrected_supp if z == "RED"]

        calculated_status = apply_activity_rules(
            critical_zones=critical_zones,
            supporting_zones=supporting_zones,
            rules=self.rules,
        )

        # Check for corrections needed
        needs_correction = False

        if misclassified_crit_to_supp or misclassified_supp_to_crit:
            needs_correction = True
            if misclassified_crit_to_supp:
                self.corrections.append(Correction(
                    client=client,
                    activity=activity,
                    horizon=horizon,
                    error_type="Misclassification",
                    description=f"{len(misclassified_crit_to_supp)} Critical test(s) incorrectly listed as Supporting",
                    old_value=f"Supporting: {', '.join(t[0] for t in misclassified_crit_to_supp)}",
                    new_value=f"Moved to Critical",
                ))
            if misclassified_supp_to_crit:
                self.corrections.append(Correction(
                    client=client,
                    activity=activity,
                    horizon=horizon,
                    error_type="Misclassification",
                    description=f"{len(misclassified_supp_to_crit)} Supporting test(s) incorrectly listed as Critical",
                    old_value=f"Critical: {', '.join(t[0] for t in misclassified_supp_to_crit)}",
                    new_value=f"Moved to Supporting",
                ))

        if calculated_status != final_status:
            needs_correction = True
            self.corrections.append(Correction(
                client=client,
                activity=activity,
                horizon=horizon,
                error_type="Status Calculation",
                description=f"Final status should be {calculated_status} based on threshold logic",
                old_value=final_status,
                new_value=calculated_status,
            ))

        # Format corrected strings
        corrected_crit_str = ", ".join(f"{t[0]} ({t[1]})" for t in corrected_crit)
        corrected_supp_str = ", ".join(f"{t[0]} ({t[1]})" for t in corrected_supp)

        return {
            'needs_correction': needs_correction,
            'corrected_critical': corrected_crit_str,
            'corrected_supporting': corrected_supp_str,
            'corrected_status': calculated_status,
        }

    def validate_report(self, report_path: str | Path) -> Path:
        """Validate an entire report and generate corrected version.

        Args:
            report_path: Path to persona_activity_report.xlsx

        Returns:
            Path to the corrected report
        """
        report_path = Path(report_path)
        if not report_path.exists():
            raise FileNotFoundError(f"Report not found: {report_path}")

        print(f"\n=== VALIDATING REPORT: {report_path.name} ===\n")

        # Load the Excel file
        xl = pd.ExcelFile(report_path)

        # Prepare output file path
        corrected_path = report_path.parent / f"{report_path.stem}_CORRECTED{report_path.suffix}"

        # Process each sheet
        with pd.ExcelWriter(corrected_path, engine='openpyxl') as writer:
            for sheet_name in xl.sheet_names:
                # Skip non-client sheets
                if sheet_name in ['Summary', 'Logic_1', 'Logic_2']:
                    df = pd.read_excel(xl, sheet_name=sheet_name)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    continue

                print(f"Validating {sheet_name}...")
                df = pd.read_excel(xl, sheet_name=sheet_name)

                # Validate each row
                for idx, row in df.iterrows():
                    activity = row['Activity']

                    # Validate 5 year
                    if '5 Year Critical Failures' in df.columns:
                        result_5y = self.validate_activity(
                            client=sheet_name,
                            activity=activity,
                            critical_failures=row['5 Year Critical Failures'],
                            supporting_failures=row['5 Year Supporting Failures'],
                            final_status=row['5 Year Final Status'],
                            horizon="5 Year",
                        )
                        if result_5y['needs_correction']:
                            df.at[idx, '5 Year Critical Failures'] = result_5y['corrected_critical']
                            df.at[idx, '5 Year Supporting Failures'] = result_5y['corrected_supporting']
                            df.at[idx, '5 Year Final Status'] = result_5y['corrected_status']

                    # Validate 10 year
                    if '10 Year Critical Failures' in df.columns:
                        result_10y = self.validate_activity(
                            client=sheet_name,
                            activity=activity,
                            critical_failures=row['10 Year Critical Failures'],
                            supporting_failures=row['10 Year Supporting Failures'],
                            final_status=row['10 Year Final Status'],
                            horizon="10 Year",
                        )
                        if result_10y['needs_correction']:
                            df.at[idx, '10 Year Critical Failures'] = result_10y['corrected_critical']
                            df.at[idx, '10 Year Supporting Failures'] = result_10y['corrected_supporting']
                            df.at[idx, '10 Year Final Status'] = result_10y['corrected_status']

                # Write corrected sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"\n[OK] Corrected report written to: {corrected_path}")
        return corrected_path

    def generate_corrections_report(self, output_path: str | Path):
        """Generate a detailed corrections report."""
        output_path = Path(output_path)

        if not self.corrections:
            print("\n[OK] No corrections needed!")
            return

        print(f"\n=== GENERATING CORRECTIONS REPORT ===")
        print(f"Total corrections: {len(self.corrections)}\n")

        # Create DataFrame from corrections
        corrections_data = []
        for c in self.corrections:
            corrections_data.append({
                'Client': c.client,
                'Activity': c.activity,
                'Time Horizon': c.horizon,
                'Error Type': c.error_type,
                'Description': c.description,
                'Old Value': c.old_value,
                'New Value': c.new_value,
            })

        df = pd.DataFrame(corrections_data)

        # Write to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Corrections', index=False)

            # Summary sheet
            summary_data = {
                'Metric': [
                    'Total Corrections',
                    'Clients Affected',
                    'Activities Affected',
                    'Misclassification Errors',
                    'Status Calculation Errors',
                ],
                'Count': [
                    len(self.corrections),
                    len(set(c.client for c in self.corrections)),
                    len(set(f"{c.client}:{c.activity}" for c in self.corrections)),
                    len([c for c in self.corrections if c.error_type == 'Misclassification']),
                    len([c for c in self.corrections if c.error_type == 'Status Calculation']),
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        print(f"[OK] Corrections report written to: {output_path}")

        # Print summary to console
        print("\n=== CORRECTIONS SUMMARY ===")
        print(f"Total corrections: {len(self.corrections)}")
        print(f"Clients affected: {len(set(c.client for c in self.corrections))}")
        print(f"Misclassification errors: {len([c for c in self.corrections if c.error_type == 'Misclassification'])}")
        print(f"Status calculation errors: {len([c for c in self.corrections if c.error_type == 'Status Calculation'])}")


def main():
    """Main execution function."""
    base_path = Path(__file__).parent

    # Input files
    report_path = base_path / "out" / "persona_activity_report.xlsx"
    limits_path = base_path / "Exercise limits optimised.xlsx"
    classifications_path = base_path / "Exercise Threshold Classifications.xlsx"
    docx_path = base_path / "Activities Threshold Logic Matrix.docx"

    # Output files
    corrections_report_path = base_path / "out" / f"validation_corrections_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    # Create validation agent
    agent = ValidationAgent(
        limits_path=limits_path,
        classifications_path=classifications_path,
        docx_path=docx_path if docx_path.exists() else None,
    )

    # Validate and correct the report
    corrected_path = agent.validate_report(report_path)

    # Generate corrections report
    agent.generate_corrections_report(corrections_report_path)

    print("\n=== VALIDATION COMPLETE ===")
    print(f"Corrected report: {corrected_path}")
    print(f"Corrections log: {corrections_report_path}")


if __name__ == "__main__":
    main()
