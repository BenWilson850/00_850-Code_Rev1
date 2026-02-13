"""Independent Verification Tool

This script independently checks the validation agent's corrections to ensure accuracy.
Uses a completely different code path to verify the same logic.
"""

import pandas as pd
from pathlib import Path
import re


class IndependentChecker:
    """Independent checker that verifies validation agent corrections."""

    def __init__(self, limits_path, classifications_path):
        self.limits_path = Path(limits_path)
        self.classifications_path = Path(classifications_path)
        self._load_data()

    def _load_data(self):
        """Load reference matrices."""
        print("Loading reference data...")

        # Load classifications
        self.class_df = pd.read_excel(
            self.classifications_path,
            sheet_name='Activities Thresholds_Rev2',
            header=0
        )
        self.class_df['Activity'] = self.class_df['Activity'].astype(str).str.strip()

        # Create lookup: for each activity, which tests are Critical vs Supporting
        self.activity_classifications = {}

        for idx, row in self.class_df.iterrows():
            activity = row['Activity']
            if pd.isna(activity) or activity == 'nan':
                continue

            critical_tests = []
            supporting_tests = []

            for col in self.class_df.columns:
                if col in ['Activity', 'Unnamed: 0', 'Key References', 'Evidence Quality']:
                    continue

                val = row[col]
                if pd.isna(val):
                    continue

                val_str = str(val).lower()
                # Normalize column name
                col_normalized = ' '.join(col.split())

                if 'critical' in val_str:
                    critical_tests.append(col_normalized)
                elif 'supporting' in val_str:
                    supporting_tests.append(col_normalized)

            self.activity_classifications[activity] = {
                'critical': critical_tests,
                'supporting': supporting_tests
            }

        print(f"Loaded classifications for {len(self.activity_classifications)} activities")

    def _parse_test_from_failure(self, failure_text):
        """Extract test name from 'VO2 Max (ml/kg/min) (RED)'."""
        # Remove zone at end
        text = re.sub(r'\s*\((RED|YELLOW|GREEN|MISSING)\)\s*$', '', failure_text)
        # Normalize whitespace
        return ' '.join(text.split())

    def _normalize_name(self, name):
        """Normalize test name for comparison."""
        return ' '.join(str(name).lower().split())

    def _test_matches(self, test_from_report, test_from_matrix):
        """Check if two test names refer to the same test."""
        norm_report = self._normalize_name(test_from_report)
        norm_matrix = self._normalize_name(test_from_matrix)

        # Exact match
        if norm_report == norm_matrix:
            return True

        # One contains the other
        if norm_report in norm_matrix or norm_matrix in norm_report:
            return True

        return False

    def verify_activity_row(self, activity, critical_failures_str, supporting_failures_str, final_status):
        """Verify one activity row.

        Returns:
            dict with verification results
        """
        issues = []

        # Get expected classifications for this activity
        if activity not in self.activity_classifications:
            return {
                'verified': None,
                'issues': [f"Activity '{activity}' not found in classifications matrix"],
                'status_correct': None,
            }

        expected = self.activity_classifications[activity]
        expected_critical = expected['critical']
        expected_supporting = expected['supporting']

        # Parse what's in the report
        critical_failures = []
        supporting_failures = []

        if pd.notna(critical_failures_str) and str(critical_failures_str).strip():
            for part in str(critical_failures_str).split(','):
                test_name = self._parse_test_from_failure(part.strip())
                if test_name:
                    critical_failures.append(test_name)

        if pd.notna(supporting_failures_str) and str(supporting_failures_str).strip():
            for part in str(supporting_failures_str).split(','):
                test_name = self._parse_test_from_failure(part.strip())
                if test_name:
                    supporting_failures.append(test_name)

        # Check each test in critical failures
        for test in critical_failures:
            is_critical = any(self._test_matches(test, exp) for exp in expected_critical)
            is_supporting = any(self._test_matches(test, exp) for exp in expected_supporting)

            if is_supporting:
                issues.append(f"MISCLASSIFICATION: '{test}' listed as Critical but should be Supporting")
            elif not is_critical:
                # Not in either list - might be OK if it's a valid test just not classified
                pass

        # Check each test in supporting failures
        for test in supporting_failures:
            is_critical = any(self._test_matches(test, exp) for exp in expected_critical)
            is_supporting = any(self._test_matches(test, exp) for exp in expected_supporting)

            if is_critical:
                issues.append(f"MISCLASSIFICATION: '{test}' listed as Supporting but should be Critical")
            elif not is_supporting:
                # Not in either list
                pass

        # Verify status calculation
        # Extract zones from the failures
        critical_zones = []
        supporting_red_count = 0

        if pd.notna(critical_failures_str):
            critical_zones = re.findall(r'\((RED|YELLOW|MISSING)\)', str(critical_failures_str))

        if pd.notna(supporting_failures_str):
            supporting_red_count = len(re.findall(r'\(RED\)', str(supporting_failures_str)))

        # Apply threshold logic
        expected_status = self._calculate_expected_status(critical_zones, supporting_red_count)

        status_correct = (final_status == expected_status)
        if not status_correct:
            issues.append(f"STATUS ERROR: Should be {expected_status}, got {final_status}")

        return {
            'verified': len(issues) == 0,
            'issues': issues,
            'status_correct': status_correct,
            'expected_status': expected_status,
        }

    def _calculate_expected_status(self, critical_zones, supporting_red_count):
        """Calculate expected status based on threshold logic.

        Rules:
        - Any critical RED -> RED
        - >3 supporting RED -> RED
        - Any critical YELLOW/MISSING -> YELLOW
        - Exactly 2 supporting RED -> YELLOW
        - All critical GREEN and <2 supporting RED -> GREEN
        - Otherwise -> YELLOW (safety fallback)
        """
        # Check for critical RED
        if 'RED' in critical_zones:
            return 'RED'

        # Check for too many supporting RED
        if supporting_red_count > 3:
            return 'RED'

        # Check for critical YELLOW or MISSING
        if 'YELLOW' in critical_zones or 'MISSING' in critical_zones:
            return 'YELLOW'

        # Check for exactly 2 supporting RED
        if supporting_red_count == 2:
            return 'YELLOW'

        # Check if all critical are GREEN
        if critical_zones and all(z == 'GREEN' for z in critical_zones) and supporting_red_count < 2:
            return 'GREEN'

        # No critical zones at all - safety fallback
        if not critical_zones:
            return 'YELLOW'

        # Other cases - fallback
        return 'YELLOW'

    def verify_report(self, corrected_report_path):
        """Verify the entire corrected report."""
        print(f"\n=== VERIFYING CORRECTED REPORT ===")
        print(f"Report: {corrected_report_path}\n")

        corrected_path = Path(corrected_report_path)
        if not corrected_path.exists():
            print(f"ERROR: Report not found: {corrected_path}")
            return

        xl = pd.ExcelFile(corrected_path)

        total_rows = 0
        total_issues = 0
        clients_with_issues = []

        for sheet_name in xl.sheet_names:
            if sheet_name in ['Summary', 'Logic_1', 'Logic_2']:
                continue

            df = pd.read_excel(xl, sheet_name=sheet_name)
            sheet_issues = 0

            for idx, row in df.iterrows():
                activity = row['Activity']

                # Check 5 year
                if '5 Year Final Status' in df.columns:
                    result = self.verify_activity_row(
                        activity,
                        row['5 Year Critical Failures'],
                        row['5 Year Supporting Failures'],
                        row['5 Year Final Status']
                    )
                    total_rows += 1
                    if result['issues']:
                        total_issues += len(result['issues'])
                        sheet_issues += len(result['issues'])
                        print(f"{sheet_name} - {activity} (5 Year):")
                        for issue in result['issues']:
                            print(f"  ! {issue}")

                # Check 10 year
                if '10 Year Final Status' in df.columns:
                    result = self.verify_activity_row(
                        activity,
                        row['10 Year Critical Failures'],
                        row['10 Year Supporting Failures'],
                        row['10 Year Final Status']
                    )
                    total_rows += 1
                    if result['issues']:
                        total_issues += len(result['issues'])
                        sheet_issues += len(result['issues'])
                        print(f"{sheet_name} - {activity} (10 Year):")
                        for issue in result['issues']:
                            print(f"  ! {issue}")

            if sheet_issues > 0:
                clients_with_issues.append(sheet_name)

        print(f"\n=== VERIFICATION SUMMARY ===")
        print(f"Total rows checked: {total_rows}")
        print(f"Total issues found: {total_issues}")
        print(f"Clients with issues: {len(clients_with_issues)}")

        if total_issues == 0:
            print("\n[OK] VALIDATION AGENT VERIFIED - All corrections are accurate!")
        else:
            print(f"\n[WARNING] Found {total_issues} potential issues")
            print(f"Affected clients: {', '.join(clients_with_issues)}")


def main():
    """Main verification function."""
    base_path = Path(__file__).parent

    # Inputs
    corrected_report = base_path / "out" / "persona_activity_report_CORRECTED.xlsx"
    limits_path = base_path / "Exercise limits optimised.xlsx"
    classifications_path = base_path / "Exercise Threshold Classifications.xlsx"

    if not corrected_report.exists():
        print(f"ERROR: Corrected report not found: {corrected_report}")
        print("Run validate_output.py first to generate the corrected report.")
        return

    # Create checker
    checker = IndependentChecker(limits_path, classifications_path)

    # Verify the corrected report
    checker.verify_report(corrected_report)


if __name__ == "__main__":
    main()
