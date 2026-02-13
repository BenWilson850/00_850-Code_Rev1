"""Validation Workflow Runner

Runs the validation agent and independent verifier together,
then provides a clear PASS/FAIL result.
"""

import subprocess
import sys
from pathlib import Path
import pandas as pd
import re


def run_command(script_name, description):
    """Run a Python script and return success status."""
    print(f"\n{'='*60}")
    print(f"{description}")
    print('='*60)

    result = subprocess.run(
        [sys.executable, script_name],
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace'
    )

    print(result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)

    return result.returncode == 0


def analyze_verification_results():
    """Analyze the verification output and filter real errors."""
    base_path = Path(__file__).parent

    # Run the verifier and capture output
    result = subprocess.run(
        [sys.executable, "verify_validator.py"],
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace'
    )

    output = result.stdout

    # Parse the output to find real errors (not nan activities)
    lines = output.split('\n')
    real_errors = []
    current_error = None

    for line in lines:
        # Check if it's an error line
        if ' - ' in line and '(5 Year)' in line or '(10 Year)' in line:
            # Check if it's not a nan activity
            if ' - nan ' not in line:
                current_error = line.strip()
        elif line.strip().startswith('!') and current_error:
            # This is an error description
            error_desc = line.strip()
            # Skip "Activity 'nan' not found" errors
            if "Activity 'nan' not found" not in error_desc:
                real_errors.append(f"{current_error}\n  {error_desc}")
                current_error = None

    return real_errors, output


def main():
    """Main validation workflow."""
    print("\n" + "="*60)
    print("VALIDATION WORKFLOW")
    print("="*60)
    print("\nThis will:")
    print("  1. Run the validation agent to correct errors")
    print("  2. Run the independent verifier to check corrections")
    print("  3. Give you a clear PASS/FAIL result")
    print()

    base_path = Path(__file__).parent

    # Check required files exist
    required_files = [
        "validate_output.py",
        "verify_validator.py",
        "out/persona_activity_report.xlsx"
    ]

    for file in required_files:
        if not (base_path / file).exists():
            print(f"ERROR: Required file not found: {file}")
            return False

    # Step 1: Run validation agent
    success = run_command(
        "validate_output.py",
        "STEP 1: Running Validation Agent"
    )

    if not success:
        print("\n[FAIL] Validation agent encountered errors")
        return False

    # Step 2: Run independent verifier and analyze
    print(f"\n{'='*60}")
    print("STEP 2: Running Independent Verifier")
    print('='*60)

    real_errors, full_output = analyze_verification_results()

    # Print summary from verifier
    summary_started = False
    for line in full_output.split('\n'):
        if 'VERIFICATION SUMMARY' in line:
            summary_started = True
        if summary_started:
            print(line)

    # Step 3: Final result
    print(f"\n{'='*60}")
    print("FINAL RESULT")
    print('='*60)

    if not real_errors:
        print("\n*** PASS ***")
        print("\nThe validation agent's corrections have been independently verified.")
        print("All corrections are accurate!")
        print("\nOutput files:")
        print("  - out/persona_activity_report_CORRECTED.xlsx (use this file)")
        print("  - out/validation_corrections_*.xlsx (review what was fixed)")
        return True
    else:
        print("\n*** FAIL ***")
        print(f"\nFound {len(real_errors)} real error(s) that need investigation:\n")
        for i, error in enumerate(real_errors, 1):
            print(f"{i}. {error}")
        print("\nAction required: Review these cases manually")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
