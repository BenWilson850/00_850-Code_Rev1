"""Monte Carlo Validation Workflow

Runs validation agent and independent verifier, provides clear pass/fail result.
"""

import subprocess
import sys
from pathlib import Path


def run_script(script_name, description):
    """Run a Python script and capture output."""
    print(f"\n{'='*60}")
    print(description)
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

    return result.returncode == 0, result.stdout


def main():
    """Main validation workflow."""
    print("\n" + "="*60)
    print("MONTE CARLO VALIDATION WORKFLOW")
    print("="*60)
    print("\nThis will:")
    print("  1. Run the validation agent to check outputs")
    print("  2. Run the independent verifier to confirm")
    print("  3. Give you a clear PASS/FAIL result")
    print()

    base_path = Path(__file__).parent

    # Check outputs exist
    decline_path = base_path / "outputs" / "MonteCarlo_Decline.xlsx"
    improvement_path = base_path / "outputs" / "MonteCarlo_Improvement.xlsx"

    if not decline_path.exists() or not improvement_path.exists():
        print("ERROR: Monte Carlo outputs not found!")
        print("Please run the Monte Carlo pipeline first:")
        print("  .\\run_monte_carlo.ps1 -Input \"Your_Client_File.xlsx\"")
        return False

    # Step 1: Run validation agent
    success1, output1 = run_script(
        "validate_monte_carlo.py",
        "STEP 1: Running Validation Agent"
    )

    if not success1:
        print("\n[FAIL] Validation agent encountered errors")
        return False

    # Check if validation passed
    validation_passed = "VALIDATION PASSED" in output1

    # Step 2: Run independent verifier
    success2, output2 = run_script(
        "verify_monte_carlo.py",
        "STEP 2: Running Independent Verifier"
    )

    if not success2:
        print("\n[FAIL] Independent verifier encountered errors")
        return False

    # Check if verifier confirmed
    verifier_passed = "NO ERRORS" in output2

    # Final result
    print(f"\n{'='*60}")
    print("FINAL RESULT")
    print('='*60)

    if validation_passed and verifier_passed:
        print("\n*** PASS ***")
        print("\nMonte Carlo outputs have been validated:")
        print("  - All directional changes are correct")
        print("  - All progressions are monotonic")
        print("  - All values are physiologically plausible")
        print("\nOutputs:")
        print("  - outputs/MonteCarlo_Decline.xlsx")
        print("  - outputs/MonteCarlo_Improvement.xlsx")
        return True
    else:
        print("\n*** FAIL ***")
        print("\nIssues found - check the validation report:")
        print("  - outputs/monte_carlo_validation_report.xlsx")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
