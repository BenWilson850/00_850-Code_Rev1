"""Script to detect and fix activity name mismatches between Excel files.

This script:
1. Reads both Excel files (limits and classifications)
2. Identifies activity name mismatches
3. Standardizes names using the limits file as source of truth
4. Creates backups and updates the classifications file
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
from difflib import SequenceMatcher
import shutil

def similarity(a: str, b: str) -> float:
    """Calculate similarity ratio between two strings (0-1)."""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def find_best_match(activity: str, candidates: list[str], threshold: float = 0.6) -> tuple[str | None, float]:
    """Find the best matching activity name from candidates."""
    best_match = None
    best_score = 0.0

    for candidate in candidates:
        score = similarity(activity, candidate)
        if score > best_score:
            best_score = score
            best_match = candidate

    if best_score >= threshold:
        return best_match, best_score
    return None, best_score

def detect_mismatches(limits_path: str, classifications_path: str):
    """Detect activity name mismatches between the two files."""
    print(f"Reading {limits_path}...")
    limits_df = pd.read_excel(limits_path, sheet_name='Activities Thresholds_Rev2', header=0)
    # Filter out NaN values before converting to string
    limits_activities = [str(a).strip() for a in limits_df['Activity'].tolist()
                        if pd.notna(a) and str(a).strip() != '' and str(a).strip() != 'nan']

    print(f"Reading {classifications_path}...")
    class_df = pd.read_excel(classifications_path, sheet_name='Activities Thresholds_Rev2', header=0)
    # Filter out NaN values before converting to string
    class_activities = [str(a).strip() for a in class_df['Activity'].tolist()
                       if pd.notna(a) and str(a).strip() != '' and str(a).strip() != 'nan']

    print(f"\nFound {len(limits_activities)} activities in limits file")
    print(f"Found {len(class_activities)} activities in classifications file\n")

    mismatches = []

    # Check each activity from limits file
    for limit_activity in limits_activities:
        if limit_activity not in class_activities:
            # Try to find a similar name
            match, score = find_best_match(limit_activity, class_activities)
            mismatches.append({
                'limits_name': limit_activity,
                'class_name': match,
                'similarity': score,
                'status': 'fuzzy_match' if match else 'missing'
            })

    # Check for activities in classifications that don't exist in limits
    for class_activity in class_activities:
        if class_activity not in limits_activities:
            match, score = find_best_match(class_activity, limits_activities)
            if not any(m['class_name'] == class_activity for m in mismatches):
                mismatches.append({
                    'limits_name': match,
                    'class_name': class_activity,
                    'similarity': score,
                    'status': 'extra_in_class'
                })

    return mismatches, limits_df, class_df

def fix_mismatches(mismatches: list[dict], limits_df: pd.DataFrame, class_df: pd.DataFrame,
                   classifications_path: str):
    """Fix the mismatches by standardizing to limits file naming."""
    if not mismatches:
        print("No mismatches found! Files are already consistent.")
        return

    print("\n=== DETECTED MISMATCHES ===\n")
    for i, m in enumerate(mismatches, 1):
        print(f"{i}. Limits: '{m['limits_name']}' <-> Classifications: '{m['class_name']}'")
        print(f"   Similarity: {m['similarity']:.2%}, Status: {m['status']}\n")

    # Create backup
    backup_path = Path(classifications_path).with_suffix('.backup.xlsx')
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = Path(classifications_path).parent / f"{Path(classifications_path).stem}_backup_{timestamp}.xlsx"

    print(f"Creating backup: {backup_path}")
    shutil.copy2(classifications_path, backup_path)

    # Apply fixes to classifications file
    print("\nApplying fixes to classifications file...")

    # Create a mapping of old names to new names
    rename_mapping = {}
    for m in mismatches:
        if m['class_name'] and m['limits_name']:
            # Use limits name as the standard
            if m['class_name'] != m['limits_name']:
                rename_mapping[m['class_name']] = m['limits_name']

    # Apply the renaming
    class_df['Activity'] = class_df['Activity'].replace(rename_mapping)

    print(f"Renamed {len(rename_mapping)} activities:")
    for old_name, new_name in rename_mapping.items():
        print(f"  '{old_name}' -> '{new_name}'")

    # Write the updated file
    print(f"\nWriting updated file to {classifications_path}")
    with pd.ExcelWriter(classifications_path, engine='openpyxl') as writer:
        class_df.to_excel(writer, sheet_name='Activities Thresholds_Rev2', index=False)

    print("\n[OK] Classification file updated successfully!")
    print(f"[OK] Backup saved to: {backup_path}")

def main():
    """Main execution function."""
    base_path = Path(__file__).parent
    limits_path = base_path / "Exercise limits optimised.xlsx"
    classifications_path = base_path / "Exercise Threshold Classifications.xlsx"

    if not limits_path.exists():
        print(f"Error: {limits_path} not found")
        return

    if not classifications_path.exists():
        print(f"Error: {classifications_path} not found")
        return

    print("=== Activity Name Standardization Script ===\n")

    mismatches, limits_df, class_df = detect_mismatches(str(limits_path), str(classifications_path))

    if mismatches:
        fix_mismatches(mismatches, limits_df, class_df, str(classifications_path))
    else:
        print("\n[OK] No mismatches detected. Files are already consistent!")

if __name__ == "__main__":
    main()
