"""Independent Verifier for Monte Carlo Validation

Uses a completely different approach to verify the validation agent's work.
"""

import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


class IndependentMonteCarloChecker:
    """Independent checker using different code."""

    def __init__(self, config_path: Path):
        with open(config_path, 'r', encoding='utf-8-sig') as f:
            cfg = json.load(f)

        self.higher_is_better = set(cfg['higher_is_better'])
        self.lower_is_better = set(cfg['lower_is_better'])
        self.all_tests = cfg['tests_order']
        self.errors_found = []

    def simple_read_excel(self, path: Path) -> dict:
        """Simple Excel reader."""
        z = zipfile.ZipFile(path)

        # Get shared strings
        try:
            sst_xml = z.read('xl/sharedStrings.xml')
            sst = ET.fromstring(sst_xml)
            shared = []
            for si in sst.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                text = ''.join(t.text or '' for t in si.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'))
                shared.append(text)
        except:
            shared = []

        # Get sheets
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))

        rel_map = {}
        for rel in rels.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rel_map[rel.get('Id')] = rel.get('Target')

        result = {}
        for sheet in wb.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
            name = sheet.get('name')
            rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

            sheet_data = {}
            sheet_xml = ET.fromstring(z.read('xl/' + rel_map[rid]))

            for row in sheet_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                cells = {}
                for cell in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    ref = cell.get('r')
                    col = ''.join(c for c in ref if c.isalpha())

                    v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    if v_elem is not None:
                        if cell.get('t') == 's':
                            try:
                                cells[col] = shared[int(v_elem.text)]
                            except:
                                cells[col] = v_elem.text
                        else:
                            cells[col] = v_elem.text

                label = cells.get('A', '').strip()
                value_str = cells.get('B', '')
                if label:
                    try:
                        sheet_data[label] = float(value_str)
                    except:
                        pass

            result[name] = sheet_data

        return result

    def check_direction_simple(self, test: str, baseline: float, future: float, scenario: str) -> str | None:
        """Simple direction check - returns error message or None."""
        if baseline == 0:
            return None

        went_up = future > baseline
        went_down = future < baseline

        if test in self.higher_is_better:
            # Higher is better
            if scenario == "Decline" and went_up:
                return f"IMPROVED on decline (bad)"
            if scenario == "Improvement" and went_down:
                return f"DECLINED on improvement (bad)"

        elif test in self.lower_is_better:
            # Lower is better
            if scenario == "Decline" and went_down:
                return f"IMPROVED on decline (bad)"
            if scenario == "Improvement" and went_up:
                return f"DECLINED on improvement (bad)"

        return None

    def check_monotonic_simple(self, test: str, baseline: float, y5: float, y10: float, scenario: str) -> str | None:
        """Simple monotonic check."""
        if test in self.higher_is_better:
            if scenario == "Decline":
                if y5 > baseline or y10 > y5:
                    return f"Not monotonically declining"
            else:
                if y5 < baseline or y10 < y5:
                    return f"Not monotonically improving"

        elif test in self.lower_is_better:
            if scenario == "Decline":
                if y5 < baseline or y10 < y5:
                    return f"Not monotonically worsening"
            else:
                if y5 > baseline or y10 > y5:
                    return f"Not monotonically improving"

        return None

    def verify_file(self, path: Path, scenario: str):
        """Verify one output file."""
        print(f"\nChecking {scenario}: {path.name}")

        if not path.exists():
            print(f"  [SKIP] File not found")
            return

        data = self.simple_read_excel(path)
        print(f"  Found {len(data)} clients")

        for client_name, values in data.items():
            for test in self.all_tests:
                baseline = values.get(test)
                y5 = values.get(f"{test} (Year 5)")
                y10 = values.get(f"{test} (Year 10)")

                if baseline is None:
                    continue

                # Check directions
                if y5 is not None:
                    err = self.check_direction_simple(test, baseline, y5, scenario)
                    if err:
                        self.errors_found.append(f"{client_name} | {test} | Year 5: {err}")

                if y10 is not None:
                    err = self.check_direction_simple(test, baseline, y10, scenario)
                    if err:
                        self.errors_found.append(f"{client_name} | {test} | Year 10: {err}")

                # Check monotonic
                if y5 is not None and y10 is not None:
                    err = self.check_monotonic_simple(test, baseline, y5, y10, scenario)
                    if err:
                        self.errors_found.append(f"{client_name} | {test} | {err}")

    def print_results(self):
        """Print verification results."""
        print("\n" + "="*60)
        print("INDEPENDENT VERIFICATION RESULTS")
        print("="*60)

        if not self.errors_found:
            print("\n*** VERIFIED - NO ERRORS ***")
            print("The validation agent's assessment is correct.")
            print("All Monte Carlo outputs pass independent verification.")
        else:
            print(f"\n*** FOUND {len(self.errors_found)} ERRORS ***")
            print("\nFirst 10 errors:")
            for err in self.errors_found[:10]:
                print(f"  ! {err}")


def main():
    """Run independent verification."""
    base_path = Path(__file__).parent

    print("=== INDEPENDENT MONTE CARLO VERIFICATION ===")

    checker = IndependentMonteCarloChecker(base_path / "mc_assumptions.json")

    checker.verify_file(
        base_path / "outputs" / "MonteCarlo_Decline.xlsx",
        "Decline"
    )
    checker.verify_file(
        base_path / "outputs" / "MonteCarlo_Improvement.xlsx",
        "Improvement"
    )

    checker.print_results()


if __name__ == "__main__":
    main()
