"""Command-line interface for BFA calculation pipeline."""

from __future__ import annotations
import argparse
from pathlib import Path
import yaml

from .io.client_reader import read_client_workbook
from .io.normative_reader import read_normative_database
from .core.test_scoring import score_all_tests
from .core.pillar_scoring import calculate_pillar_functional_ages
from .core.bfa_calculation import calculate_bfa
from .core.healthspan_index import calculate_healthspan_index, categorize_healthspan_index
from .report.output_writer import create_output_row, write_results_csv, write_results_excel


def load_config(config_dir: Path) -> dict:
    """Load configuration files."""
    with open(config_dir / 'weights.yaml') as f:
        weights = yaml.safe_load(f)

    with open(config_dir / 'healthspan_categories.yaml') as f:
        healthspan_config = yaml.safe_load(f)

    return {
        'pillar_weights': weights['pillar_weights'],
        'subtest_weights': weights['subtest_weights'],
        'healthspan_index': healthspan_config['healthspan_index'],
        'categories': healthspan_config['categories']
    }


def process_client(
    client_tests,
    norm_data,
    config
):
    """Process a single client through the BFA pipeline."""
    # Score individual tests
    individual_scores = score_all_tests(
        client_tests,
        norm_data,
        config['subtest_weights']['metabolic']
    )

    # Calculate pillar scores
    pillars = calculate_pillar_functional_ages(
        individual_scores,
        client_tests,
        config['subtest_weights']
    )

    # Calculate BFA
    bfa = calculate_bfa(pillars, config['pillar_weights'])

    # Calculate Healthspan Index
    healthspan_index, _ = calculate_healthspan_index(
        client_tests.age,
        bfa,
        config['healthspan_index']
    )

    # Determine category
    healthspan_category = categorize_healthspan_index(
        healthspan_index,
        config['categories']
    )

    return pillars, bfa, healthspan_index, healthspan_category


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description='Calculate Biological Functional Age and Healthspan Index from client test data.'
    )
    parser.add_argument(
        'client_file',
        type=str,
        help='Path to client test data Excel workbook'
    )
    parser.add_argument(
        '-n', '--normative-db',
        type=str,
        default='Normative Database.xlsx',
        help='Path to normative database (default: Normative Database.xlsx)'
    )
    parser.add_argument(
        '-c', '--config-dir',
        type=str,
        default='configs',
        help='Path to configuration directory (default: configs)'
    )
    parser.add_argument(
        '-o', '--output',
        type=str,
        default='output/bfa_results.csv',
        help='Output file path (.csv or .xlsx) (default: output/bfa_results.csv)'
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Verbose output'
    )

    args = parser.parse_args()

    # Resolve paths
    client_file = Path(args.client_file)
    normative_db = Path(args.normative_db)
    config_dir = Path(args.config_dir)
    output_path = Path(args.output)

    # Validate inputs
    if not client_file.exists():
        print(f"Error: Client file not found: {client_file}")
        return 1

    if not normative_db.exists():
        print(f"Error: Normative database not found: {normative_db}")
        return 1

    if not config_dir.exists():
        print(f"Error: Config directory not found: {config_dir}")
        return 1

    # Load configuration
    if args.verbose:
        print("Loading configuration...")
    config = load_config(config_dir)

    # Load normative database
    if args.verbose:
        print(f"Loading normative database from: {normative_db}")
    norm_data = read_normative_database(normative_db)

    # Read client data
    if args.verbose:
        print(f"Reading client data from: {client_file}")
    clients = read_client_workbook(client_file)
    print(f"Found {len(clients)} client(s)")

    # Process each client
    results = []
    for i, client in enumerate(clients, 1):
        if args.verbose:
            print(f"\nProcessing client {i}/{len(clients)}: {client.name}")

        try:
            pillars, bfa, healthspan_index, healthspan_category = process_client(
                client,
                norm_data,
                config
            )

            # Create output row
            result_row = create_output_row(
                client,
                pillars,
                bfa,
                healthspan_index,
                healthspan_category
            )
            results.append(result_row)

            if args.verbose:
                if bfa is not None:
                    print(f"  BFA: {bfa:.1f} (Chronological: {client.age})")
                    print(f"  Healthspan Index: {healthspan_index:.0f} ({healthspan_category})")
                else:
                    print(f"  BFA: INCOMPLETE (missing test data)")

        except Exception as e:
            print(f"Error processing client {client.name}: {e}")
            if args.verbose:
                import traceback
                traceback.print_exc()
            continue

    # Write results
    if not results:
        print("Error: No results to write")
        return 1

    if output_path.suffix == '.xlsx':
        write_results_excel(results, output_path)
    else:
        write_results_csv(results, output_path)

    return 0


if __name__ == '__main__':
    exit(main())
