#!/usr/bin/env python3
"""
CLI script for document generation
Script CLI para generacion de documentos
"""

import argparse
import json
import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from modules.generate import generate
from modules.plugin_loader import load_plugin, list_available_plugins


def main():
    """Main CLI entry point / Punto de entrada CLI principal"""
    parser = argparse.ArgumentParser(
        description="Generate Carta de Manifestacion documents from CLI"
    )

    parser.add_argument(
        "--plugin",
        "-p",
        default="carta_manifestacion",
        help="Plugin ID to use (default: carta_manifestacion)"
    )

    parser.add_argument(
        "--data",
        "-d",
        required=True,
        help="Path to JSON file with input data"
    )

    parser.add_argument(
        "--output",
        "-o",
        default="output",
        help="Output directory (default: output)"
    )

    parser.add_argument(
        "--template",
        "-t",
        help="Custom template path (optional)"
    )

    parser.add_argument(
        "--no-validate",
        action="store_true",
        help="Skip input validation"
    )

    parser.add_argument(
        "--list-plugins",
        action="store_true",
        help="List available plugins and exit"
    )

    args = parser.parse_args()

    # List plugins if requested
    if args.list_plugins:
        plugins = list_available_plugins()
        print("Available plugins:")
        for p in plugins:
            print(f"  - {p}")
        return 0

    # Load input data
    data_path = Path(args.data)
    if not data_path.exists():
        print(f"Error: Data file not found: {data_path}")
        return 1

    try:
        with open(data_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        return 1

    # Setup output directory
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Template path
    template_path = Path(args.template) if args.template else None

    # Generate document
    print(f"Generating document with plugin: {args.plugin}")
    print(f"Input data: {data_path}")
    print(f"Output directory: {output_dir}")

    result = generate(
        plugin_id=args.plugin,
        data=data,
        output_dir=output_dir,
        template_path=template_path,
        should_validate=not args.no_validate
    )

    if result.success:
        print(f"\nSuccess! Document generated: {result.output_path}")
        print(f"Trace ID: {result.trace_id}")
        print(f"Duration: {result.duration_ms}ms")
        return 0
    else:
        print(f"\nError: {result.error}")
        if result.validation_errors:
            print("Validation errors:")
            for err in result.validation_errors:
                print(f"  - {err}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
