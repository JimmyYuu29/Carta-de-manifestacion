#!/usr/bin/env python3
"""
CLI script for plugin validation
Script CLI para validacion de plugins
"""

import argparse
import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from modules.plugin_loader import load_plugin, list_available_plugins


def validate_plugin(plugin_id: str) -> bool:
    """
    Validate a plugin configuration
    Validar configuracion de un plugin

    Args:
        plugin_id: ID of the plugin to validate

    Returns:
        True if valid, False otherwise
    """
    print(f"\n{'='*60}")
    print(f"Validating plugin: {plugin_id}")
    print(f"{'='*60}")

    errors = []
    warnings = []

    try:
        plugin = load_plugin(plugin_id)

        # Check manifest
        print("\n[Manifest]")
        manifest = plugin.manifest
        if not manifest:
            errors.append("manifest.yaml is empty or missing")
        else:
            required_manifest = ['plugin_id', 'version', 'name']
            for field in required_manifest:
                if field not in manifest:
                    errors.append(f"Manifest missing required field: {field}")
                else:
                    print(f"  {field}: {manifest[field]}")

        # Check fields
        print("\n[Fields]")
        fields = plugin.fields.get("fields", {})
        if not fields:
            warnings.append("No fields defined in fields.yaml")
        else:
            print(f"  Total fields: {len(fields)}")

            # Check required fields have labels
            for name, spec in fields.items():
                if not spec.get("label"):
                    warnings.append(f"Field '{name}' has no label")

        # Check logic
        print("\n[Logic/Rules]")
        rules = plugin.logic.get("rules", {})
        print(f"  Total rules: {len(rules)}")

        # Check template
        print("\n[Template]")
        template_path = plugin.get_template_path()
        if template_path.exists():
            print(f"  Path: {template_path}")
            print(f"  Exists: Yes")
        else:
            errors.append(f"Template not found at: {template_path}")

        # Check config
        print("\n[Config]")
        config = plugin.config
        sections = config.get("sections", [])
        print(f"  Sections: {len(sections)}")
        oficinas = config.get("oficinas", {})
        print(f"  Oficinas: {len(oficinas)}")

        # Check formatting
        print("\n[Formatting]")
        formatting = plugin.formatting
        field_formats = formatting.get("fields", {})
        print(f"  Field formats: {len(field_formats)}")
        colors = formatting.get("colors", {})
        print(f"  Color mappings: {len(colors)}")

    except Exception as e:
        errors.append(f"Error loading plugin: {e}")

    # Print results
    print(f"\n{'='*60}")
    print("Validation Results")
    print(f"{'='*60}")

    if errors:
        print(f"\nERRORS ({len(errors)}):")
        for err in errors:
            print(f"  [X] {err}")

    if warnings:
        print(f"\nWARNINGS ({len(warnings)}):")
        for warn in warnings:
            print(f"  [!] {warn}")

    if not errors and not warnings:
        print("\n[OK] Plugin is valid with no issues!")

    elif not errors:
        print(f"\n[OK] Plugin is valid with {len(warnings)} warning(s)")

    else:
        print(f"\n[FAIL] Plugin has {len(errors)} error(s)")
        return False

    return True


def main():
    """Main CLI entry point / Punto de entrada CLI principal"""
    parser = argparse.ArgumentParser(
        description="Validate plugin configurations"
    )

    parser.add_argument(
        "--plugin",
        "-p",
        help="Plugin ID to validate (validates all if not specified)"
    )

    parser.add_argument(
        "--list",
        "-l",
        action="store_true",
        help="List available plugins"
    )

    args = parser.parse_args()

    # List plugins if requested
    if args.list:
        plugins = list_available_plugins()
        print("Available plugins:")
        for p in plugins:
            print(f"  - {p}")
        return 0

    # Validate specific plugin or all
    if args.plugin:
        success = validate_plugin(args.plugin)
        return 0 if success else 1
    else:
        plugins = list_available_plugins()
        if not plugins:
            print("No plugins found!")
            return 1

        all_valid = True
        for plugin_id in plugins:
            if not validate_plugin(plugin_id):
                all_valid = False

        return 0 if all_valid else 1


if __name__ == "__main__":
    sys.exit(main())
