#!/usr/bin/env python3
"""
Run the Streamlit application
Ejecutar la aplicacion Streamlit
"""

import subprocess
import sys
from pathlib import Path

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))


def main():
    """Launch the Streamlit app / Lanzar la aplicacion Streamlit"""
    app_path = PROJECT_ROOT / "ui" / "streamlit_app" / "app.py"

    if not app_path.exists():
        print(f"Error: App not found at {app_path}")
        sys.exit(1)

    print("Starting Carta de Manifestacion Generator...")
    print(f"App path: {app_path}")

    try:
        subprocess.run([
            sys.executable, "-m", "streamlit", "run",
            str(app_path),
            "--server.headless=true"
        ], check=True)
    except KeyboardInterrupt:
        print("\nApp stopped.")
    except subprocess.CalledProcessError as e:
        print(f"Error running app: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
