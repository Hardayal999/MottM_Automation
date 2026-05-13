"""
run.py
======
Entry point. Double-click this (after setting up .env) or run from a terminal:

    python run.py

All the actual logic lives in the visure/ package. Keeping this file tiny
means the moving parts stay testable and importable from elsewhere.
"""

import sys
import traceback

from visure.extractor import run_interactive_extraction


def main() -> int:
    try:
        run_interactive_extraction()
        return 0
    except KeyboardInterrupt:
        print("\n\nInterrupted by user.")
        return 130   # standard exit code for Ctrl-C
    except Exception:
        # Print the full traceback so the user can copy/paste it back to us.
        print("\n--- Unexpected error ---", file=sys.stderr)
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
