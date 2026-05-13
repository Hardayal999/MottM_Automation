"""
config.py
=========
Loads configuration from the .env file in the project root.

WHY a separate config module?
-----------------------------
Every other file in this package can `from visure.config import settings` and
get the same loaded values. We never read environment variables in random
places — one source of truth, one place to debug if something is wrong.
"""

from dataclasses import dataclass
from pathlib import Path
from dotenv import load_dotenv
import os
import sys


# Find the .env file. It should sit next to run.py, two levels up from this file.
#    visure_extractor/             <-- .env lives here
#    └── visure/
#        └── config.py             <-- we are here
PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"

# Load the .env file into os.environ. If the file doesn't exist, load_dotenv
# silently does nothing — we explicitly check below so the user gets a clear error.
load_dotenv(ENV_PATH)


@dataclass(frozen=True)
class Settings:
    """All runtime configuration in one immutable object.

    Using a dataclass (instead of bare module-level variables) makes it
    obvious what the script needs and makes it easy to swap in a fake
    Settings for testing later.
    """
    username: str
    password: str
    base_url: str
    output_dir: Path
    license_type: str


def _require(name: str) -> str:
    """Read an env var and fail loudly with a useful message if it's missing."""
    value = os.getenv(name)
    if not value:
        print(
            f"\nERROR: Required environment variable {name!r} is missing.\n"
            f"Check your .env file at: {ENV_PATH}\n"
            f"Copy .env.example to .env and fill in real values.\n",
            file=sys.stderr,
        )
        sys.exit(1)
    return value


def load_settings() -> Settings:
    """Build a Settings object from the loaded environment.

    Called once at startup. Anything that needs config imports `settings`
    from this module rather than reading os.environ directly.
    """
    if not ENV_PATH.exists():
        print(
            f"\nERROR: No .env file found at {ENV_PATH}\n"
            f"Copy .env.example to .env and fill in real values:\n"
            f"    cp .env.example .env\n",
            file=sys.stderr,
        )
        sys.exit(1)

    output_dir = Path(os.getenv("OUTPUT_DIR", "./output"))
    # Resolve relative paths against the project root, not the cwd.
    # That way the script works the same whether you double-click it or
    # run it from any terminal location.
    if not output_dir.is_absolute():
        output_dir = (PROJECT_ROOT / output_dir).resolve()

    # Strip any trailing slash from base_url so we can build URLs cleanly.
    base_url = _require("VISURE_BASE_URL").rstrip("/")

    return Settings(
        username=_require("VISURE_USERNAME"),
        password=_require("VISURE_PASSWORD"),
        base_url=base_url,
        output_dir=output_dir,
        license_type=os.getenv("LICENSE_TYPE", "AUTHORING"),
    )


# Load once on import. Every other module sees the same `settings` instance.
settings = load_settings()
