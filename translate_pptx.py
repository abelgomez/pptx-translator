#!/usr/bin/env python
"""Script entry point: ``python translate_pptx.py input.pptx -t en``."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from pptx_translator.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
