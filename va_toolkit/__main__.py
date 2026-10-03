"""Allow ``python -m va_toolkit``."""

import sys

from .cli import main

sys.exit(main())
