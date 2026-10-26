"""Allow `python -m tidytable`."""
import sys

from .cli import main

sys.exit(main())
