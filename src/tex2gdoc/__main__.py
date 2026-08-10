"""Entry point for `python -m tex2gdoc`.

The module can no longer be run as a loose file, because `self_test()` imports
the calibration registry as a sibling. Running it as a package is what makes
that import resolve.
"""

import sys

from .tex2gdoc import main

if __name__ == "__main__":
    sys.exit(main())
