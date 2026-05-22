"""Spec §12 acceptance-number tests.

Placeholder in v0.1; populated in v1.0 once calcs.py is implemented.
Each test asserts the Python mirror of an Excel formula produces the
expected acceptance number to whole-dollar precision.
"""

from div296 import __version__


def test_package_importable():
    assert __version__ == "0.1.0"
