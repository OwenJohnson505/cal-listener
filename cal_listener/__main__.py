"""Run with: python -m cal_listener

PyInstaller bundles this as the entry point, but at runtime the bundled
script runs as ``__main__`` (no parent package), so we must import the
daemon by absolute path — relative imports raise
``ImportError: attempted relative import with no known parent package``.
"""
from cal_listener.daemon import main

if __name__ == "__main__":
    main()
