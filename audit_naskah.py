#!/usr/bin/env python3
"""Wrapper CLI untuk audit_naskah."""
import os
import sys

scripts_dir = os.path.join(os.path.dirname(__file__), "scripts")
if scripts_dir not in sys.path:
    sys.path.insert(0, scripts_dir)

import audit_naskah

if __name__ == "__main__":
    sys.exit(audit_naskah.main())
