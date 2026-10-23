#!/usr/bin/env python3
"""Wrapper CLI untuk editor_soal_v1_85."""
import os
import sys

# Tambahkan folder scripts ke sys.path
scripts_dir = os.path.join(os.path.dirname(__file__), "scripts")
if scripts_dir not in sys.path:
    sys.path.insert(0, scripts_dir)

import editor_soal_v1_85

if __name__ == "__main__":
    # Jalankan modul utama editor_soal
    if hasattr(editor_soal_v1_85, "main"):
        sys.exit(editor_soal_v1_85.main())
    else:
        # Panggil via subprocess jika modul memakai __main__ guard langsung
        import subprocess
        script_path = os.path.join(scripts_dir, "editor_soal_v1_85.py")
        sys.exit(subprocess.run([sys.executable, script_path] + sys.argv[1:]).returncode)
