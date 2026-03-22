"""
check_packages.py — Verify required Python packages are installed before running audit scripts.
Returns exit code 0 if all OK, 1 if any missing.
"""
import sys
import importlib

REQUIRED = [
    ("pandas",   "pandas>=1.3",   "pip install pandas"),
    ("openpyxl", "openpyxl>=3.0", "pip install openpyxl"),
    ("xlrd",     "xlrd>=2.0",     "pip install xlrd"),
    ("numpy",    "numpy>=1.21",   "pip install numpy"),
]

missing = []
for module, friendly, install_cmd in REQUIRED:
    try:
        importlib.import_module(module)
    except ImportError:
        missing.append((friendly, install_cmd))

if missing:
    print("\n❌ Missing required Python packages:\n")
    for friendly, cmd in missing:
        print(f"   {friendly:20s}  →  {cmd}")
    print("\nPlease install them and retry.\n")
    sys.exit(1)
else:
    print("✅ All required packages are installed (pandas, openpyxl, xlrd, numpy)")
    sys.exit(0)
