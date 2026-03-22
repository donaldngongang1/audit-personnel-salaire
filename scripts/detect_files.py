"""
detect_files.py — Scan a directory for required audit source files.
Outputs a JSON summary to stdout. Writes/updates .audit-session.json.
Also stores accounting_plan (SYSCOHADA or PCG) if passed via --plan.

Usage: python detect_files.py [--dir /path/to/audit/folder] [--plan SYSCOHADA|PCG]
"""
import argparse, glob, json, os, sys

def find_files(directory, patterns):
    """Find files matching any of the given glob patterns (case-insensitive on Windows)."""
    matches = []
    for pattern in patterns:
        found = glob.glob(os.path.join(directory, pattern))
        # Also try uppercase variants
        found += glob.glob(os.path.join(directory, pattern.upper()))
        found += glob.glob(os.path.join(directory, pattern.lower()))
        matches.extend(found)
    # Deduplicate by resolved absolute path
    seen = set()
    unique = []
    for f in matches:
        key = os.path.abspath(f).lower()
        if key not in seen:
            seen.add(key)
            unique.append(os.path.abspath(f))
    return unique

FILE_SPECS = {
    "balance_generale": {
        "label_fr": "Balance Générale",
        "label_en": "Trial Balance",
        "patterns": ["BG*.xlsx", "Balance*.xlsx", "BALANCE*.xlsx", "bg*.xlsx"],
        "required": True,
    },
    "grand_livre": {
        "label_fr": "Grand Livre Général",
        "label_en": "General Ledger",
        "patterns": ["Grand Livre*.xls*", "GL*.xls*", "GRAND LIVRE*.xls*", "grand*livre*.xls*"],
        "required": True,
    },
    "livre_paie": {
        "label_fr": "Livre de Paie",
        "label_en": "Payroll Register",
        "patterns": ["LIVREPAIE*.CSV", "Livre*Paie*.csv", "LivrePaie*.csv", "livrepaie*.csv"],
        "required": True,
    },
    "charges_patronales": {
        "label_fr": "Charges Patronales",
        "label_en": "Employer Charges",
        "patterns": ["CHARGESPATRON*.CSV", "Charges*Patron*.csv", "chargespatron*.csv"],
        "required": True,
    },
    "feuille_travail": {
        "label_fr": "Feuille de Travail FT-P-2",
        "label_en": "Working Paper FT-P-2",
        "patterns": ["FT-P-2*.xlsx", "FT_P_2*.xlsx"],
        "required": False,  # Will be created from template if missing
    },
}

def main():
    parser = argparse.ArgumentParser(description="Detect audit source files")
    parser.add_argument("--dir",  default=".", help="Directory to scan")
    parser.add_argument("--plan", default=None, choices=["SYSCOHADA", "PCG"],
                        help="Accounting plan: SYSCOHADA (66x) or PCG (64x)")
    parser.add_argument("--json-only", action="store_true", help="Output JSON only")
    args = parser.parse_args()

    scan_dir = os.path.abspath(args.dir)

    if not os.path.isdir(scan_dir):
        print(json.dumps({"error": f"Directory not found: {scan_dir}"}))
        sys.exit(1)

    results = {}
    all_found = True

    for key, spec in FILE_SPECS.items():
        matches = find_files(scan_dir, spec["patterns"])
        if len(matches) == 1:
            status = "found"
            path = matches[0]
        elif len(matches) > 1:
            status = "multiple"
            path = matches[0]  # Default to first; agent will ask user to confirm
        else:
            status = "not_found"
            path = None
            if spec["required"]:
                all_found = False

        results[key] = {
            "status": status,
            "path": path,
            "all_matches": matches,
            "label_fr": spec["label_fr"],
            "label_en": spec["label_en"],
            "required": spec["required"],
        }

    output = {
        "scan_dir": scan_dir,
        "all_found": all_found,
        "files": results,
    }

    print(json.dumps(output, indent=2, ensure_ascii=False))

    # Update .audit-session.json if it exists
    session_path = os.path.join(scan_dir, ".audit-session.json")
    session = {}
    if os.path.exists(session_path):
        with open(session_path, encoding="utf-8") as f:
            session = json.load(f)

    session.setdefault("working_dir", scan_dir)
    session.setdefault("files", {})
    for key, info in results.items():
        if info["path"]:
            session["files"][key] = info["path"]

    # Store accounting plan if provided
    if args.plan:
        session["accounting_plan"] = args.plan
        account_range = "66x" if args.plan == "SYSCOHADA" else "64x"
        print(f"Accounting plan: {args.plan} (charges de personnel = comptes {account_range})")

    with open(session_path, "w", encoding="utf-8") as f:
        json.dump(session, f, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main()
