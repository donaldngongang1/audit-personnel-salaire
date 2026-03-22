"""
build_extracts.py — Orchestrator: run all four extract parsers in sequence.
Usage: python build_extracts.py [--sheet balance|gl|paie|charges|all]
"""
import argparse, subprocess, sys, os

def run_script(script_name):
    print(f"\n{'='*60}")
    print(f"Running {script_name}...")
    print('='*60)
    result = subprocess.run([sys.executable, os.path.join("scripts", script_name)], check=False)
    if result.returncode != 0:
        print(f"❌ {script_name} failed with exit code {result.returncode}")
        return False
    return True

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--sheet", default="all",
                        choices=["balance", "gl", "paie", "charges", "all"])
    args = parser.parse_args()

    scripts_map = {
        "balance": "parse_balance.py",
        "gl":      "parse_grand_livre.py",
        "charges": "parse_charges_patronales.py",
        "paie":    "parse_livre_paie.py",
    }

    if args.sheet == "all":
        to_run = list(scripts_map.values())
    else:
        to_run = [scripts_map[args.sheet]]

    all_ok = True
    for script in to_run:
        ok = run_script(script)
        if not ok:
            all_ok = False

    print("\n" + "="*60)
    if all_ok:
        print("✅ All extractions completed successfully")
    else:
        print("❌ One or more extractions failed — check output above")
    print("="*60)

if __name__ == "__main__":
    main()
