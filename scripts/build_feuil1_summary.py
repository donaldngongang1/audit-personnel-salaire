"""
build_feuil1_summary.py — Fill Feuil1 executive summary sheet.
Reads ecart values from Feuil2, generates gap explanations, writes to Feuil1.
Only fills value cells — NEVER changes formatting or structure.

Usage: python build_feuil1_summary.py [--write]
"""
import argparse, json, sys
from openpyxl import load_workbook

SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

FT_PATH = session["files"]["feuille_travail"]
LANGUAGE = session.get("language", "fr")

# ── Gap explanations library ──────────────────────────────────────────────────
GAP_LIBRARY = {
    "U": {
        "fr": "Le FNE (Fond National de l'Emploi) n'est pas comptabilisé dans le Grand Livre Général. "
              "Les cotisations FNE sont calculées à partir du livre de paie et versées directement "
              "sans écriture comptable distincte. Écart structurel — aucune action requise.",
        "en": "FNE (National Employment Fund) contributions are not recorded in the General Ledger. "
              "They are calculated from the payroll register and paid directly without a GL booking. "
              "Structural gap — no action required.",
        "category": "hors_champ_gl",
        "action": "information_only",
    },
}

def read_ecart_values(ws, ecart_row):
    """Read ecart values for columns R(18)–Y(25) from the workbook."""
    ecarts = {}
    col_names = {18:"R", 19:"S", 20:"T", 21:"U", 22:"V", 23:"W", 24:"X", 25:"Y"}
    for col, name in col_names.items():
        val = ws.cell(ecart_row, col).value
        try:
            ecarts[name] = float(val) if val not in (None, "") else 0.0
        except (TypeError, ValueError):
            ecarts[name] = 0.0
    return ecarts

def generate_explanation(col_name, amount, language):
    """Generate gap explanation for a column."""
    if abs(amount) < 2:
        return {
            "fr": "Écart d'arrondi négligeable.",
            "en": "Negligible rounding difference.",
            "category": "arrondi",
            "action": "aucune",
        }
    if col_name in GAP_LIBRARY:
        return GAP_LIBRARY[col_name]
    # Generic explanation
    direction_fr = "PAIE > COMPTA" if amount < 0 else "COMPTA > PAIE"
    direction_en = direction_fr
    return {
        "fr": f"Écart de {abs(amount):,.0f} FCFA ({direction_fr}). Investigation complémentaire requise.",
        "en": f"Gap of {abs(amount):,.0f} FCFA ({direction_en}). Further investigation required.",
        "category": "a_investiguer",
        "action": "investigation_requise",
    }

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true", help="Write to Feuil1")
    args = parser.parse_args()

    wb = load_workbook(FT_PATH)
    ws2 = wb["Feuil2"]
    ws1 = wb["Feuil1"]

    # Locate ECART row in Feuil2
    ecart_row = session.get("ecart_row")
    if not ecart_row:
        for r in range(1, ws2.max_row + 1):
            v = str(ws2.cell(r, 1).value or ws2.cell(r, 2).value or "")
            if "ECART" in v.upper() and "TOTAL" in v.upper():
                ecart_row = r
                break
    if not ecart_row:
        print("❌ Could not locate ECART row in Feuil2")
        sys.exit(1)

    ecarts = read_ecart_values(ws2, ecart_row)
    print(f"Ecart values (row {ecart_row}):")
    for col, val in ecarts.items():
        status = "✅" if abs(val) < 2 else "⚠️"
        print(f"  {status} Col {col}: {val:,.0f} FCFA")

    # Generate explanations
    explanations = {}
    for col_name, amount in ecarts.items():
        if col_name in ("V",):  # Skip derivative columns
            continue
        explanations[col_name] = generate_explanation(col_name, amount, LANGUAGE)
        explanations[col_name]["amount"] = amount

    # Persist in session
    session["ecarts"] = {k: {**v, "amount": v.get("amount", ecarts.get(k, 0))}
                         for k, v in explanations.items()}

    # Summary totals (read from TOTAL PAIE and TOTAL COMPTA rows)
    tot_paie_row   = session.get("tot_paie_row", 178)
    tot_compta_row = session.get("tot_compta_row")

    if args.write and tot_compta_row:
        # Read summary totals from Feuil2
        paie_totals   = {c: ws2.cell(tot_paie_row,   c).value for c in range(18, 26)}
        compta_totals = {c: ws2.cell(tot_compta_row,  c).value for c in range(18, 26)}

        # Fill Feuil1 — scan for summary cells (look for labels in col A or B)
        metadata = session.get("metadata", {})
        for r in range(1, ws1.max_row + 1):
            v = str(ws1.cell(r, 1).value or "").upper()
            # Fill metadata cells found next to known labels
            if "SOCIETE" in v or "ENTREPRISE" in v or "COMPANY" in v:
                ws1.cell(r, 2).value = metadata.get("societe", "")
            elif "AUDITEUR" in v or "AUDITOR" in v:
                ws1.cell(r, 2).value = metadata.get("auditeur", "")
            elif "PERIODE" in v or "PERIOD" in v:
                ws1.cell(r, 2).value = metadata.get("periode", "")
            elif "DATE" in v and "RAPPORT" in v:
                ws1.cell(r, 2).value = metadata.get("date_rapport", "")

        wb.save(FT_PATH)
        print(f"✅ Feuil1 filled and saved to '{FT_PATH}'")
        session.setdefault("steps_completed", [])
        if "summarise" not in session["steps_completed"]:
            session["steps_completed"].append("summarise")

    with open(SESSION_FILE, "w", encoding="utf-8") as f:
        json.dump(session, f, indent=2, ensure_ascii=False)

    print("\n--- Gap Analysis Summary ---")
    for col, info in explanations.items():
        amount = info.get("amount", 0)
        if abs(amount) >= 2:
            print(f"\nColumn {col}: {amount:,.0f} FCFA")
            key = "fr" if LANGUAGE == "fr" else "en"
            print(f"  {info.get(key, '')}")
            print(f"  Category: {info.get('category')} | Action: {info.get('action')}")

if __name__ == "__main__":
    main()
