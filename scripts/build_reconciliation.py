"""
build_reconciliation.py — Build Feuil2 reconciliation workpaper.
Sections: PAIE (employee rows), COMPTABILITE (GL accounts), TOTAL rows, ECART row.

Usage: python build_reconciliation.py [--section paie|compta|all]
"""
import argparse, json, re, sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gcl

# ── Load session ────────────────────────────────────────────────────────────────
SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

FT_PATH = session["files"]["feuille_travail"]

# ── Styling constants ───────────────────────────────────────────────────────────
NUM_FMT  = "#,##0;(#,##0);\"-\""
BLANK_COLS = [14, 15, 16, 17]   # N, O, P, Q — always blank

def med(c="1F4E79"):  return Side(style="medium", color=c)
def thn(c="D9D9D9"):  return Side(style="thin",   color=c)

BDR_DATA = Border(bottom=thn(), right=thn())
BDR_TOT  = Border(top=med(),  bottom=med(), right=thn())

RGT = Alignment(horizontal="right",  vertical="center")
LFT = Alignment(horizontal="left",   vertical="center")

FDATA   = Font(name="Arial", size=10)
FWHITE  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
FTOT_P  = Font(name="Arial", bold=True, size=10, color="1F4E79")
FTOT_C  = Font(name="Arial", bold=True, size=10, color="1F4E79")

FILL_ALT    = PatternFill("solid", start_color="F5F8FC")
FILL_TOT_P  = PatternFill("solid", start_color="D6E4F0")
FILL_TOT_C  = PatternFill("solid", start_color="C6EFCE")
FILL_ECART  = PatternFill("solid", start_color="843C0C")
FILL_NONE   = PatternFill(fill_type=None)

def blank_cell(ws, row, col, fill=None):
    c = ws.cell(row=row, column=col)
    c.value = None; c.font = FDATA; c.number_format = "General"
    c.alignment = RGT; c.border = BDR_DATA
    c.fill = fill if fill else FILL_NONE

def num_cell(ws, row, col, value, font, fill, bdr=BDR_DATA):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font; c.fill = fill if fill else FILL_NONE
    c.number_format = NUM_FMT; c.alignment = RGT; c.border = bdr
    return c

def formula_cell(ws, row, col, formula, font, fill, bdr=BDR_DATA):
    c = ws.cell(row=row, column=col, value=formula)
    c.font = font; c.fill = fill if fill else FILL_NONE
    c.number_format = NUM_FMT; c.alignment = RGT; c.border = bdr
    return c

# ── Locate key rows in Feuil2 ──────────────────────────────────────────────────
def find_feuil2_rows(ws):
    """Scan column A/B for section labels to determine row layout."""
    rows = {}
    for r in range(1, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or ws.cell(r, 2).value or "")
        vu = v.upper()
        if "TOTAL PAIE" in vu and "COMPTA" not in vu and "TOT_PAIE" not in rows:
            rows["TOT_PAIE"] = r
        elif "TOTAL COMPTABILITE" in vu and "TOT_COMPTA" not in rows:
            rows["TOT_COMPTA"] = r
        elif "ECART" in vu and "TOTAL" in vu and "ECART_ROW" not in rows:
            rows["ECART_ROW"] = r
    # DATA_START = row after the PAIE header (typically row 4)
    rows.setdefault("DATA_START", 4)
    return rows

# ── Load payroll pivot data ─────────────────────────────────────────────────────
def load_paie_data():
    lp = pd.read_csv(".livre_paie_pivot.csv", dtype={"Matricule": str})
    cp = pd.read_csv(".charges_patronales_pivot.csv", dtype={"Matricule": str})
    merged = pd.merge(lp, cp, on=["Matricule", "Nom", "Prenom"], how="outer").fillna(0)
    # Ensure expected columns
    for col in ["SAL BRUT",
                "CF/P (Crédit Foncier Patronal)",
                "FNE (Fond National Emploi)",
                "CNPS/P (Pension Vieillesse)",
                "AF (Allocation Familiale)",
                "AT (Accident de Travail)"]:
        if col not in merged.columns:
            merged[col] = 0
    return merged.sort_values("Matricule").reset_index(drop=True)

# ── Load COMPTA data from Balance Générale ──────────────────────────────────────
def load_compta_data():
    bg_path = session["files"]["balance_generale"]
    df = pd.read_excel(bg_path, dtype={"Compte": str})
    df.columns = [str(c).strip() for c in df.columns]
    compte_col = next((c for c in df.columns if c.lower() in ["compte","account"]), df.columns[0])
    mvtd_col   = df.columns[4] if len(df.columns) > 4 else None
    mvtc_col   = df.columns[5] if len(df.columns) > 5 else None
    df[compte_col] = df[compte_col].astype(str).str.strip()
    df["NetSolde"] = pd.to_numeric(df.get(mvtd_col, 0), errors="coerce").fillna(0) - \
                     pd.to_numeric(df.get(mvtc_col, 0), errors="coerce").fillna(0)
    account_map = {row[compte_col]: round(row["NetSolde"]) for _, row in df.iterrows()}
    return account_map

# ── Section definitions ────────────────────────────────────────────────────────
# Each tuple: (account_code, libelle, target_col_index)
# target_col_index: R=18, S=19, T=20, U=21, W=23, X=24
GROUP_A_ACCOUNTS = [
    "661110", "661120", "661130", "661200", "661210",
    "661220", "661300", "661380", "661410", "661800",
    "663101", "663102", "663410",
]
GROUP_B_ACCOUNTS = [
    ("664120", "CNPS Pension Vieillesse (AV)", 19),
    ("664110", "CNPS Allocation Familiale (AF)", 23),
    ("664130", "CNPS Accident de Travail (AT)", 24),
]
GROUP_C_ACCOUNTS = [
    ("664380", "Provision CF/P", 20),
    ("FNE_NA", "FNE non comptabilisé", 21),  # Always 0
]
GROUP_D_ACCOUNTS = [
    ("668420", "Charges sociales diverses 1", 0),
    ("668430", "Charges sociales diverses 2", 0),
    ("668700", "Autres charges de personnel", 0),
]

def build_paie_section(ws, paie_df, row_map):
    data_start = row_map.get("DATA_START", 4)
    tot_paie   = row_map.get("TOT_PAIE", 178)
    paie_end   = tot_paie - 1

    print(f"Building PAIE section: rows {data_start}–{paie_end} ({len(paie_df)} employees)")

    for i, emp in enumerate(paie_df.itertuples(index=False)):
        r = data_start + i
        if r >= tot_paie:
            print(f"⚠️ More employees than PAIE rows available! Stopping at row {r-1}")
            break
        alt = FILL_ALT if (i % 2 == 0) else None

        # Blank N/O/P/Q
        for col in BLANK_COLS:
            blank_cell(ws, r, col, fill=alt)

        # Write numeric columns
        brut = round(getattr(emp, "SAL BRUT", 0) or 0)
        cnps = round(getattr(emp, "CNPS/P (Pension Vieillesse)", 0) or 0)
        cfp  = round(getattr(emp, "CF/P (Crédit Foncier Patronal)", 0) or 0)
        fne  = round(getattr(emp, "FNE (Fond National Emploi)", 0) or 0)
        af   = round(getattr(emp, "AF (Allocation Familiale)", 0) or 0)
        at   = round(getattr(emp, "AT (Accident de Travail)", 0) or 0)

        num_cell(ws, r, 18, brut, FDATA, alt)   # R = SAL BRUT
        num_cell(ws, r, 19, cnps, FDATA, alt)   # S = CNPS/P
        num_cell(ws, r, 20, cfp,  FDATA, alt)   # T = CF/P
        num_cell(ws, r, 21, fne,  FDATA, alt)   # U = FNE
        formula_cell(ws, r, 22, f"=T{r}+U{r}", FDATA, alt)  # V = T+U
        num_cell(ws, r, 23, af,   FDATA, alt)   # W = AF
        num_cell(ws, r, 24, at,   FDATA, alt)   # X = AT
        formula_cell(ws, r, 25, f"=R{r}+S{r}+T{r}+U{r}+W{r}+X{r}", FDATA, alt)  # Y

    # TOTAL PAIE row
    r = tot_paie
    for col in BLANK_COLS:
        c = ws.cell(r, col)
        c.value = None; c.fill = FILL_TOT_P; c.border = BDR_TOT; c.font = FTOT_P

    for col in range(18, 25):  # R–X
        formula_cell(ws, r, col,
                     f"=SUM({gcl(col)}{data_start}:{gcl(col)}{paie_end})",
                     FTOT_P, FILL_TOT_P, BDR_TOT)
    formula_cell(ws, r, 25,
                 f"=SUM(Y{data_start}:Y{paie_end})",
                 FTOT_P, FILL_TOT_P, BDR_TOT)

    print(f"✅ PAIE section written ({len(paie_df)} employees + TOTAL row {r})")
    return tot_paie

def build_compta_section(ws, account_map, start_row):
    """Write all COMPTA groups starting at start_row. Returns TOTAL COMPTA row number."""
    r = start_row
    group_subtotals = []

    def write_group(accounts_spec, group_label, is_group_a=False):
        nonlocal r
        group_start = r
        for spec in accounts_spec:
            alt = FILL_ALT if ((r - group_start) % 2 == 0) else None
            for col in BLANK_COLS:
                blank_cell(ws, r, col, fill=alt)
            # Write account code
            if is_group_a:
                acct = spec
                amount = round(account_map.get(acct, 0))
                # Identity cols (A=1, B=2 = account label)
                ws.cell(r, 1).value = acct
                num_cell(ws, r, 18, amount, FDATA, alt)  # R
                formula_cell(ws, r, 22, f"=T{r}+U{r}", FDATA, alt)
                formula_cell(ws, r, 25, f"=R{r}+S{r}+T{r}+U{r}+W{r}+X{r}", FDATA, alt)
            else:
                acct, libelle, target_col = spec
                amount = 0 if acct == "FNE_NA" else round(account_map.get(acct, 0))
                ws.cell(r, 1).value = "" if acct == "FNE_NA" else acct
                ws.cell(r, 2).value = libelle
                if target_col > 0:
                    num_cell(ws, r, target_col, amount, FDATA, alt)
                formula_cell(ws, r, 22, f"=T{r}+U{r}", FDATA, alt)
                formula_cell(ws, r, 25, f"=R{r}+S{r}+T{r}+U{r}+W{r}+X{r}", FDATA, alt)
            r += 1

        # Subtotal row
        sub_r = r
        for col in range(18, 26):
            formula_cell(ws, sub_r, col,
                         f"=SUM({gcl(col)}{group_start}:{gcl(col)}{r-1})",
                         FTOT_C, FILL_TOT_C, BDR_TOT)
        group_subtotals.append(sub_r)
        print(f"  {group_label}: rows {group_start}–{r-1}, subtotal row {sub_r}")
        r += 2  # gap row + next group start

    write_group(GROUP_A_ACCOUNTS, "Group A (661-663)", is_group_a=True)
    write_group(GROUP_B_ACCOUNTS, "Group B (CNPS)")
    write_group(GROUP_C_ACCOUNTS, "Group C (CF/P + FNE)")
    # Group D — info only, grey
    write_group(GROUP_D_ACCOUNTS, "Group D (668xxx — info)")

    # TOTAL COMPTABILITE = A + B + C subtotals (D excluded)
    tot_compta_row = r
    if len(group_subtotals) >= 3:
        a_sub, b_sub, c_sub = group_subtotals[0], group_subtotals[1], group_subtotals[2]
        for col in range(18, 26):
            formula_cell(ws, tot_compta_row, col,
                         f"=R{a_sub}+R{b_sub}+R{c_sub}" if col == 18 else
                         f"={gcl(col)}{a_sub}+{gcl(col)}{b_sub}+{gcl(col)}{c_sub}",
                         FTOT_C, FILL_TOT_C, BDR_TOT)

    print(f"✅ COMPTA section written. TOTAL COMPTA: row {tot_compta_row}")
    return tot_compta_row

def build_ecart_row(ws, tot_compta_row, tot_paie_row):
    ecart_row = tot_compta_row + 3
    for col in BLANK_COLS:
        c = ws.cell(ecart_row, col)
        c.value = None; c.fill = FILL_ECART; c.border = BDR_TOT; c.font = FWHITE

    for col in range(18, 25):
        formula_cell(ws, ecart_row, col,
                     f"={gcl(col)}{tot_compta_row}-{gcl(col)}{tot_paie_row}",
                     FWHITE, FILL_ECART, BDR_TOT)
    formula_cell(ws, ecart_row, 25,
                 f"=Y{tot_compta_row}-Y{tot_paie_row}",
                 FWHITE, FILL_ECART, BDR_TOT)

    print(f"✅ ECART row: {ecart_row}")
    return ecart_row

# ── Main ────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--section", default="all", choices=["paie", "compta", "all"])
    args = parser.parse_args()

    wb = load_workbook(FT_PATH)
    ws = wb["Feuil2"]

    row_map = find_feuil2_rows(ws)
    print(f"Row layout: {row_map}")

    tot_paie_row = row_map.get("TOT_PAIE", 178)
    compta_start = tot_paie_row + 3  # Leave gap after TOTAL PAIE

    # Unmerge COMPTA region before writing
    to_unmerge = [mr for mr in list(ws.merged_cells.ranges)
                  if mr.min_row >= compta_start]
    for mr in to_unmerge:
        ws.unmerge_cells(str(mr))
    if to_unmerge:
        print(f"Unmerged {len(to_unmerge)} merged cell regions in COMPTA area")

    if args.section in ("paie", "all"):
        paie_df = load_paie_data()
        build_paie_section(ws, paie_df, row_map)

    tot_compta_row = None
    if args.section in ("compta", "all"):
        account_map = load_compta_data()
        tot_compta_row = build_compta_section(ws, account_map, compta_start)

    if args.section == "all" and tot_compta_row:
        build_ecart_row(ws, tot_compta_row, tot_paie_row)

    wb.save(FT_PATH)
    print(f"\n✅ Saved: {FT_PATH}")

    # Update session
    session.setdefault("steps_completed", [])
    if "reconcile" not in session["steps_completed"]:
        session["steps_completed"].append("reconcile")
    if tot_compta_row:
        session["tot_compta_row"] = tot_compta_row
    session["tot_paie_row"] = tot_paie_row
    with open(SESSION_FILE, "w", encoding="utf-8") as f:
        json.dump(session, f, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main()
