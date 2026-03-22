"""
build_reconciliation.py — Build Feuil2 reconciliation workpaper (v1.1 — SYSCOHADA corrections).

Section A source: Balance Générale (BG xlsx) — col4=MvtDebit, col5=MvtCredit
Section B/C/D source: Grand Livre (GL xls) — ALL journals

Excel formulas:
  V = =T{r}+U{r}
  Y = =R{r}+S{r}+V{r}+W{r}+X{r}  ← V not T+U separately
  TOTAL PAIE: =SUM(...)
  Subtotals: =SUM(...)
  TOTAL COMPTA: =col{stA}+col{stB}+col{stC}
  ECART: =col{compta}-col{paie}

Usage: python build_reconciliation.py [--section paie|compta|all]
"""
import argparse, io, json, subprocess, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
import xlrd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gcl

# ── Load session ───────────────────────────────────────────────────────────────
SESSION_FILE = ".audit-session.json"
with open(SESSION_FILE, encoding="utf-8") as f:
    session = json.load(f)

FT_PATH  = session["files"]["feuille_travail"]
BG_PATH  = session["files"]["balance_generale"]
GL_PATH  = session["files"]["grand_livre"]

# ── Styling ────────────────────────────────────────────────────────────────────
NUM_FMT    = "#,##0;(#,##0);\"-\""
BLANK_COLS = [14, 15, 16, 17]  # N, O, P, Q — always blank

def med(c="1F4E79"): return Side(style="medium", color=c)
def thn(c="D9D9D9"): return Side(style="thin",   color=c)

BDR_DATA = Border(bottom=thn(), right=thn())
BDR_TOT  = Border(top=med(), bottom=med(), right=thn())

RGT = Alignment(horizontal="right", vertical="center")
LFT = Alignment(horizontal="left",  vertical="center")

FDATA   = Font(name="Arial", size=10)
FWHITE  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
FTOT_P  = Font(name="Arial", bold=True, size=10, color="1F4E79")
FTOT_C  = Font(name="Arial", bold=True, size=10, color="1F4E79")
FGRP    = Font(name="Arial", bold=True, size=10, color="1F4E79")

FILL_ALT   = PatternFill("solid", start_color="F5F8FC")
FILL_TOT_P = PatternFill("solid", start_color="D6E4F0")
FILL_TOT_C = PatternFill("solid", start_color="C6EFCE")
FILL_ECART = PatternFill("solid", start_color="843C0C")
FILL_GRP   = PatternFill("solid", start_color="EBF3FB")
FILL_STOT  = PatternFill("solid", start_color="2E75B6")
FILL_NONE  = PatternFill(fill_type=None)

def blank_cell(ws, row, col, fill=None):
    c = ws.cell(row, col)
    c.value = None; c.font = FDATA; c.number_format = "General"
    c.alignment = RGT; c.border = BDR_DATA
    c.fill = fill if fill else FILL_NONE

def num_cell(ws, row, col, value, font, fill, bdr=BDR_DATA):
    c = ws.cell(row, col, value=value)
    c.font = font; c.fill = fill if fill else FILL_NONE
    c.number_format = NUM_FMT; c.alignment = RGT; c.border = bdr

def formula_cell(ws, row, col, formula, font, fill, bdr=BDR_DATA):
    c = ws.cell(row, col, value=formula)
    c.font = font; c.fill = fill if fill else FILL_NONE
    c.number_format = NUM_FMT; c.alignment = RGT; c.border = bdr

def label_cell(ws, row, col, value, font, fill, bdr=BDR_DATA):
    c = ws.cell(row, col, value=value)
    c.font = font; c.fill = fill if fill else FILL_NONE
    c.alignment = LFT; c.border = bdr

# ── Locate key rows ────────────────────────────────────────────────────────────
def find_feuil2_rows(ws):
    rows = {}
    for r in range(1, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or ws.cell(r, 2).value or "")
        vu = v.upper()
        if "TOTAL PAIE" in vu and "COMPTA" not in vu and "TOT_PAIE" not in rows:
            rows["TOT_PAIE"] = r
    rows.setdefault("DATA_START", 4)
    return rows

# ── Load payroll pivot ─────────────────────────────────────────────────────────
def load_paie_data():
    lp = pd.read_csv(".livre_paie_pivot.csv", dtype={"Matricule": str})
    cp = pd.read_csv(".charges_patronales_pivot.csv", dtype={"Matricule": str})

    # Normalize charges patronales column names (saved with underscores)
    COL_MAP = {
        "Credit_Foncier_Patronal":        "CF_P",
        "Fond_National_de_lemploi_FNE":   "FNE",
        "Pension_Vieillesse_CNPS":        "CNPS_P",
        "Allocation_Familiale":           "AF",
        "Accident_de_Travail":            "AT",
    }
    cp = cp.rename(columns={k: v for k, v in COL_MAP.items() if k in cp.columns})

    merged = pd.merge(lp, cp, on=["Matricule", "Nom", "Prenom"], how="outer").fillna(0)
    for col in ["SAL_BRUT", "CF_P", "FNE", "CNPS_P", "AF", "AT"]:
        if col not in merged.columns:
            merged[col] = 0
    return merged.sort_values("Matricule").reset_index(drop=True)

# ── Load BG amounts (Section A: 661x+663x) ────────────────────────────────────
def load_bg_section_a():
    df = pd.read_excel(BG_PATH, dtype={0: str})
    df.columns = [str(c).strip() for c in df.columns]
    compte_col = df.columns[0]
    df[compte_col] = df[compte_col].astype(str).str.strip()
    df["MvtDebit"]  = pd.to_numeric(df.iloc[:, 4], errors="coerce").fillna(0)
    df["MvtCredit"] = pd.to_numeric(df.iloc[:, 5], errors="coerce").fillna(0)
    df["SoldeNet"]  = df["MvtDebit"] - df["MvtCredit"]

    # Filter 661x + 663x
    mask = df[compte_col].str.match(r'^661|^663')
    df_a = df[mask & (df["SoldeNet"].abs() > 0)].copy()

    account_map = {str(row[compte_col]): {"libelle": str(row.iloc[1]), "solde": round(row["SoldeNet"])}
                   for _, row in df_a.iterrows()}
    return account_map, round(df_a["SoldeNet"].sum())

# ── Load GL amounts (Sections B, C, D) ────────────────────────────────────────
def load_gl_sections():
    """Read GL for all 664x+668x accounts, ALL journals."""
    xls = xlrd.open_workbook(GL_PATH)
    sheet_names = xls.sheet_names()
    sheet_name = "Sage" if "Sage" in sheet_names else sheet_names[0]

    df = pd.read_excel(GL_PATH, sheet_name=sheet_name, engine="xlrd", dtype={0: str})
    df.columns = [str(c).strip() for c in df.columns]
    compte_col = df.columns[0]
    df[compte_col] = df[compte_col].astype(str).str.strip()

    debit_col  = next((c for c in df.columns if "debit"  in c.lower()), df.columns[8] if len(df.columns) > 8 else None)
    credit_col = next((c for c in df.columns if "credit" in c.lower()), df.columns[9] if len(df.columns) > 9 else None)
    libelle_col = df.columns[4] if len(df.columns) > 4 else df.columns[1]

    df["Debit"]  = pd.to_numeric(df[debit_col],  errors="coerce").fillna(0) if debit_col  else 0
    df["Credit"] = pd.to_numeric(df[credit_col], errors="coerce").fillna(0) if credit_col else 0

    # Aggregate by account (ALL journals)
    mask_664 = df[compte_col].str.match(r'^664')
    mask_668 = df[compte_col].str.match(r'^668')
    df_6xx = df[mask_664 | mask_668].copy()

    agg = df_6xx.groupby(compte_col).agg(
        Debit=("Debit", "sum"),
        Credit=("Credit", "sum"),
        Libelle=(libelle_col, "first")
    ).reset_index()
    agg["SoldeNet"] = agg["Debit"] - agg["Credit"]

    account_map = {str(row[compte_col]): {"libelle": str(row["Libelle"]), "solde": round(row["SoldeNet"])}
                   for _, row in agg.iterrows()}
    return account_map

# ── PAIE section ───────────────────────────────────────────────────────────────
def build_paie_section(ws, paie_df, data_start, tot_paie_row):
    paie_end = tot_paie_row - 1
    print(f"Building PAIE section: rows {data_start}–{paie_end} ({len(paie_df)} employees)")

    for i, emp in enumerate(paie_df.itertuples(index=False)):
        r = data_start + i
        if r >= tot_paie_row:
            print(f"⚠️ More employees than PAIE rows — stopping at row {r-1}")
            break
        alt = FILL_ALT if (i % 2 == 0) else None

        for col in BLANK_COLS:
            blank_cell(ws, r, col, fill=alt)

        brut = round(getattr(emp, "SAL_BRUT",  0) or 0)
        cnps = round(getattr(emp, "CNPS_P",    0) or 0)
        cfp  = round(getattr(emp, "CF_P",      0) or 0)
        fne  = round(getattr(emp, "FNE",       0) or 0)
        af   = round(getattr(emp, "AF",        0) or 0)
        at   = round(getattr(emp, "AT",        0) or 0)

        num_cell(ws, r, 18, brut, FDATA, alt)   # R = SAL BRUT
        num_cell(ws, r, 19, cnps, FDATA, alt)   # S = CNPS/P
        num_cell(ws, r, 20, cfp,  FDATA, alt)   # T = CF/P
        num_cell(ws, r, 21, fne,  FDATA, alt)   # U = FNE
        formula_cell(ws, r, 22, f"=T{r}+U{r}", FDATA, alt)               # V = CF/P+FNE
        num_cell(ws, r, 23, af,   FDATA, alt)   # W = AF
        num_cell(ws, r, 24, at,   FDATA, alt)   # X = AT
        formula_cell(ws, r, 25, f"=R{r}+S{r}+V{r}+W{r}+X{r}", FDATA, alt)  # Y (uses V!)

    # TOTAL PAIE row
    r = tot_paie_row
    for col in BLANK_COLS:
        c = ws.cell(r, col)
        c.value = None; c.fill = FILL_TOT_P; c.border = BDR_TOT; c.font = FTOT_P

    for col in range(18, 25):
        formula_cell(ws, r, col,
                     f"=SUM({gcl(col)}{data_start}:{gcl(col)}{paie_end})",
                     FTOT_P, FILL_TOT_P, BDR_TOT)
    formula_cell(ws, r, 25,
                 f"=SUM(Y{data_start}:Y{paie_end})",
                 FTOT_P, FILL_TOT_P, BDR_TOT)

    print(f"  TOTAL PAIE at row {r} with SUM formulas (rows {data_start}:{paie_end})")

# ── COMPTA section ─────────────────────────────────────────────────────────────
def write_group_label(ws, row, label):
    for col in range(1, 26):
        c = ws.cell(row, col)
        c.value = None; c.fill = FILL_GRP; c.border = BDR_TOT; c.font = FGRP
    ws.cell(row, 1, value=label).alignment = LFT

def write_account_row(ws, row, compte, libelle, col_idx, amount, fill=None):
    for col in BLANK_COLS:
        blank_cell(ws, row, col, fill=fill)
    label_cell(ws, row, 1, compte, FDATA, fill)
    label_cell(ws, row, 2, libelle, FDATA, fill)
    if col_idx > 0:
        num_cell(ws, row, col_idx, amount, FDATA, fill)
    # V and Y formulas
    formula_cell(ws, row, 22, f"=T{row}+U{row}", FDATA, fill)
    formula_cell(ws, row, 25, f"=R{row}+S{row}+V{row}+W{row}+X{row}", FDATA, fill)

def write_subtotal_row(ws, row, label, group_start, group_end):
    for col in BLANK_COLS:
        c = ws.cell(row, col)
        c.value = None; c.fill = FILL_STOT; c.border = BDR_TOT; c.font = FWHITE
    label_cell(ws, row, 1, label, FWHITE, FILL_STOT, BDR_TOT)
    for col in range(18, 26):
        formula_cell(ws, row, col,
                     f"=SUM({gcl(col)}{group_start}:{gcl(col)}{group_end})",
                     FWHITE, FILL_STOT, BDR_TOT)

def build_compta_section(ws, bg_map, gl_map, start_row):
    r = start_row
    subtotal_rows = []

    # ── GROUP A — Rémunérations directes 661x+663x (source: BG) ──────────────
    write_group_label(ws, r, "Section A — Rémunérations directes et avantages (661x + 663x) — Source: Balance Générale")
    r += 1
    group_a_start = r

    A_ACCOUNTS = ["661110","661120","661130","661200","661210","661220",
                  "661300","661380","661410","661800","663101","663102","663410"]
    for i, acct in enumerate(A_ACCOUNTS):
        info = bg_map.get(acct, {"libelle": acct, "solde": 0})
        fill = FILL_ALT if (i % 2 == 0) else None
        write_account_row(ws, r, acct, info["libelle"], 18, info["solde"], fill=fill)
        r += 1

    group_a_end = r - 1
    write_subtotal_row(ws, r, "Sous-total A — 661x+663x", group_a_start, group_a_end)
    subtotal_rows.append(r)
    print(f"  Group A: rows {group_a_start}–{group_a_end}, subtotal row {r}")
    r += 2  # spacer

    # ── GROUP B — Cotisations CNPS (source: GL tous journaux) ─────────────────
    write_group_label(ws, r, "Section B — Cotisations CNPS (664110 AF | 664120 CNPS/P | 664130 AT) — Source: Grand Livre")
    r += 1
    group_b_start = r

    B_ACCOUNTS = [
        ("664120", "CNPS Pension Vieillesse (AV)", 19),   # S
        ("664110", "CNPS Allocation Familiale (AF)", 23),  # W
        ("664130", "CNPS Accident de Travail (AT)", 24),   # X
    ]
    for i, (acct, fallback_lib, col_idx) in enumerate(B_ACCOUNTS):
        info = gl_map.get(acct, {"libelle": fallback_lib, "solde": 0})
        fill = FILL_ALT if (i % 2 == 0) else None
        write_account_row(ws, r, acct, info["libelle"], col_idx, info["solde"], fill=fill)
        r += 1

    group_b_end = r - 1
    write_subtotal_row(ws, r, "Sous-total B — CNPS", group_b_start, group_b_end)
    subtotal_rows.append(r)
    print(f"  Group B: rows {group_b_start}–{group_b_end}, subtotal row {r}")
    r += 2

    # ── GROUP C — CF/P + FNE (source: GL) ────────────────────────────────────
    write_group_label(ws, r, "Section C — Crédit Foncier Patronal & FNE (664380 + FNE=0) — Source: Grand Livre")
    r += 1
    group_c_start = r

    # CF/P from GL
    cfp_info = gl_map.get("664380", {"libelle": "Provisions Crédit Foncier Patronal", "solde": 0})
    write_account_row(ws, r, "664380", cfp_info["libelle"], 20, cfp_info["solde"], fill=FILL_ALT)
    r += 1

    # FNE — structural zero
    write_account_row(ws, r, "—", "FNE — non comptabilisé en GL (retenue salariale hors 66x)", 21, 0, fill=None)
    ws.cell(r, 2).font = Font(name="Arial", italic=True, size=10, color="806000")
    r += 1

    group_c_end = r - 1
    write_subtotal_row(ws, r, "Sous-total C — CF/P+FNE", group_c_start, group_c_end)
    subtotal_rows.append(r)
    print(f"  Group C: rows {group_c_start}–{group_c_end}, subtotal row {r}")
    r += 2

    # ── GROUP D — Autres charges sociales 668x (info only) ───────────────────
    write_group_label(ws, r, "Section D — Autres charges sociales (668x) — Informatif uniquement — NON inclus dans TOTAL")
    r += 1
    group_d_start = r

    D_ACCOUNTS = ["668420", "668430", "668700"]
    for i, acct in enumerate(D_ACCOUNTS):
        info = gl_map.get(acct, {"libelle": acct, "solde": 0})
        fill = FILL_ALT if (i % 2 == 0) else None
        write_account_row(ws, r, acct, info["libelle"], 18, info["solde"], fill=fill)
        r += 1

    group_d_end = r - 1
    # Subtotal D (informative, grey)
    FILL_D = PatternFill("solid", start_color="808080")
    for col in range(1, 26):
        c = ws.cell(r, col)
        c.value = None; c.fill = FILL_D; c.border = BDR_TOT; c.font = FWHITE
    ws.cell(r, 1, value="Sous-total D — 668x (hors rapprochement)").alignment = LFT
    for col in range(18, 26):
        formula_cell(ws, r, col,
                     f"=SUM({gcl(col)}{group_d_start}:{gcl(col)}{group_d_end})",
                     FWHITE, FILL_D, BDR_TOT)
    print(f"  Group D (info): rows {group_d_start}–{group_d_end}, subtotal row {r}")
    r += 2

    # ── TOTAL COMPTABILITE = A + B + C (D excluded) ───────────────────────────
    tot_compta_row = r
    for col in range(1, 26):
        c = ws.cell(r, col)
        c.value = None; c.fill = FILL_TOT_C; c.border = BDR_TOT; c.font = FTOT_C
    ws.cell(r, 1, value="TOTAL COMPTABILITE (A+B+C)").alignment = LFT

    stA, stB, stC = subtotal_rows[0], subtotal_rows[1], subtotal_rows[2]
    for col in BLANK_COLS:
        blank_cell(ws, r, col, fill=FILL_TOT_C)
    for col in range(18, 26):
        formula_cell(ws, r, col,
                     f"={gcl(col)}{stA}+{gcl(col)}{stB}+{gcl(col)}{stC}",
                     FTOT_C, FILL_TOT_C, BDR_TOT)

    print(f"  TOTAL COMPTABILITE: row {tot_compta_row} = SUM(subtotals A{stA}+B{stB}+C{stC})")
    return tot_compta_row

# ── ECART row ─────────────────────────────────────────────────────────────────
def build_ecart_row(ws, tot_compta_row, tot_paie_row):
    r = tot_compta_row + 3
    for col in range(1, 26):
        c = ws.cell(r, col)
        c.value = None; c.fill = FILL_ECART; c.border = BDR_TOT; c.font = FWHITE
    ws.cell(r, 1, value="ECART TOTAL (COMPTABILITE - PAIE)").alignment = LFT
    for col in BLANK_COLS:
        blank_cell(ws, r, col, fill=FILL_ECART)
    for col in range(18, 26):
        formula_cell(ws, r, col,
                     f"={gcl(col)}{tot_compta_row}-{gcl(col)}{tot_paie_row}",
                     FWHITE, FILL_ECART, BDR_TOT)
    print(f"  ECART row: {r} = COMPTA{tot_compta_row}-PAIE{tot_paie_row}")
    return r

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--section", default="all", choices=["paie", "compta", "all"])
    args = parser.parse_args()

    wb = load_workbook(FT_PATH)
    ws = wb["Feuil2"]

    row_map = find_feuil2_rows(ws)
    data_start   = row_map.get("DATA_START", 4)
    tot_paie_row = row_map.get("TOT_PAIE", 178)
    compta_start = tot_paie_row + 3

    # Unmerge all in COMPTA region
    to_unmerge = [mr for mr in list(ws.merged_cells.ranges) if mr.min_row >= compta_start]
    for mr in to_unmerge:
        ws.unmerge_cells(str(mr))
    if to_unmerge:
        print(f"Unmerged {len(to_unmerge)} merged regions")

    tot_compta_row = None
    ecart_row = None

    if args.section in ("paie", "all"):
        paie_df = load_paie_data()
        build_paie_section(ws, paie_df, data_start, tot_paie_row)

    if args.section in ("compta", "all"):
        bg_map, _ = load_bg_section_a()
        gl_map     = load_gl_sections()
        tot_compta_row = build_compta_section(ws, bg_map, gl_map, compta_start)

    if args.section == "all" and tot_compta_row:
        ecart_row = build_ecart_row(ws, tot_compta_row, tot_paie_row)

    wb.save(FT_PATH)
    print(f"\n✅ Feuil2 saved: {FT_PATH}")

    # Update session
    session.setdefault("steps_completed", [])
    if "reconcile" not in session["steps_completed"]:
        session["steps_completed"].append("reconcile")

    session["feuil2_build"] = {
        "row_total_paie":   tot_paie_row,
        "row_total_compta": tot_compta_row,
        "row_ecart":        ecart_row,
        "source_section_A": "BG",
        "source_sections_BCD": "GL_all_journals",
    }
    with open(SESSION_FILE, "w", encoding="utf-8") as f:
        json.dump(session, f, indent=2, ensure_ascii=False)

    # ── Auto-trigger gap analysis (Correction 6) ──────────────────────────────
    if args.section == "all":
        run_mode = session.get("run_mode", "interactive")
        print(f"\n>>> Triggering gap analysis (mode: {run_mode})...")
        result = subprocess.run(
            [sys.executable, "scripts/build_feuil1_summary.py", "--write",
             "--row-paie", str(tot_paie_row),
             "--row-compta", str(tot_compta_row or ""),
             "--row-ecart",  str(ecart_row or ""),
             "--mode", run_mode],
            check=False
        )
        if result.returncode != 0:
            print("⚠️ Gap analysis (build_feuil1_summary) returned non-zero exit code — check output above")
        else:
            print("✅ Gap analysis and Feuil1 completed")

if __name__ == "__main__":
    main()
