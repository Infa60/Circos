"""
Lit un Excel (.xlsx) avec colonnes:
  - ArtNb
  - spastic, ataxic, dyskinetic, mixte

RÈGLES PAR CELLULE :
- "0" (ou 0.0, 0,00) -> ne rien écrire.
- "???" -> écrire une ligne sous typeCPNA (Y=170, color=lorange) pour cette cellule (pas de typeSpastic..IV).
- "X" ou toute valeur != 0 -> écrire sous le type de la colonne :
    spastic -> typeSpastic   (Y=280, color=vvdorange)
    ataxic -> typeAtaxic  (Y=470, color=vdorange)
    dyskinetic -> typeDyskinetic (Y=170, color=dorange)
    mixte -> typeMixte  (Y=170, color=orange)
- Cellule vide -> reporter en # ERRORS.

Sortie TXT (UTF-8 BOM) groupée par sections :
  # typeSpastic
  # typeAtaxic
  # typeDyskinetic
  # typeMixte
  # typeCPNA
  # ERRORS
"""

from pathlib import Path
import re

# ========== CONFIG ==========
INPUT_XLSX   = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
SHEET_NAME   = 0
OUTPUT_TXT   = r"C:\Circos_project\Circos_review\cp_type.links.txt"
OUTPUT_SUMMARY = r"C:\Circos_project\Circos_review\cp_type.numbers.txt"
OUTPUT_CHR = r"C:\Circos_project\Circos_review\cp_type.data.txt"

SORT_WITHIN_SECTIONS = True
DEDUP_WITHIN_SECTIONS = True
Y1, Y2, Y3, Y4, YNA = 460, 30, 40, 10, 300
# ============================

COL_ART = "ArtNb"
GMFCS_COLS = ["Spastic", "Ataxic", "Dyskinetic", "Mixte"]

SECTION_MAP = {
    "Spastic": ("typeSpastic",   Y1, "vvdorange"),
    "Ataxic": ("typeAtaxic",  Y2, "vdorange"),
    "Dyskinetic": ("typeDyskinetic", Y3, "dorange"),
    "Mixte": ("typeMixte",  Y4, "orange"),
}
TYPE_NA = ("typeCPNA", YNA, "lorange")
SECTIONS_ORDER = ["typeSpastic", "typeAtaxic", "typeDyskinetic", "typeMixte", "typeCPNA"]

# Coordonnées pour le résumé
SUMMARY_Y = {
    "typeSpastic":   Y1,
    "typeAtaxic":  Y2,
    "typeDyskinetic": Y3,
    "typeMixte":  Y4,
    "typeCPNA":  YNA,
}

def as_art_label(raw) -> str:
    s = "" if raw is None else str(raw).strip()
    if s == "":
        return ""
    m = re.match(r"^[Aa]rt\s*([0-9]+)$", s)
    if m:
        return f"art{m.group(1)}"
    m = re.match(r"^([0-9]+)$", s)
    if m:
        return f"art{m.group(1)}"
    return s if s.lower().startswith("art") else f"art{s}"

def is_zero_like(s: str) -> bool:
    txt = s.strip().replace(",", ".")
    try:
        return float(txt) == 0.0
    except ValueError:
        return False

def main():
    try:
        import pandas as pd
    except ImportError:
        raise SystemExit("Ce script requiert pandas et openpyxl.\nInstalle :  pip install pandas openpyxl")
    try:
        df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME, engine="openpyxl")
    except Exception as e:
        raise SystemExit(f"Erreur lecture Excel : {e}")

    missing = [c for c in [COL_ART] + GMFCS_COLS if c not in df.columns]
    if missing:
        raise SystemExit(f"Colonnes manquantes : {missing}\nColonnes trouvées : {list(df.columns)}")

    bucket = {sec: [] for sec in SECTIONS_ORDER}
    errors = []

    print(f"📄 Lignes lues (hors entête) : {len(df)}")

    for idx, row in df.iterrows():
        art = as_art_label(row.get(COL_ART))
        if not art:
            errors.append(("(art vide)", COL_ART))
            continue

        for col in GMFCS_COLS:
            val = row.get(col)
            if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
                errors.append((art, col))
                continue

            sval = str(val).strip()
            low = sval.lower()

            # "???" => typeCPNA
            if low == "???":
                tlabel, y, color = TYPE_NA
                bucket[tlabel].append(f"{art}\t0\t25\t{tlabel}\t0\t{y}\tcolor={color}")
                continue

            if is_zero_like(sval):
                continue

            tlabel, y, color = SECTION_MAP[col]
            bucket[tlabel].append(f"{art}\t0\t25\t{tlabel}\t0\t{y}\tcolor={color}")

    # Dédoublonnage
    if DEDUP_WITHIN_SECTIONS:
        for sec in SECTIONS_ORDER:
            bucket[sec] = list(dict.fromkeys(bucket[sec]))

    # Tri
    if SORT_WITHIN_SECTIONS:
        def art_key(line: str):
            m = re.match(r"^art(\d+)", line)
            return (0, int(m.group(1))) if m else (1, line.lower())
        for sec in SECTIONS_ORDER:
            bucket[sec] = sorted(bucket[sec], key=art_key)

    # Écriture du fichier principal
    out_path = Path(OUTPUT_TXT)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="") as fw:

        first = True
        for sec in SECTIONS_ORDER:
            lines = bucket[sec]
            if not lines:
                continue
            if not first:
                fw.write("\n")
            fw.write(f"# {sec}\n")
            fw.writelines(line + "\n" for line in lines)
            first = False
        if errors:
            fw.write("\n# ERRORS (cellules vides)\n")
            for art, col in errors:
                fw.write(f"{art}\t{col}\t<empty>\n")

    counts = {sec: len(bucket[sec]) for sec in SECTIONS_ORDER}
    print("✅ Fichier écrit :", out_path)
    print("  Lignes par section :", counts)

    # === AJOUT : création du fichier résumé ===
    out_sum = Path(OUTPUT_SUMMARY)
    out_sum.parent.mkdir(parents=True, exist_ok=True)
    with out_sum.open("w", encoding="utf-8", newline="") as fw:
        for sec in SECTIONS_ORDER:
            y = SUMMARY_Y[sec]
            n = counts.get(sec, 0)
            fw.write(f"{sec}\t0\t{y}\t{n} color=black\n")

    print("✅ Fichier résumé écrit :", out_sum)

    # === AJOUT : création du fichier CHR ===
    out_chr = Path(OUTPUT_CHR)
    out_chr.parent.mkdir(parents=True, exist_ok=True)

    # Récupère les couleurs depuis tes mappings existants
    color_I = SECTION_MAP["Spastic"][2]  # vvdorange
    color_II = SECTION_MAP["Ataxic"][2]  # vdorange
    color_III = SECTION_MAP["Dyskinetic"][2]  # dorange
    color_IV = SECTION_MAP["Mixte"][2]  # orange
    color_NA = TYPE_NA[2]  # lorange

    #with out_chr.open("w", encoding="utf-8-sig", newline="") as fw:
    with out_chr.open("w", encoding="utf-8", newline="") as fw:
        fw.write("# chr - CHRNAME CHRLABEL START END COLOR\n")
        fw.write(f"chr -\ttypeSpastic\tSpastic\t0\t{Y1}\t{color_I}\n")
        fw.write(f"chr -\ttypeAtaxic\tAtaxic\t0\t{Y2}\t{color_II}\n")
        fw.write(f"chr -\ttypeDyskinetic\tDyskinetic\t0\t{Y3}\t{color_III}\n")
        fw.write(f"chr -\ttypeMixte\tMixte\t0\t{Y4}\t{color_IV}\n")
        fw.write(f"chr -\ttypeCPNA\tNA\t0\t{YNA}\t{color_NA}\n")

    print("✅ Fichier chr écrit :", out_chr)

    if errors:
        print(f"⚠️  {len(errors)} cellules vides signalées (voir section # ERRORS)")

if __name__ == "__main__":
    main()