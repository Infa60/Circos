"""
Lit un Excel (.xlsx) avec colonnes:
  - ArtNb
  - Hemiplegic, Diplegic, Quadriplegic, mixte

RÈGLES PAR CELLULE :
- "0" (ou 0.0, 0,00) -> ne rien écrire.
- "???" -> écrire une ligne sous typeLateralityNA (Y=170, color=lyellow) pour cette cellule (pas de typeHemiplegic..IV).
- "X" ou toute valeur != 0 -> écrire sous le type de la colonne :
    Hemiplegic -> typeHemiplegic   (Y=280, color=vvdyellow)
    Diplegic -> typeDiplegic  (Y=470, color=vdyellow)
    Quadriplegic -> typeQuadriplegic (Y=170, color=dyellow)
    mixte -> typeMixte  (Y=170, color=yellow)
- Cellule vide -> reporter en # ERRORS.

Sortie TXT (UTF-8 BOM) groupée par sections :
  # typeHemiplegic
  # typeDiplegic
  # typeQuadriplegic
  # typeMixte
  # typeLateralityNA
  # ERRORS
"""

from pathlib import Path
import re

# ========== CONFIG ==========
INPUT_XLSX   = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
SHEET_NAME   = 0
OUTPUT_TXT   = r"C:\Circos_project\Circos_review\assessment_tools.links.txt"
OUTPUT_SUMMARY = r"C:\Circos_project\Circos_review\assessment_tools.numbers.txt"
OUTPUT_CHR = r"C:\Circos_project\Circos_review\assessment_tools.data.txt"

SORT_WITHIN_SECTIONS = True
DEDUP_WITHIN_SECTIONS = True
Y1, Y2, Y3, Y4, Y5, Y6, Y7, Y8 = 470, 370, 60, 130, 20, 70, 70, 70
# ============================

COL_ART = "ArtNb"
GMFCS_COLS = ["Optoelectronique", "Force-plate", "IMU", "EMG", "Wii-fit", "Heart-rate-monitor",
              "Indirect-calorimetry", "Autres"]

SECTION_MAP = {
    "Optoelectronique": ("typeOptoelectronique", Y1, "vvdgreen"),  # vert très foncé
    "Force-plate": ("typeForce_plate", Y2, "vdgreen"),
    "IMU": ("typeIMU", Y3, "dgreen"),
    "EMG": ("typeEMG", Y4, "green"),
    "Wii-fit": ("typeWii_fit", Y5, "dpgreen"),
    "Heart-rate-monitor": ("typeHeart_rate_monitor", Y6, "lgreen"),
    "Indirect-calorimetry": ("typeIndirect_calorimetry", Y7, "vlgreen"),
    "Autres": ("typeAutres", Y8, "vvlgreen"),
}

SECTIONS_ORDER = ["typeOptoelectronique", "typeForce_plate", "typeIMU", "typeEMG", "typeWii_fit",
                  "typeHeart_rate_monitor", "typeIndirect_calorimetry", "typeAutres"]

# Coordonnées pour le résumé
SUMMARY_Y = {
    "typeOptoelectronique":   Y1,
    "typeForce_plate":  Y2,
    "typeIMU": Y3,
    "typeEMG":  Y4,
    "typeWii_fit": Y5,
    "typeHeart_rate_monitor": Y6,
    "typeIndirect_calorimetry": Y7,
    "typeAutres": Y8
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

            # "???" => typeLateralityNA
            #if low == "???":
                #tlabel, y, color = TYPE_NA
                #bucket[tlabel].append(f"{art}\t0\t25\t{tlabel}\t0\t{y}\tcolor={color}")
                #continue

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
    color_I = SECTION_MAP["Optoelectronique"][2]  # vvdyellow
    color_II = SECTION_MAP["Force-plate"][2]  # vdyellow
    color_III = SECTION_MAP["IMU"][2]  # dyellow
    color_IV = SECTION_MAP["EMG"][2]  # dyellow
    color_V = SECTION_MAP["Wii-fit"][2]  # dyellow
    color_VI = SECTION_MAP["Heart-rate-monitor"][2]  # dyellow
    color_VII = SECTION_MAP["Indirect-calorimetry"][2]  # dyellow
    color_VIII = SECTION_MAP["Autres"][2]  # dyellow

    #with out_chr.open("w", encoding="utf-8-sig", newline="") as fw:
    with out_chr.open("w", encoding="utf-8", newline="") as fw:
        fw.write("# chr - CHRNAME CHRLABEL START END COLOR\n")
        fw.write(f"chr -\ttypeOptoelectronique\tOptoelectronique\t0\t{Y1}\t{color_I}\n")
        fw.write(f"chr -\ttypeForce_plate\tForce-plate\t0\t{Y2}\t{color_II}\n")
        fw.write(f"chr -\ttypeIMU\tIMU\t0\t{Y3}\t{color_III}\n")
        fw.write(f"chr -\ttypeEMG\tEMG\t0\t{Y4}\t{color_IV}\n")
        fw.write(f"chr -\ttypeWii_fit\tWii-fit\t0\t{Y5}\t{color_V}\n")
        fw.write(f"chr -\ttypeHeart_rate_monitor\tHeart-rate-monitor\t0\t{Y6}\t{color_VI}\n")
        fw.write(f"chr -\ttypeIndirect_calorimetry\tIndirect-calorimetry\t0\t{Y7}\t{color_VII}\n")
        fw.write(f"chr -\ttypeAutres\tAutres\t0\t{Y8}\t{color_VIII}\n")

    print("✅ Fichier chr écrit :", out_chr)


    if errors:
        print(f"⚠️  {len(errors)} cellules vides signalées (voir section # ERRORS)")

if __name__ == "__main__":
    main()