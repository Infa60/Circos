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
OUTPUT_TXT   = r"C:\Circos_project\Circos_review\tasks_type.links.txt"
OUTPUT_SUMMARY = r"C:\Circos_project\Circos_review\tasks_type.numbers.txt"
OUTPUT_CHR = r"C:\Circos_project\Circos_review\tasks_type.data.txt"

SORT_WITHIN_SECTIONS = True
DEDUP_WITHIN_SECTIONS = True
Y1, Y2, Y3, Y4, Y5, Y6, Y7, Y8, Y9, Y10, Y11, Y12, Y13 = 230, 160, 100, 90, 50, 30, 30, 20, 20, 20, 10, 10, 10
# ============================

COL_ART = "ArtNb"
GMFCS_COLS = ["Sit-to-stand", "Running", "Cycling", "Stair-negotiation", "Obstacle-clearance", "TUG",
              "Game", "One-leg-standing", "Jumping", "Stepping-target", "Squat", "Hopping", "GMFM-E"
]

SECTION_MAP = {
    "Sit-to-stand": ("typeSit-to-stand", Y1, "vvdblue"),  # bleu très foncé
    "Running": ("typeRunning", Y2, "vdblue"),
    "Cycling": ("typeCycling", Y3, "dblue"),
    "Stair-negotiation": ("typeStair-negotiation", Y4, "blue"),
    "Obstacle-clearance": ("typeObstacle-clearance", Y5, "lblue"),
    "TUG": ("typeTUG", Y6, "dpblue"),
    "Game": ("typeGame", Y7, "vlblue"),
    "One-leg-standing": ("typeOne-leg-standing", Y8, "vvlblue"),
    "Jumping": ("typeJumping", Y9, "vlblue"),
    "Stepping-target": ("typeStepping-target", Y10, "vlblue"),
    "Squat": ("typeSquat", Y11, "vlblue"),
    "Hopping": ("typeHopping", Y12, "vlblue"),
    "GMFM-E": ("typeGMFM-E", Y13, "vvlblue"),
}

SECTIONS_ORDER = ["typeSit-to-stand", "typeRunning", "typeCycling", "typeStair-negotiation",
                  "typeObstacle-clearance", "typeTUG", "typeGame", "typeOne-leg-standing", "typeJumping",
                  "typeStepping-target", "typeSquat", "typeHopping", "typeGMFM-E"]

# Coordonnées pour le résumé
SUMMARY_Y = {
    "typeSit-to-stand":   Y1,
    "typeRunning":  Y2,
    "typeCycling": Y3,
    "typeStair-negotiation":  Y4,
    "typeObstacle-clearance": Y5,
    "typeTUG": Y6,
    "typeGame": Y7,
    "typeOne-leg-standing": Y8,
    "typeJumping": Y9,
    "typeStepping-target": Y10,
    "typeSquat": Y11,
    "typeHopping": Y12,
    "typeGMFM-E": Y13,
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
    color_I = SECTION_MAP["Sit-to-stand"][2]  # vvdyellow
    color_II = SECTION_MAP["Running"][2]  # vdyellow
    color_III = SECTION_MAP["Cycling"][2]  # dyellow
    color_IV = SECTION_MAP["Stair-negotiation"][2]  # dyellow
    color_V = SECTION_MAP["Obstacle-clearance"][2]  # dyellow
    color_VI = SECTION_MAP["TUG"][2]  # dyellow
    color_VII = SECTION_MAP["Game"][2]  # dyellow
    color_VIII = SECTION_MAP["One-leg-standing"][2]  # dyellow
    color_IX = SECTION_MAP["Jumping"][2]  # dyellow
    color_X = SECTION_MAP["Stepping-target"][2]  # dyellow
    color_XI = SECTION_MAP["Squat"][2]  # dyellow
    color_XII = SECTION_MAP["Hopping"][2]  # dyellow
    color_XIII = SECTION_MAP["GMFM-E"][2]  # dyellow


    #with out_chr.open("w", encoding="utf-8-sig", newline="") as fw:
    with out_chr.open("w", encoding="utf-8", newline="") as fw:
        fw.write("# chr - CHRNAME CHRLABEL START END COLOR\n")
        fw.write(f"chr -\ttypeSit-to-stand\tSit-to-stand\t0\t{Y1}\t{color_I}\n")
        fw.write(f"chr -\ttypeRunning\tRunning\t0\t{Y2}\t{color_II}\n")
        fw.write(f"chr -\ttypeCycling\tCycling\t0\t{Y3}\t{color_III}\n")
        fw.write(f"chr -\ttypeStair-negotiation\tStair-negotiation\t0\t{Y4}\t{color_IV}\n")
        fw.write(f"chr -\ttypeObstacle-clearance\tObstacle-clearance\t0\t{Y5}\t{color_V}\n")
        fw.write(f"chr -\ttypeTUG\tTUG\t0\t{Y5}\t{color_VI}\n")
        fw.write(f"chr -\ttypeGame\tGame\t0\t{Y7}\t{color_VII}\n")
        fw.write(f"chr -\ttypeOne-leg-standing\tOne-leg-standing\t0\t{Y8}\t{color_VIII}\n")
        fw.write(f"chr -\ttypeJumping\tJumping\t0\t{Y9}\t{color_IX}\n")
        fw.write(f"chr -\ttypeStepping-target\tStepping-target\t0\t{Y10}\t{color_X}\n")
        fw.write(f"chr -\ttypeSquat\tSquat\t0\t{Y11}\t{color_XI}\n")
        fw.write(f"chr -\ttypeHopping\tHopping\t0\t{Y12}\t{color_XII}\n")
        fw.write(f"chr -\ttypeGMFM-E\tGMFM-E\t0\t{Y13}\t{color_XIII}\n")

    print("✅ Fichier chr écrit :", out_chr)

    if errors:
        print(f"⚠️  {len(errors)} cellules vides signalées (voir section # ERRORS)")

if __name__ == "__main__":
    main()