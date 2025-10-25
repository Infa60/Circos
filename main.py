from dataclasses import dataclass
from typing import List, Dict, Tuple
from pathlib import Path
import pandas as pd
import re
from make_articles_data import main as build_articles

# ===================== Utils =====================

def as_art_label(raw) -> str:
    """Normalise l’ID d’article -> 'artNN'."""
    s = "" if raw is None else str(raw).strip()
    if s == "":
        return ""
    m = re.match(r"^[Aa]rt\s*([0-9]+)$", s) or re.match(r"^([0-9]+)$", s)
    if m:
        return f"art{m.group(1)}"
    return s if s.lower().startswith("art") else f"art{s}"

def is_zero_like(val) -> bool:
    """Vaut True si val est un 0 (0, 0.0, '0,0', etc.)."""
    if val is None:
        return False
    try:
        return float(str(val).strip().replace(",", ".")) == 0.0
    except:
        return False

NA_TOKENS = {"", "na", "n/a", "nan", "-", "--", "?", "??"}  # '???' traité à part si special_na

# ===================== Modèles de config =====================

@dataclass
class Section:
    excel_col: str      # nom exact de la colonne Excel
    tlabel: str         # identifiant Circos (sans espace)
    y: int              # longueur/position Y (source unique de vérité)
    color: str          # couleur Circos (nom reconnu)

@dataclass
class TrackConfig:
    input_xlsx: str
    sheet: int
    col_art: str
    sections: List[Section]
    out_prefix: str
    # options
    treat_empty_as_error: bool = True
    dedup: bool = True
    sort_in_section: bool = True
    # Cas spécial: valeur textuelle "???"
    # Si défini -> ("tlabel_na", y_na, color_na) pour router les "???"
    special_na: Tuple[str, int, str] | None = None

# ===================== Moteur générique =====================

def build_track(cfg: TrackConfig):
    # Dérivés automatiques (pas de duplication)
    cols = [s.excel_col for s in cfg.sections]
    section_map: Dict[str, Tuple[str, int, str]] = {s.excel_col: (s.tlabel, s.y, s.color) for s in cfg.sections}
    sections_order: List[str] = [s.tlabel for s in cfg.sections]
    summary_y: Dict[str, int] = {s.tlabel: s.y for s in cfg.sections}

    # Ajouter la section spéciale NA (si présente) à l’ordre/summary
    if cfg.special_na:
        tlabel_na, y_na, color_na = cfg.special_na
        if tlabel_na not in sections_order:
            sections_order.append(tlabel_na)
            summary_y[tlabel_na] = y_na

    # Lecture Excel
    try:
        df = pd.read_excel(cfg.input_xlsx, sheet_name=cfg.sheet, engine="openpyxl")
    except Exception as e:
        raise SystemExit(f"Erreur lecture Excel : {e}")

    missing = [c for c in [cfg.col_art] + cols if c not in df.columns]
    if missing:
        raise SystemExit(f"Colonnes manquantes : {missing}\nColonnes trouvées : {list(df.columns)}")

    bucket: Dict[str, List[str]] = {t: [] for t in sections_order}
    errors: List[Tuple[str, str]] = []

    # Remplissage
    for _, row in df.iterrows():
        art = as_art_label(row.get(cfg.col_art))
        if not art:
            if cfg.treat_empty_as_error:
                errors.append(("(art vide)", cfg.col_art))
            continue

        for col in cols:
            val = row.get(col)

            # Détection NA générales
            if val is None or (isinstance(val, float) and pd.isna(val)):
                if cfg.treat_empty_as_error:
                    errors.append((art, col))
                continue

            sval = str(val).strip()
            low = sval.lower()

            # "???" -> route vers special_na si défini
            if low == "???":
                if cfg.special_na:
                    tlabel_na, y_na, color_na = cfg.special_na
                    bucket[tlabel_na].append(f"{art}\t0\t25\t{tlabel_na}\t0\t{y_na}\tcolor={color_na}")
                else:
                    if cfg.treat_empty_as_error:
                        errors.append((art, col))
                continue

            # Autres NA textuels -> considérer vide
            if (isinstance(val, str) and low in NA_TOKENS):
                if cfg.treat_empty_as_error:
                    errors.append((art, col))
                continue

            # 0 -> ignorer
            if is_zero_like(sval):
                continue

            # Sinon -> on compte la présence dans la section
            tlabel, y, color = section_map[col]
            bucket[tlabel].append(f"{art}\t0\t25\t{tlabel}\t0\t{y}\tcolor={color}")

    # Dédupe
    if cfg.dedup:
        for t in sections_order:
            bucket[t] = list(dict.fromkeys(bucket[t]))

    # Tri interne
    if cfg.sort_in_section:
        def art_key(line: str):
            m = re.match(r"^art(\d+)", line)
            return (0, int(m.group(1))) if m else (1, line.lower())
        for t in sections_order:
            bucket[t] = sorted(bucket[t], key=art_key)

    # Chemins sortie
    out_links = Path(f"{cfg.out_prefix}.links.txt")
    out_nums  = Path(f"{cfg.out_prefix}.numbers.txt")
    out_chr   = Path(f"{cfg.out_prefix}.data.txt")
    for p in (out_links, out_nums, out_chr):
        p.parent.mkdir(parents=True, exist_ok=True)

    # .links
    with out_links.open("w", encoding="utf-8", newline="") as fw:
        first = True
        for t in sections_order:
            lines = bucket[t]
            if not lines:
                continue
            if not first:
                fw.write("\n")
            fw.write(f"# {t}\n")
            fw.writelines(line + "\n" for line in lines)
            first = False
        if errors:
            fw.write("\n# ERRORS (cellules vides)\n")
            for art, col in errors:
                fw.write(f"{art}\t{col}\t<empty>\n")

    # .numbers
    with out_nums.open("w", encoding="utf-8", newline="") as fw:
        for t in sections_order:
            y = summary_y[t]
            n = len(bucket[t])
            fw.write(f"{t}\t0\t{y}\t{n} color=black\n")

    # .data (karyotype)
    with out_chr.open("w", encoding="utf-8", newline="") as fw:
        fw.write("# chr - CHRNAME CHRLABEL START END COLOR\n")
        # sections normales
        for s in cfg.sections:
            #pretty = re.sub(r"[_-]+", " ", s.tlabel.replace("type", ""))
            pretty = s.tlabel.replace("type", "")
            fw.write(f"chr -\t{s.tlabel}\t{pretty}\t0\t{s.y}\t{s.color}\n")
        # section spéciale NA si définie
        if cfg.special_na:
            tlabel_na, y_na, color_na = cfg.special_na
            #pretty_na = re.sub(r"[_-]+", " ", tlabel_na.replace("type", ""))
            pretty_na = tlabel_na.replace("type", "")
            fw.write(f"chr -\t{tlabel_na}\t{pretty_na}\t0\t{y_na}\t{color_na}\n")

    print("✅ Écrit :", out_links, out_nums, out_chr)

# ===================== CONFIGS =====================

EXCEL_PATH = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
SHEET_IDX = 0
COL_ART = "ArtNb"

# 1) GMFCS LEVEL (rouge)
gmfcs_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=SHEET_IDX,
    col_art=COL_ART,
    sections=[
        Section("GMFCS-I",   "typeGMFCS-I",   490, "vvdred"),
        Section("GMFCS-II",  "typeGMFCS-II",  530,  "vdred"),
        Section("GMFCS-III", "typeGMFCS-III", 220,  "dred"),
        Section("GMFCS-IV",  "typeGMFCS-IV",  90,  "red"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\gmfcs_level",
    special_na = ("typeGMFCS-Unspecified", 120, "vlred"),  # "???" traités ici

)


# 2️) CP TYPE (orange)
cp_type_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Spastic",     "typeSpastic",     460, "vvdorange"),
        Section("Dyskinetic",  "typeDyskinetic",  40, "vdorange"),
        Section("Ataxic",      "typeAtaxic",      30,  "dorange"),
        Section("Mixed",       "typeMixed",       10,  "orange"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\cp_type",
    special_na = ("typeType-Unspecified", 300, "lorange"),  # "???" traités ici

)

# 3️) LATERALITY (jaune)
laterality_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Hemiplegic", "typeHemiplegic", 400, "vvdyellow"),
        Section("Diplegic",  "typeDiplegic",  530, "dyellow"),
        Section("Quadriplegic", "typeQuadriplegic", 70, "yellow"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\cp_laterality",
    special_na = ("typeLaterality-Unspecified", 70, "lyellow"),  # "???" traités ici

)

# 4️) ASSESSMENT TOOL (vert)
assessment_tools_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Optoelectronique",       "typeOptoelectronique",       470, "vvdgreen"),
        Section("Force-plate",            "typeForce_plate",            370, "vdgreen"),
        Section("IMU",                    "typeIMU",                    60,  "dgreen"),
        Section("EMG",                    "typeEMG",                    130, "green"),
        Section("Wii-fit",                "typeWii_fit",                20,  "dpgreen"),
        Section("Heart-rate-monitor",     "typeHeart_rate_monitor",     70,  "lgreen"),
        Section("Indirect-calorimetry",   "typeIndirect_calorimetry",   70,  "vlgreen"),
        Section("Other",                  "typeOther",                  70,  "vvlgreen"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\assessment_tools"
)

# 5) ASSESSMENT TYPE (violet)
assessment_type_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Spatiotemporal",    "typeSpatiotemporal",    490, "vvdpurple"),
        Section("Kinematic",         "typeKinematic",         470, "vdpurple"),
        Section("Kinetic",           "typeKinetic",           240, "dpurple"),
        Section("Electromyographic", "typeElectromyographic", 120, "purple"),
        Section("Metabolic",         "typeMetabolic",         100, "lpurple"),
        Section("Stability",         "typeStability",         160, "vlpurple"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\assessment_type"
)

# 6️) TASK TYPE (bleu)
task_type_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Sit-to-stand",      "typeSit-to-stand",      230, "vvdblue"),
        Section("Running",           "typeRunning",           160, "vdblue"),
        Section("Cycling",           "typeCycling",           100, "dblue"),
        Section("Stair-negotiation", "typeStair-negotiation", 90,  "blue"),
        Section("Obstacle-clearance","typeObstacle-clearance",50,  "lblue"),
        Section("TUG",               "typeTUG",               30,  "dpblue"),
        Section("Game",              "typeGame",              30,  "vlblue"),
        Section("One-leg-standing",  "typeOne-leg-standing",  20,  "vvlblue"),
        Section("Jumping",           "typeJumping",           20,  "vlblue"),
        Section("Stepping-target",   "typeStepping-target",   20,  "vlblue"),
        Section("Squat",             "typeSquat",             10,  "vlblue"),
        Section("Hopping",           "typeHopping",           10,  "vlblue"),
        Section("GMFM-E",            "typeGMFM-E",            10,  "vvlblue"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\tasks_type"
)

# ---------- Lancement global ----------
if __name__ == "__main__":
    all_cfgs = [
        gmfcs_cfg,
        cp_type_cfg,
        laterality_cfg,
        assessment_tools_cfg,
        assessment_type_cfg,
        task_type_cfg,
    ]
    print("\n=== Génération du fichier articles.data.txt ===")
    build_articles()

    for cfg in all_cfgs:
        print(f"\n=== Construction du track : {cfg.out_prefix} ===")
        build_track(cfg)
