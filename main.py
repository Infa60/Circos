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

def build_track(cfg: TrackConfig, start_line, end_line):
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
                    bucket[tlabel_na].append(f"{art}\t{start_line}\t{end_line}\t{tlabel_na}\t0\t{y_na}\tcolor={color_na}")
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
            bucket[tlabel].append(f"{art}\t{start_line}\t{end_line}\t{tlabel}\t0\t{y}\tcolor={color}")

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
        Section("GMFCS-I",   "typeGMFCS-I",   374, "120,0,0"),
        Section("GMFCS-II",  "typeGMFCS-II",  400,  "160,20,20"),
        Section("GMFCS-III", "typeGMFCS-III", 203,  "200,40,40"),
        Section("GMFCS-IV",  "typeGMFCS-IV",  95,  "230,80,80"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\gmfcs_level",
    special_na = ("typeGMFCS-Unspecified", 139, "255,150,150"),  # "???" traités ici

)

# 2️) CP TYPE (orange)
cp_type_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Spastic",     "typeSpastic",     349, "120,45,0"),
        Section("Dyskinetic",  "typeDyskinetic",  89, "165,60,0"),
        Section("Ataxic",      "typeAtaxic",      82,  "210,85,0"),
        Section("Mixed",       "typeMixed",       70,  "240,120,30"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\cp_type",
    special_na = ("typeType-Unspecified", 254, "255,160,80"),  # "???" traités ici

)

# 3️) LATERALITY (jaune)
laterality_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Hemiplegic", "typeHemiplegic", 317, "150,120,0"),
        Section("Diplegic",  "typeDiplegic",  393, "200,160,0"),
        Section("Quadriplegic", "typeQuadriplegic", 108, "240,200,20"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\cp_laterality",
    special_na = ("typeLaterality-Unspecified", 108, "255,230,80"),  # "???" traités ici

)

# 4️) ASSESSMENT TOOL (vert)
assessment_tools_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[
        Section("Optoelectronique",       "typeOptoelectronique",       355, "40,90,40"),
        Section("Force-plate",            "typeForce-plate",            298, "55,120,55"),
        Section("EMG",                    "typeEMG",                    139, "70,150,70"),
        Section("Heart-rate-monitor",       "typeHeart-rate-monitor",       108,   "90,180,90"),
        Section("Indirect-calorimetry",     "typeIndirect-calorimetry",     108,  "120,200,120"),
        Section("IMU",                       "typeIMU",                     101,  "150,220,150"),
        Section("Wii-fit",                  "typeWii-fit",                  76,   "180,240,180"),
        Section("Other-tools",            "typeOther-tools",            108,  "210,255,210"),
    ],
    out_prefix=r"C:\Circos_project\Circos_review\assessment_tools"
)

# 5) ASSESSMENT TYPE (violet)
assessment_type_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[

        Section("Spatiotemporal",    "typeSpatiotemporal",    374, "40,90,190"),    # bleu soutenu mais lisible
        Section("Kinematic",         "typeKinematic",         355, "55,120,210"),   # bleu moyen
        Section("Kinetic",           "typeKinetic",           215, "70,150,230"),   # bleu clair équilibré
        Section("Stability",         "typeStability",         165, "90,170,240"),   # bleu lumineux
        Section("Electromyographic", "typeElectromyographic", 139, "120,190,250"),  # bleu clair
        Section("Metabolic",         "typeMetabolic",         127, "170,215,255" ), # bleu très clair / pastel

    ],
    out_prefix=r"C:\Circos_project\Circos_review\assessment_type"
)

# 6️) TASK TYPE (bleu)
task_type_cfg = TrackConfig(
    input_xlsx=EXCEL_PATH,
    sheet=0,
    col_art="ArtNb",
    sections=[


        #Section("Sit-to-stand",      "typeSit-to-stand",      209, "0,82,170"),   # foncé mais pas trop
        #Section("Running",           "typeRunning",           165, "0,95,190"),   # un peu plus clair
        #Section("Cycling",           "typeCycling",           127, "0,110,210"),  # bleu moyen
        #Section("Stair-negotiation", "typeStair-negotiation", 120, "0,125,230"),  # bleu vif
        #Section("Obstacle-clearance","typeObstacle-clearance", 95,  "40,145,245"),# bleu lumineux
        #Section("TUG",               "typeTUG",                82,  "80,165,255"),# bleu clair
        #Section("Game",              "typeGame",               82,  "120,185,255"),# bleu doux
        #Section("One-leg-standing",  "typeOne-leg-standing",   76,  "155,205,255"),# bleu pastel
        #Section("Jumping",           "typeJumping",            76,  "185,220,255"),# bleu pâle
        #Section("Stepping-target",   "typeStepping-target",    76,  "210,235,255"),# bleu très pâle
        #Section("Other-tasks",       "typeOther-tasks",        82,  "235,245,255"),# presque blanc

        Section("Sit-to-stand",      "typeSit-to-stand",      209, "90,0,140"),    # violet profond
        Section("Running",           "typeRunning",           165, "110,20,160"),  # violet soutenu
        Section("Cycling",           "typeCycling",           120, "130,40,180"),  # violet moyen
        Section("Stair-negotiation", "typeStair-negotiation", 120, "150,60,200"),  # violet vif
        Section("Obstacle-clearance","typeObstacle-clearance", 95,  "170,90,215"), # violet lumineux
        Section("TUG",               "typeTUG",                82,  "190,120,225"),# violet clair
        Section("Game",              "typeGame",               82,  "205,145,235"),# mauve doux
        Section("One-leg-standing",  "typeOne-leg-standing",   76,  "220,170,245"),# lavande pastel
        Section("Jumping",           "typeJumping",            76,  "230,190,250"),# lavande pâle
        Section("Stepping-target",   "typeStepping-target",    76,  "240,210,255"),# mauve très pâle
        Section("Other-tasks",       "typeOther-tasks",        82,  "250,230,255"),# presque blanc rosé


        #Section("Squat",             "typeSquat",             10,  "vlblue"),
        #Section("Hopping",           "typeHopping",           10,  "vlblue"),
        #Section("GMFM-E",            "typeGMFM-E",            10,  "vvlblue"),
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

    start_art_line = 0
    end_art_line = 9

    for cfg in all_cfgs:
        print(f"\n=== Construction du track : {cfg.out_prefix} ===")
        build_track(cfg, start_art_line, end_art_line)
        start_art_line = end_art_line + 1
        end_art_line = start_art_line + 9
