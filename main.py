from dataclasses import dataclass
from typing import List, Dict, Tuple
from pathlib import Path
import pandas as pd
import re
from circos_make_articles_data import generate_articles_karyotype
from circos_conf_builder import generate_circos_conf

# ==========================================
#              USER CONFIGURATION
# ==========================================

# 1. PARAMÈTRES FICHIER EXCEL
EXCEL_PATH = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
SHEET_IDX = 0       # Index de la feuille
COL_ART = "ArtNb"   # Nom de la colonne ID article
COL_REF = "ref"     # Name of the Reference column

# 2. PARAMÈTRES DE SORTIE
OUTPUT_DIR = r"C:\Circos_project\Circos_review"

# 3. PARAMÈTRES D'ÉCHELLE (VISUALISATION)
# Taille visuelle min/max des sections (indépendant du nombre réel d'articles)
VISUAL_MIN_SIZE = 70
VISUAL_MAX_SIZE = 400


# ==========================================
#           CLASSES & UTILITAIRES
# ==========================================

def as_art_label(raw) -> str:
    """Normalise l'ID d'article -> 'artNN'."""
    s = "" if raw is None else str(raw).strip()
    if s == "": return ""
    m = re.match(r"^[Aa]rt\s*([0-9]+)$", s) or re.match(r"^([0-9]+)$", s)
    if m: return f"art{m.group(1)}"
    return s if s.lower().startswith("art") else f"art{s}"


def is_zero_like(val) -> bool:
    """Vaut True si la valeur est 0 ou vide."""
    if val is None: return False
    try:
        return float(str(val).strip().replace(",", ".")) == 0.0
    except:
        return False


NA_TOKENS = {"", "na", "n/a", "nan", "-", "--", "?", "??"}


@dataclass
class Section:
    excel_col: str  # Nom colonne Excel
    tlabel: str  # Identifiant Circos
    color: str  # Couleur


@dataclass
class TrackConfig:
    name: str  # Nom affichage
    sections: List[Section]
    subdir: str  # Sous-dossier sortie

    treat_empty_as_error: bool = True
    dedup: bool = True
    sort_in_section: bool = True
    special_na: Tuple[str, str] | None = None  # ("label_na", "couleur_na")

    @property
    def out_prefix(self):
        return str(Path(OUTPUT_DIR) / self.subdir)


# ==========================================
#           MOTEUR DE GÉNÉRATION
# ==========================================

def get_article_boundaries():
    """Trouve le premier et le dernier article (ex: art1, art53)."""
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_IDX, engine="openpyxl")
        valid_arts = []
        for val in df[COL_ART]:
            label = as_art_label(val)
            if label:
                valid_arts.append(label)

        def sort_key(x):
            m = re.match(r"^art(\d+)", x)
            return int(m.group(1)) if m else 0

        valid_arts.sort(key=sort_key)

        if valid_arts:
            return valid_arts[0], valid_arts[-1]
        return None, None
    except Exception as e:
        print(f"[ERREUR] Impossible de lire les articles: {e}")
        return None, None


def get_counts_for_config(cfg: TrackConfig) -> Dict[str, int]:
    """Compte le nombre d'articles par section pour calculer le min/max global."""
    section_map = {s.excel_col: s.tlabel for s in cfg.sections}
    sections_order = [s.tlabel for s in cfg.sections]

    tlabel_na = None
    if cfg.special_na:
        tlabel_na = cfg.special_na[0]
        if tlabel_na not in sections_order: sections_order.append(tlabel_na)

    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_IDX, engine="openpyxl")
    except Exception as e:
        print(f"[ERREUR] Lecture Excel ({cfg.name}): {e}")
        return {}

    cols_to_check = [s.excel_col for s in cfg.sections]
    bucket = {t: [] for t in sections_order}

    for _, row in df.iterrows():
        art = as_art_label(row.get(COL_ART))
        if not art: continue

        for col in cols_to_check:
            val = row.get(col)
            if val is None or (isinstance(val, float) and pd.isna(val)): continue
            sval = str(val).strip()
            low = sval.lower()

            if low == "???":
                if cfg.special_na: bucket[tlabel_na].append(art)
                continue

            if (isinstance(val, str) and low in NA_TOKENS): continue
            if is_zero_like(sval): continue

            tlabel = section_map[col]
            bucket[tlabel].append(art)

    counts = {}
    for t in sections_order:
        if cfg.dedup: bucket[t] = list(dict.fromkeys(bucket[t]))
        counts[t] = len(bucket[t])

    return counts


def build_track(cfg: TrackConfig, start_line, end_line, global_min, global_max):
    """Génère les fichiers Circos et retourne les bornes (premier, dernier label)."""

    section_map = {s.excel_col: (s.tlabel, s.color) for s in cfg.sections}
    sections_order = [s.tlabel for s in cfg.sections]

    tlabel_na, color_na = (None, None)
    if cfg.special_na:
        tlabel_na, color_na = cfg.special_na
        if tlabel_na not in sections_order: sections_order.append(tlabel_na)

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_IDX, engine="openpyxl")
    cols_to_check = [s.excel_col for s in cfg.sections]

    bucket = {t: [] for t in sections_order}
    errors = []

    # 1. Remplissage
    for _, row in df.iterrows():
        art = as_art_label(row.get(COL_ART))
        if not art:
            if cfg.treat_empty_as_error: errors.append(("(art vide)", COL_ART))
            continue

        for col in cols_to_check:
            val = row.get(col)
            if val is None or (isinstance(val, float) and pd.isna(val)):
                if cfg.treat_empty_as_error: errors.append((art, col))
                continue

            sval = str(val).strip()
            low = sval.lower()

            if low == "???":
                if cfg.special_na:
                    bucket[tlabel_na].append(
                        f"{art}\t{start_line}\t{end_line}\t{tlabel_na}\t0\tSIZE_PLACEHOLDER\tcolor={color_na}")
                elif cfg.treat_empty_as_error:
                    errors.append((art, col))
                continue

            if (isinstance(val, str) and low in NA_TOKENS):
                if cfg.treat_empty_as_error: errors.append((art, col))
                continue
            if is_zero_like(sval): continue

            tlabel, color = section_map[col]
            bucket[tlabel].append(f"{art}\t{start_line}\t{end_line}\t{tlabel}\t0\tSIZE_PLACEHOLDER\tcolor={color}")

    # 2. Calculs & Tris
    real_counts = {}
    for t in sections_order:
        if cfg.dedup: bucket[t] = list(dict.fromkeys(bucket[t]))
        if cfg.sort_in_section:
            def art_key(line: str):
                parts = line.split('\t')
                a = parts[0]
                m = re.match(r"^art(\d+)", a)
                return (0, int(m.group(1))) if m else (1, a.lower())

            bucket[t] = sorted(bucket[t], key=art_key)
        real_counts[t] = len(bucket[t])

    # 3. Mise à l'échelle & Bornes
    scaled_sizes = {}
    active_labels = []

    for t in sections_order:
        count = real_counts[t]
        if count == 0:
            scaled_sizes[t] = 0
            continue

        active_labels.append(t)

        if global_max == global_min:
            scaled_size = VISUAL_MAX_SIZE
        else:
            scaled_size = int(VISUAL_MIN_SIZE + (count - global_min) * (VISUAL_MAX_SIZE - VISUAL_MIN_SIZE) / (
                        global_max - global_min))

        scaled_sizes[t] = scaled_size

        new_lines = []
        for line in bucket[t]:
            new_lines.append(line.replace("SIZE_PLACEHOLDER", str(scaled_size)))
        bucket[t] = new_lines

    # 4. Écriture
    base_path = Path(OUTPUT_DIR) / cfg.subdir
    base_path.parent.mkdir(parents=True, exist_ok=True)

    f_links = base_path.with_name(f"{cfg.subdir}.links.txt")
    f_nums = base_path.with_name(f"{cfg.subdir}.numbers.txt")
    f_data = base_path.with_name(f"{cfg.subdir}.data.txt")

    with f_links.open("w", encoding="utf-8", newline="") as fw:
        for t in sections_order:
            if not bucket[t]: continue
            fw.write(f"# {t} (Real: {real_counts[t]}, Scaled: {scaled_sizes[t]})\n")
            fw.writelines(line + "\n" for line in bucket[t])
        if errors:
            fw.write("\n# ERRORS\n")
            for art, col in errors: fw.write(f"{art}\t{col}\t<empty>\n")

    with f_nums.open("w", encoding="utf-8", newline="") as fw:
        for t in sections_order:
            if scaled_sizes[t] > 0:
                fw.write(f"{t}\t0\t{scaled_sizes[t]}\t{real_counts[t]} color=black\n")

    with f_data.open("w", encoding="utf-8", newline="") as fw:
        fw.write("# chr - CHRNAME CHRLABEL START END COLOR\n")
        meta_info = {s.tlabel: (s.tlabel.replace("type", ""), s.color) for s in cfg.sections}
        if cfg.special_na: meta_info[tlabel_na] = (tlabel_na.replace("type", ""), color_na)

        for t in sections_order:
            if t not in meta_info or scaled_sizes[t] == 0: continue
            pretty, color = meta_info[t]
            fw.write(f"chr -\t{t}\t{pretty}\t0\t{scaled_sizes[t]}\t{color}\n")

    # Retourne les bornes pour le fichier de config
    if active_labels:
        return active_labels[0], active_labels[-1]
    return None, None


# ==========================================
#           CONFIGURATIONS DES BLOCS
# ==========================================

gmfcs_config = TrackConfig(
    name="GMFCS Level",
    subdir="gmfcs_level",
    special_na=("typeGMFCS-Unspecified", "255,150,150"),
    sections=[
        Section("GMFCS-I", "typeGMFCS-I", "120,0,0"),
        Section("GMFCS-II", "typeGMFCS-II", "160,20,20"),
        Section("GMFCS-III", "typeGMFCS-III", "200,40,40"),
        Section("GMFCS-IV", "typeGMFCS-IV", "230,80,80"),
    ]
)

cp_type_config = TrackConfig(
    name="CP Type",
    subdir="cp_type",
    special_na=("typeType-Unspecified", "255,160,80"),
    sections=[
        Section("Spastic", "typeSpastic", "120,45,0"),
        Section("Dyskinetic", "typeDyskinetic", "165,60,0"),
        Section("Ataxic", "typeAtaxic", "210,85,0"),
        Section("Mixed", "typeMixed", "240,120,30"),
    ]
)

laterality_config = TrackConfig(
    name="Laterality",
    subdir="cp_laterality",
    special_na=("typeLaterality-Unspecified", "255,230,80"),
    sections=[
        Section("Hemiplegic", "typeHemiplegic", "150,120,0"),
        Section("Diplegic", "typeDiplegic", "200,160,0"),
        Section("Quadriplegic", "typeQuadriplegic", "240,200,20"),
    ]
)

tools_config = TrackConfig(
    name="Assessment Tools",
    subdir="assessment_tools",
    sections=[
        Section("Optoelectronic", "typeOptoelectronic", "40,90,40"),
        Section("Force-plate", "typeForce-plate", "55,120,55"),
        Section("EMG", "typeEMG", "70,150,70"),
        Section("Heart-rate-monitor", "typeHeart-rate-monitor", "90,180,90"),
        Section("Indirect-calorimetry", "typeIndirect-calorimetry", "120,200,120"),
        Section("IMU", "typeIMU", "150,220,150"),
        Section("Wii-fit", "typeWii-fit", "180,240,180"),
        Section("Other-tools", "typeOther-tools", "210,255,210"),
    ]
)

assessment_type_config = TrackConfig(
    name="Assessment Type",
    subdir="assessment_type",
    sections=[
        Section("Spatiotemporal", "typeSpatiotemporal", "40,90,190"),
        Section("Kinematic", "typeKinematic", "55,120,210"),
        Section("Kinetic", "typeKinetic", "70,150,230"),
        Section("Stability", "typeStability", "90,170,240"),
        Section("Electromyographic", "typeElectromyographic", "120,190,250"),
        Section("Metabolic", "typeMetabolic", "170,215,255"),
    ]
)

# tasks_config = TrackConfig(
#     name="Task Type",
#     subdir="tasks_type",
#     sections=[
#         Section("Sit-to-stand", "typeSit-to-stand", "90,0,140"),
#         Section("Running", "typeRunning", "110,20,160"),
#         Section("Cycling", "typeCycling", "130,40,180"),
#         Section("Stair-negotiation", "typeStair-negotiation", "150,60,200"),
#         Section("Obstacle-clearance", "typeObstacle-clearance", "170,90,215"),
#         Section("Time-Up-and-Go", "typeTime-Up-and-Go", "190,120,225"),
#         Section("Game", "typeGame", "205,145,235"),
#         Section("One-leg-standing", "typeOne-leg-standing", "220,170,245"),
#         Section("Jumping", "typeJumping", "230,190,250"),
#         Section("Stepping-target", "typeStepping-target", "240,210,255"),
#         Section("Other-tasks", "typeOther-tasks", "250,230,255"),
#     ]
# )

tasks_config = TrackConfig(
    name="Task Type",
    subdir="tasks_type",
    sections=[
        Section("Sit-to-stand", "typeSit-to-stand", "50,0,80"),
        Section("Running", "typeRunning", "65,10,95"),
        Section("Cycling", "typeCycling", "80,20,110"),
        Section("Stair-negotiation", "typeStair-negotiation", "95,30,125"),
        Section("Obstacle-clearance", "typeObstacle-clearance", "110,40,140"),
        Section("Game", "typeGame", "125,55,155"),
        Section("Jumping", "typeJumping", "140,70,170"),
        Section("Time-Up-and-Go", "typeTime-Up-and-Go", "160,90,185"),
        Section("One-leg-standing", "typeOne-leg-standing", "180,110,200"),
        Section("Stepping-target", "typeStepping-target", "200,135,210"),
        Section("Hopping", "typeHopping", "220,160,220"),
        Section("Squat", "typeSquat", "240,190,230"),
        Section("Kicking-a-ball", "typeKicking-a-ball", "255,220,240"),
    ]
)

# ==========================================
#           EXÉCUTION PRINCIPALE
# ==========================================

if __name__ == "__main__":

    # 4. CHOIX DE L'ORDRE D'EXÉCUTION (MODIFIEZ L'ORDRE ICI)
    ACTIVE_TRACKS = [
        gmfcs_config,
        cp_type_config,
        laterality_config,
        tools_config,
        assessment_type_config,
        tasks_config
    ]

    print("\n=== PHASE 0: Generating Articles Karyotype ===")
    generate_articles_karyotype(
        excel_path=EXCEL_PATH,
        sheet_idx=SHEET_IDX,
        output_dir=OUTPUT_DIR,
        col_art=COL_ART,
        col_ref=COL_REF,
        end_value=60  # You can adjust default article size here
    )

    print("\n=== PHASE 1 : Analyse Globale (Calcul des Min/Max) ===")
    all_counts = []

    for cfg in ACTIVE_TRACKS:
        print(f"Scan de : {cfg.name}...")
        counts_dict = get_counts_for_config(cfg)
        valid_vals = [v for v in counts_dict.values() if v > 0]
        all_counts.extend(valid_vals)

    if not all_counts:
        GLOBAL_MIN, GLOBAL_MAX = 0, 1
        print("[INFO] Aucune donnée trouvée.")
    else:
        GLOBAL_MIN = min(all_counts)
        GLOBAL_MAX = max(all_counts)

    print(f"\n>>> BORNES GLOBALES : Min={GLOBAL_MIN}, Max={GLOBAL_MAX}")
    print(f">>> CIBLE VISUELLE : [{VISUAL_MIN_SIZE} - {VISUAL_MAX_SIZE}]\n")

    print("=== PHASE 2 : Génération & Récupération des Bornes ===")
    boundary_map = {}

    # Bornes des Articles
    first_art, last_art = get_article_boundaries()
    if first_art:
        print(f"Articles : {first_art} -> {last_art}")
        boundary_map['articles'] = (first_art, last_art)

    # Bornes des Tracks
    start_art_line = 0
    end_art_line = 9

    for cfg in ACTIVE_TRACKS:
        print(f"Traitement : {cfg.name}")
        first_lbl, last_lbl = build_track(cfg, start_art_line, end_art_line, GLOBAL_MIN, GLOBAL_MAX)

        if first_lbl and last_lbl:
            boundary_map[cfg.subdir] = (first_lbl, last_lbl)

        start_art_line = end_art_line + 1
        end_art_line = start_art_line + 9

    print("\n=== PHASE 3 : Création automatique de circos.conf ===")
    generate_circos_conf(
        output_dir=OUTPUT_DIR,
        active_tracks=ACTIVE_TRACKS,
        boundary_map=boundary_map
    )

    print("\n✅ Terminé avec succès.")