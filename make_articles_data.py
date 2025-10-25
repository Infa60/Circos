"""
Lit un Excel et écrit un fichier karyotype Circos :
  chr -  art<ArtNb>  <ref>  0  100  black

- Colonne "ref" : libellé complet (ex: "Abdolrahmani et al. 2017")
- Colonne "ArtNb" : numéro d'article (ex: 1, 2, 3)

Option : remplace les espaces et caractères spéciaux par "_"
ex : "Abdolrahmani et al. 2017" -> "Abdolrahmani__et_al._2017"

Sortie : UTF-8 (sans BOM), séparateur = tabulation.
Tri croissant sur le numéro d’article.
"""

from pathlib import Path
import re

# ========== À MODIFIER ==========
INPUT_XLSX  = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
SHEET_NAME  = 0
OUTPUT_TXT  = r"C:\Circos_project\Circos_review\articles.data.txt"
END_VALUE   = 100
COLOR_VALUE = "black"
COL_REF     = "ref"
COL_ARTNB   = "ArtNb"
REPLACE_SPACES = True
# ================================


def normalize_ref(ref: str) -> str:
    """Nettoie et transforme le texte du champ ref."""
    s = str(ref).strip()
    if not s:
        return s
    if REPLACE_SPACES:
        s = re.sub(r"[^\w\.]", "_", s)   # garde lettres et points
        s = re.sub(r"_+", "_", s)       # compresse plusieurs underscores
    return s


def art_label_from(raw) -> str:
    """Construit 'artN' à partir de la valeur ArtNb."""
    if raw is None:
        return ""
    s = str(raw).strip()
    if s == "":
        return ""
    m = re.match(r"^[Aa]rt\s*([0-9]+)$", s)
    if m:
        return f"art{m.group(1)}"
    m = re.match(r"^([0-9]+)$", s)
    if m:
        return f"art{m.group(1)}"
    return s if s.lower().startswith("art") else f"art{s}"


def art_number(raw) -> int:
    """Extrait le numéro d'article sous forme d'entier pour trier."""
    s = str(raw).strip()
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else 999999  # grand nombre si non numérique


def main():
    try:
        import pandas as pd
    except ImportError:
        raise SystemExit("Installe d'abord pandas et openpyxl :\n  pip install pandas openpyxl")

    try:
        df = pd.read_excel(
            INPUT_XLSX,
            sheet_name=SHEET_NAME,
            engine="openpyxl",
            dtype=str,
            keep_default_na=False,
            na_values=[]
        )
    except Exception as e:
        raise SystemExit(f"Erreur lecture Excel : {e}")

    df.columns = [str(c).strip() for c in df.columns]
    for need in (COL_REF, COL_ARTNB):
        if need not in df.columns:
            raise SystemExit(f"Colonne manquante : '{need}'\nColonnes trouvées : {list(df.columns)}")

    # Trie croissant sur ArtNb (numérique si possible)
    df["__num__"] = df[COL_ARTNB].apply(art_number)
    df = df.sort_values("__num__").drop(columns="__num__")

    out_path = Path(OUTPUT_TXT)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    lines_written = 0
    errors = []

    with out_path.open("w", encoding="utf-8", newline="") as fw:
        for _, row in df.iterrows():
            ref_raw = str(row.get(COL_REF, "")).strip()
            art = art_label_from(row.get(COL_ARTNB))
            ref = normalize_ref(ref_raw)

            if not ref or not art:
                errors.append((row.get(COL_ARTNB, ""), row.get(COL_REF, "")))
                continue

            fw.write(f"chr -\t{art}\t{ref}\t0\t{END_VALUE}\t{COLOR_VALUE}\n")
            lines_written += 1

    print(f"✅ Fichier écrit : {out_path}")
    print(f"   Lignes écrites : {lines_written}")
    if errors:
        print(f"⚠️  Lignes ignorées (ArtNb ou ref vide) : {len(errors)}")


if __name__ == "__main__":
    main()
