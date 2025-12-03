import pandas as pd
import re
from pathlib import Path

def generate_articles_karyotype(excel_path, sheet_idx, output_dir, col_art, col_ref="ref", end_value=100):
    """
    Generates the articles.data.txt (Karyotype) file from the Excel data.
    
    Args:
        excel_path (str): Path to the Excel file.
        sheet_idx (int): Sheet index.
        output_dir (str): Output directory.
        col_art (str): Name of the Article ID column (e.g., 'ArtNb').
        col_ref (str): Name of the Reference column (e.g., 'ref').
        end_value (int): Visual size for each article segment (default: 100).
    """
    
    # --- Internal Helper Functions ---
    def normalize_ref(ref: str) -> str:
        """Cleans and transforms the reference text."""
        s = str(ref).strip()
        if not s: return s
        # Replace special characters with dashes/underscores for Circos safety
        # (Though labels can contain spaces, cleaner strings avoid parsing errors)
        s = re.sub(r"[^\w\.,&]", "-", s)
        s = re.sub(r"_+", "_", s)
        return s

    def art_label_from(raw) -> str:
        """Standardizes article label to 'artN'."""
        if raw is None: return ""
        s = str(raw).strip()
        if s == "": return ""
        # Matches 'Art 1', 'art1', '1', etc.
        m = re.match(r"^[Aa]rt\s*([0-9]+)$", s) or re.match(r"^([0-9]+)$", s)
        if m: return f"art{m.group(1)}"
        return s if s.lower().startswith("art") else f"art{s}"

    def art_number(raw) -> int:
        """Extracts integer for correct sorting (1, 2, 10 instead of 1, 10, 2)."""
        s = str(raw).strip()
        m = re.search(r"(\d+)", s)
        return int(m.group(1)) if m else 999999

    # --- Main Logic ---
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_idx, engine="openpyxl")
    except Exception as e:
        print(f"[ERROR Articles] Cannot read Excel: {e}")
        return

    # Check columns
    if col_art not in df.columns:
        print(f"[ERROR Articles] Column '{col_art}' not found in Excel.")
        return
    
    if col_ref not in df.columns:
        print(f"[WARN Articles] Column '{col_ref}' not found. Using '{col_art}' as label.")
        col_ref = col_art # Fallback

    # Sorting
    # Create a temporary column to sort numerically
    df["__num__"] = df[col_art].apply(art_number)
    df = df.sort_values("__num__").drop(columns="__num__")

    # Writing
    out_path = Path(output_dir) / "articles.data.txt"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    count = 0
    with out_path.open("w", encoding="utf-8", newline="") as fw:
        for _, row in df.iterrows():
            art_label = art_label_from(row.get(col_art))
            ref_label = normalize_ref(row.get(col_ref, ""))
            
            if not art_label: continue
            
            # Format: chr - art1 Label 0 100 black
            # 'chr' indicates this is a chromosome definition in Circos
            # '-' is the parent (none here)
            # art_label is the ID used for links
            # ref_label is the text displayed
            fw.write(f"chr -\t{art_label}\t{ref_label}\t0\t{end_value}\tblack\n")
            count += 1
            
    print(f"✅ Articles file generated: {out_path} ({count} articles)")