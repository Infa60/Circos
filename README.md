# Circos Pipeline — CP Review Project

##  Overview
This repository automates the generation of **Circos input files** (`.data.txt`, `.links.txt`, `.numbers.txt`) from an Excel sheet and an optional BibTeX export.  
It visualizes relationships between research articles and experimental variables (GMFCS level, CP type, laterality, assessment tools, etc.) for a systematic review.

### Main components
- **`Extract_name_bibfile.py`** → extracts *short references* and *titles* from a `.bib` file.  
- **`make_articles_data.py`** → creates `articles.data.txt` (Circos “chromosomes” = articles).  
- **`main.py`** → orchestrates the generation of all Circos tracks and manages `"???"` cases.  
