"""
pipeline_refs.py — Extrait les bibliographies de tous les PDF d'un dossier.
- Ignore automatiquement les faux PDFs AppleDouble (._*.pdf)
- Vérifie la signature %PDF- et la taille minimale
- Lance/Utilise GROBID via Docker Desktop (port 8070)
- Gère les erreurs 500 (retry) et journalise les échecs
- Exporte refs_by_source.csv et refs_unique.csv (avec enrichissement DOI optionnel)

Exécuter : Run ▶ dans PyCharm (ou `python pipeline_refs.py`)
Prérequis : Docker Desktop lancé (🐳 running)
"""

import sys, os, time, re, csv, socket, subprocess
from pathlib import Path

# ========== CONFIG UTILISATEUR ==========
PDF_DIR = Path(r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text")
OUT_DIR = PDF_DIR / "output"
GROBID_PORT = 8070
GROBID_IMAGE = "lfoppiano/grobid:0.8.0"

START_CONTAINER = True          # True: tente de démarrer le conteneur GROBID si absent
ENRICH_WITH_CROSSREF = True     # False pour ne pas interroger Crossref (plus rapide, offline)
MIN_BYTES = 5 * 1024            # taille minimale d'un PDF "utile" (5 Ko)
# =======================================

# ---- auto-install paquets manquants (PyCharm friendly) ----
def ensure_pkg(pkg_name):
    mod = pkg_name.split("==")[0]
    try:
        __import__(mod)
    except ImportError:
        print(f"→ Installation du paquet manquant : {pkg_name}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name])

for dep in ["requests", "pandas", "rapidfuzz", "docker"]:
    ensure_pkg(dep)

import requests
import pandas as pd
from rapidfuzz import process, fuzz
import docker
from xml.etree import ElementTree as ET


# ---------- utilitaires généraux ----------
def port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0

def wait_http_ready(url: str, timeout: int = 120) -> bool:
    start = time.time()
    while time.time() - start < timeout:
        try:
            r = requests.get(url, timeout=3)
            # 404/405 possibles, suffisent à indiquer que le service répond
            if r.status_code in (200, 404, 405):
                return True
        except Exception:
            pass
        time.sleep(1)
    return False

def check_docker_desktop_running() -> None:
    try:
        docker.from_env().ping()
    except Exception as e:
        raise RuntimeError(
            "Docker Desktop ne semble pas actif ou accessible.\n"
            "→ Lance Docker Desktop (icône 🐳 doit afficher 'Running').\n"
            f"Détail : {e}"
        )


# ---------- nettoyage/filtrage AppleDouble & PDFs ----------
def purge_apple_double(pdf_dir: Path) -> int:
    """
    Supprime tous les faux PDFs AppleDouble (._*.pdf) dans le dossier.
    Retourne le nombre de fichiers supprimés.
    """
    removed = 0
    for p in pdf_dir.glob("._*.pdf"):
        try:
            p.unlink()
            removed += 1
        except Exception:
            pass
    return removed

def looks_like_pdf(path: Path) -> bool:
    """
    Ignore ._* et vérifie que le fichier existe, qu'il est assez grand,
    et qu'il commence par la signature %PDF-.
    """
    if path.name.startswith("._"):
        return False
    try:
        if not path.exists():
            return False
        if path.stat().st_size < MIN_BYTES:
            return False
        with path.open("rb") as f:
            return f.read(5) == b"%PDF-"
    except Exception:
        return False


# ---------- 1) Gérer GROBID via Docker ----------
def ensure_grobid_running(port: int, image: str, start_container: bool):
    base_url = f"http://localhost:{port}"
    if port_in_use(port):
        print(f"✔ GROBID détecté sur {base_url}")
        return None  # on n'a pas lancé le conteneur

    if not start_container:
        if not wait_http_ready(base_url, timeout=10):
            raise RuntimeError(
                f"GROBID n'est pas joignable sur {base_url} et START_CONTAINER=False.\n"
                f">> Lance GROBID manuellement ou mets START_CONTAINER=True."
            )
        print(f"✔ GROBID prêt sur {base_url}")
        return None

    print("→ Vérification Docker Desktop…")
    check_docker_desktop_running()

    print("→ Lancement du conteneur GROBID…")
    client = docker.from_env()
    try:
        client.images.get(image)
    except docker.errors.ImageNotFound:
        print(f"  Téléchargement de l'image {image} (une seule fois)…")
        client.images.pull(image)

    container = client.containers.run(
        image,
        detach=True,
        tty=True,
        remove=True,
        ports={f"{port}/tcp": port},
        name=f"grobid-{port}",
    )
    ok = wait_http_ready(base_url)
    if not ok:
        raise RuntimeError("GROBID n'a pas répondu à temps après démarrage du conteneur.")
    print(f"✔ GROBID prêt sur {base_url}")
    return container


# ---------- 2) Appels GROBID + parsing TEI ----------
NS = {"tei": "http://www.tei-c.org/ns/1.0"}
GROBID_URL = f"http://localhost:{GROBID_PORT}/api/processFulltextDocument"

def norm_txt(t: str) -> str:
    if not t:
        return ""
    t = re.sub(r"\s+", " ", t.strip())
    t = re.sub(r"[^\w\s]", "", t)
    return t.lower()

def call_grobid(pdf_path: Path, retries: int = 2, backoff: float = 2.0) -> str:
    """
    Envoie le PDF à GROBID avec quelques retries en cas d'erreurs 500/502/503.
    """
    last_err = None
    for attempt in range(retries + 1):
        try:
            with pdf_path.open("rb") as f:
                files = {"input": (pdf_path.name, f, "application/pdf")}
                data = {
                    "consolidateHeader": "1",
                    "consolidateCitations": "1",
                }
                r = requests.post(GROBID_URL, files=files, data=data, timeout=240)
                r.raise_for_status()
                return r.text
        except requests.exceptions.HTTPError as e:
            last_err = e
            code = getattr(e.response, "status_code", None)
            if code in (500, 502, 503) and attempt < retries:
                time.sleep(backoff * (attempt + 1))
                continue
            raise
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            last_err = e
            if attempt < retries:
                time.sleep(backoff * (attempt + 1))
                continue
            raise
    raise last_err or RuntimeError("Échec GROBID inconnu")

def parse_refs_from_tei(tei_text: str):
    root = ET.fromstring(tei_text)
    out = []
    for b in root.findall(".//tei:listBibl/tei:biblStruct", NS):
        # Titre
        t = b.find(".//tei:title[@level='a']", NS) or b.find(".//tei:title", NS)
        title = t.text.strip() if t is not None and t.text else None
        # Année
        d = b.find(".//tei:date", NS)
        year = None
        if d is not None:
            w = d.attrib.get("when")
            if w:
                year = w[:4]
            elif d.text:
                m = re.search(r"(\d{4})", d.text)
                year = m.group(1) if m else None
        # DOI
        doi_el = b.find('.//tei:idno[@type="DOI"]', NS)
        doi = doi_el.text.strip() if doi_el is not None and doi_el.text else None
        # Premier auteur (simple)
        fa = None
        pers = b.find(".//tei:author//tei:persName", NS)
        if pers is not None:
            fn = pers.find("./tei:forename", NS)
            sn = pers.find("./tei:surname", NS)
            fa = " ".join([x.text.strip() for x in [fn, sn] if x is not None and x.text])

        out.append({
            "title": title,
            "title_norm": norm_txt(title),
            "year": year,
            "doi": doi,
            "first_author": fa
        })
    return out


# ---------- 3) Enrichissement Crossref (optionnel) ----------
def crossref_enrich(title: str):
    if not title:
        return None
    try:
        r = requests.get(
            "https://api.crossref.org/works",
            params={"query.bibliographic": title, "rows": 1},
            timeout=20
        )
        j = r.json()
        items = j.get("message", {}).get("items", [])
        if not items:
            return None
        it = items[0]
        doi = it.get("DOI")
        year = None
        dp = it.get("issued", {}).get("date-parts", [])
        if dp and dp[0]:
            year = dp[0][0]
        return {"doi": doi, "year": str(year) if year else None}
    except Exception:
        return None


# ---------- 4) Pipeline principal ----------
def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # 0) Purge proactive des AppleDouble ._*.pdf
    removed = purge_apple_double(PDF_DIR)
    print(f"🧹 AppleDouble supprimés : {removed}")

    # 1) GROBID prêt ?
    container = ensure_grobid_running(GROBID_PORT, GROBID_IMAGE, START_CONTAINER)

    # 2) Lister & filtrer les PDF
    all_candidates = sorted(PDF_DIR.glob("*.pdf"))
    pdfs = [p for p in all_candidates if looks_like_pdf(p)]

    print(f"📄 PDFs trouvés (brut) : {len(all_candidates)}")
    print(f"✅ PDFs valides (après filtre) : {len(pdfs)}")

    if not pdfs:
        print("⚠️ Aucun PDF valide après filtrage.")
        return

    failures = []
    all_rows = []

    # 3) Traitement
    for i, pdf in enumerate(pdfs, 1):
        print(f"[{i}/{len(pdfs)}] {pdf.name}")
        try:
            if not pdf.exists():
                raise FileNotFoundError("Disparu avant ouverture (OneDrive ?)")
            tei = call_grobid(pdf)
            (OUT_DIR / (pdf.stem + ".tei.xml")).write_text(tei, encoding="utf-8")

            refs = parse_refs_from_tei(tei)
            for r in refs:
                r["source_pdf"] = pdf.name
            all_rows.extend(refs)

        except Exception as e:
            print(f"   ⚠️ Échec sur {pdf.name}: {e}")
            failures.append({"source_pdf": pdf.name, "error": str(e)})
        time.sleep(0.05)

    if not all_rows:
        print("⚠️ Aucune référence extraite depuis les TEI.")
        # journal des échecs si existant
        if failures:
            pd.DataFrame(failures).to_csv(OUT_DIR / "grobid_failures.csv",
                                          index=False, encoding="utf-8-sig")
            print(f"📝 Journal des échecs : {OUT_DIR/'grobid_failures.csv'} ({len(failures)} fichiers)")
        return

    # 4) CSV complet (par source)
    cols = ["source_pdf", "title", "year", "doi", "first_author", "title_norm"]
    df_all = pd.DataFrame(all_rows, columns=cols)
    df_all.to_csv(OUT_DIR / "refs_by_source.csv", index=False, encoding="utf-8-sig")
    print(f"✅ Écrit : {OUT_DIR / 'refs_by_source.csv'} ({len(df_all)} lignes)")

    # 5) Dédoublonnage (DOI puis titre normalisé)
    seen_doi, seen_title = set(), set()
    uniq = []
    for r in all_rows:
        if r["doi"]:
            key = r["doi"].lower().strip()
            if key in seen_doi:
                continue
            seen_doi.add(key)
            uniq.append(r)
            continue
        t = r["title_norm"]
        if t and t in seen_title:
            continue
        if t:
            seen_title.add(t)
        uniq.append(r)

    # 6) Enrichissement DOI via Crossref (optionnel)
    if ENRICH_WITH_CROSSREF:
        added = 0
        for r in uniq:
            if r.get("doi"):
                continue
            hit = crossref_enrich(r.get("title"))
            if hit and hit.get("doi"):
                r["doi"] = hit["doi"]
                if not r.get("year") and hit.get("year"):
                    r["year"] = hit["year"]
                added += 1
                time.sleep(0.12)
        print(f"🔎 DOIs ajoutés via Crossref : {added}")

    df_uniq = pd.DataFrame(uniq, columns=cols)
    df_uniq.to_csv(OUT_DIR / "refs_unique.csv", index=False, encoding="utf-8-sig")
    print(f"✅ Écrit : {OUT_DIR / 'refs_unique.csv'} ({len(df_uniq)} lignes)")

    # 7) Journal des échecs
    if failures:
        pd.DataFrame(failures).to_csv(OUT_DIR / "grobid_failures.csv",
                                      index=False, encoding="utf-8-sig")
        print(f"📝 Journal des échecs : {OUT_DIR/'grobid_failures.csv'} ({len(failures)} fichiers)")

    print("\n🎉 Terminé. Dossier de sortie :", OUT_DIR)


if __name__ == "__main__":
    main()
