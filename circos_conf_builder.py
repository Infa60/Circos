import os
from pathlib import Path


def generate_circos_conf(output_dir, active_tracks, boundary_map, main_article_file="articles.data.txt"):
    """
    Génère le fichier circos.conf avec des espacements automatiques
    basés sur l'ordre des tracks et leurs contenus réels.
    """

    # 1. Karyotype (Liste des fichiers de données)
    # On inclut seulement les tracks qui ont généré des données (présents dans boundary_map)
    valid_tracks = [t for t in active_tracks if t.subdir in boundary_map]

    track_data_files = [f"{t.subdir}.data.txt" for t in valid_tracks]
    karyotype_string = f"{main_article_file}, {', '.join(track_data_files)}"

    # 2. Plots (Textes / Labels)
    plots_block = ""
    for t in valid_tracks:
        plots_block += f"""
    <plot>
        type           = text
        file           = {t.subdir}.numbers.txt
        r1             = 1200p
        r0             = 710p
        label_font     = bold
        label_size     = 20p
        label_parallel = no
        rpadding       = 0p
        padding        = 0p
    </plot>"""

    # 3. Links (Rubans)
    links_block = ""
    for t in valid_tracks:
        links_block += f"""
    <link>
        file          = {t.subdir}.links.txt
        radius        = dims(ideogram,radius) - 70p
        bezier_radius = 0r
        crest         = 0.3
        thickness     = 1p
        ribbon        = yes
    </link>"""

    # 4. CALCUL AUTOMATIQUE DES ESPACES (SPACING)
    # Logique : Fin Element A -> Debut Element B = 5r

    spacing_block = "default = 0.003r\n"

    # On construit la liste ordonnée des blocs pour le cercle
    ordered_boundaries = []

    # 4.1. Articles (Toujours en premier si présent)
    if 'articles' in boundary_map:
        start, end = boundary_map['articles']
        ordered_boundaries.append({'name': 'articles', 'end': end, 'start': start})

    # 4.2. Tracks (dans l'ordre défini dans le main)
    for t in valid_tracks:
        if t.subdir in boundary_map:
            start, end = boundary_map[t.subdir]
            ordered_boundaries.append({'name': t.subdir, 'end': end, 'start': start})

    # 4.3. Création des règles Pairwise (Boucle circulaire)
    if len(ordered_boundaries) > 1:
        for i in range(len(ordered_boundaries)):
            current = ordered_boundaries[i]
            # Le suivant (avec modulo pour revenir au premier à la fin)
            next_block = ordered_boundaries[(i + 1) % len(ordered_boundaries)]

            # Règle : Fin du courant -> Début du suivant
            spacing_block += f"""
        # Espace entre {current['name']} et {next_block['name']}
        <pairwise {current['end']},{next_block['start']}>
            spacing = 5r
        </pairwise>"""

    # 5. Contenu final du fichier
    conf_content = f"""# ----------------------------------------------
# CONFIGURATION AUTOMATIQUE GENEREE
# ----------------------------------------------
<image> 
    dir* = . 
    radius* = 1500p 
    svg* = yes 
    angle_orientation* = counterclockwise 
    <<include etc/image.conf>> 
</image>

data_out_of_range* = trim

# ----------------------------------------------
# DATA
# ----------------------------------------------
karyotype = {karyotype_string}
chromosomes_units           = 1
chromosomes_display_default = yes
chromosomes_scale           = /art/:0.85

# ----------------------------------------------
# IDEOGRAM
# ----------------------------------------------
<ideogram>
    <spacing>
        {spacing_block}
    </spacing>
    radius    = 0.50r
    thickness = 50p
    fill      = yes
</ideogram>

show_ticks       = no
show_tick_labels = no

# ----------------------------------------------
# LABELS (Noms des sections sur le cercle)
# ----------------------------------------------
<ideogram>
    show_label     = yes
    label_font     = default
    label_radius   = dims(ideogram,radius) + 10p
    label_center   = no
    label_with_tag = no
    label_size     = 17
    label_parallel = no
</ideogram>

# ----------------------------------------------
# PLOTS (Nombres / Totaux)
# ----------------------------------------------
<plots>
{plots_block}
</plots>

# ----------------------------------------------
# LINKS (Rubans de connexion)
# ----------------------------------------------
<links>
{links_block}
</links>

track_defaults* = undef
<<include etc/housekeeping.conf>>
<<include etc/colors_fonts_patterns.conf>>
"""

    out_path = Path(output_dir) / "circos.conf"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(conf_content)

    print(f"✅ Fichier de configuration généré : {out_path}")