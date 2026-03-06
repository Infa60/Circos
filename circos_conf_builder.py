"""
================================================================================
CIRCOS CONFIGURATION BUILDER
================================================================================

Description:
This script dynamically generates the master `circos.conf` file required by the
Circos Perl engine to render the final circular visualization. It eliminates the
need for manual tweaking of the configuration file.

Key Features:
1. Dynamic Karyotype & Tracks: Only includes tracks that successfully generated
   data, preventing Circos from crashing due to missing files.
2. Automated Spacing (The Pairwise Logic): This is the most crucial part of
   the script. Circos plots can look messy if spacing isn't handled correctly.
   This script calculates exactly where one category (e.g., GMFCS levels) ends
   and where the next (e.g., Topography) begins. It then injects `<pairwise>`
   rules to enforce a distinct, consistent visual gap (5r) between different
   groups, while keeping elements within the same group tightly packed.
3. Block Generation: Automatically constructs the `<plot>` (for text/numbers)
   and `<link>` (for the internal connecting ribbons) blocks based on the active
   tracks defined in the orchestrator.

Output:
A ready-to-use `circos.conf` file saved in the specified output directory.
================================================================================
"""

import os
from pathlib import Path


def generate_circos_conf(output_dir, active_tracks, boundary_map, main_article_file="articles.data.txt"):
    """
    Generates the circos.conf file with automatic spacing based on
    track order and their actual contents.
    """

    # 1. Karyotype (List of data files)
    # We only include tracks that generated data (present in boundary_map)
    valid_tracks = [t for t in active_tracks if t.subdir in boundary_map]

    track_data_files = [f"{t.subdir}.data.txt" for t in valid_tracks]
    karyotype_string = f"{main_article_file}, {', '.join(track_data_files)}"

    # 2. Plots (Texts / Labels)
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

    # 3. Links (Ribbons)
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

    # 4. AUTOMATIC SPACING CALCULATION
    # Logic: End of Element A -> Start of Element B = 5r

    spacing_block = "default = 0.003r\n"

    # Build an ordered list of blocks for the circle
    ordered_boundaries = []

    # 4.1. Articles (Always first if present)
    if 'articles' in boundary_map:
        start, end = boundary_map['articles']
        ordered_boundaries.append({'name': 'articles', 'end': end, 'start': start})

    # 4.2. Tracks (in the order defined in main)
    for t in valid_tracks:
        if t.subdir in boundary_map:
            start, end = boundary_map[t.subdir]
            ordered_boundaries.append({'name': t.subdir, 'end': end, 'start': start})

    # 4.3. Creation of Pairwise rules (Circular loop)
    if len(ordered_boundaries) > 1:
        for i in range(len(ordered_boundaries)):
            current = ordered_boundaries[i]
            # The next one (with modulo to loop back to the first at the end)
            next_block = ordered_boundaries[(i + 1) % len(ordered_boundaries)]

            # Rule: End of current -> Start of next
            spacing_block += f"""
        # Space between {current['name']} and {next_block['name']}
        <pairwise {current['end']},{next_block['start']}>
            spacing = 5r
        </pairwise>"""

    # 5. Final file content
    conf_content = f"""# ----------------------------------------------
# AUTOMATICALLY GENERATED CONFIGURATION
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
# LABELS (Section names on the circle)
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
# PLOTS (Numbers / Totals)
# ----------------------------------------------
<plots>
{plots_block}
</plots>

# ----------------------------------------------
# LINKS (Connection ribbons)
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

    print(f"✅ Configuration file generated: {out_path}")