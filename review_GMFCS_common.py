import pandas as pd

# 1. Load the Excel file
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"

try:
    df = pd.read_excel(file_path, sheet_name='Global_overview')
except Exception as e:
    print(f"Error loading file: {e}")
    exit()


# 2. Function to normalize and check status
def check_status(value):
    s_val = str(value).strip().upper()

    if s_val == '0' or s_val == '0.0':
        return 'ABSENT'
    if s_val == '???' or s_val == 'NAN' or s_val == '':
        return 'UNKNOWN'
    if s_val == 'X':
        return 'PRESENT'
    try:
        float_val = float(value)
        if float_val > 0:
            return 'PRESENT'
        elif float_val == 0:
            return 'ABSENT'
    except (ValueError, TypeError):
        pass
    return 'UNKNOWN'


# 3. Apply the logic to the GMFCS columns
target_cols = ['GMFCS-I', 'GMFCS-II', 'GMFCS-III', 'GMFCS-IV']

if not all(col in df.columns for col in target_cols):
    print(f"Error: The sheet must contain columns: {target_cols}")
else:
    for col in target_cols:
        df[f'{col}_status'] = df[col].apply(check_status)

    # --- FILTERS ---

    # Group 1: GMFCS I & II ONLY
    group_I_II = df[
        (df['GMFCS-I_status'] == 'PRESENT') &
        (df['GMFCS-II_status'] == 'PRESENT') &
        (df['GMFCS-III_status'] == 'ABSENT') &
        (df['GMFCS-IV_status'] == 'ABSENT')
        ]

    # Group 2: GMFCS I & II & III ONLY
    group_I_II_III = df[
        (df['GMFCS-I_status'] == 'PRESENT') &
        (df['GMFCS-II_status'] == 'PRESENT') &
        (df['GMFCS-III_status'] == 'PRESENT') &
        (df['GMFCS-IV_status'] == 'ABSENT')
        ]

    # Group 3: GMFCS I & II & III & IV (All)
    group_All = df[
        (df['GMFCS-I_status'] == 'PRESENT') &
        (df['GMFCS-II_status'] == 'PRESENT') &
        (df['GMFCS-III_status'] == 'PRESENT') &
        (df['GMFCS-IV_status'] == 'PRESENT')
        ]

    # Group 4: GMFCS III & IV ONLY
    group_III_IV = df[
        (df['GMFCS-III_status'] == 'PRESENT') &
        (df['GMFCS-IV_status'] == 'PRESENT') &
        (df['GMFCS-I_status'] == 'ABSENT') &
        (df['GMFCS-II_status'] == 'ABSENT')
        ]

    # Group 5: GMFCS II & III & IV ONLY (Added per request)
    # Logic: II=PRESENT, III=PRESENT, IV=PRESENT, I=ABSENT
    group_II_III_IV = df[
        (df['GMFCS-II_status'] == 'PRESENT') &
        (df['GMFCS-III_status'] == 'PRESENT') &
        (df['GMFCS-IV_status'] == 'PRESENT') &
        (df['GMFCS-I_status'] == 'ABSENT')
        ]

    # --- DISPLAY RESULTS ---
    cols_to_show = ['ArtNb', 'ref', 'title'] + target_cols

    print(f"--- 1. GMFCS I + II ONLY: {len(group_I_II)} articles ---")
    if not group_I_II.empty:
        print(group_I_II[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    print(f"--- 2. GMFCS I + II + III ONLY: {len(group_I_II_III)} articles ---")
    if not group_I_II_III.empty:
        print(group_I_II_III[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    print(f"--- 3. GMFCS I + II + III + IV (All): {len(group_All)} articles ---")
    if not group_All.empty:
        print(group_All[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    print(f"--- 4. GMFCS III + IV ONLY: {len(group_III_IV)} articles ---")
    if not group_III_IV.empty:
        print(group_III_IV[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    print(f"--- 5. GMFCS II + III + IV ONLY: {len(group_II_III_IV)} articles ---")
    if not group_II_III_IV.empty:
        print(group_II_III_IV[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")