import pandas as pd

# 1. Load the Excel file
# Replace 'your_file.xlsx' with your actual file path
# We specify the sheet_name='Global Overview' as requested
file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"
try:
    df = pd.read_excel(file_path, sheet_name='Global_overview')
except Exception as e:
    print(f"Error loading file: {e}")
    # Stop execution if file/sheet not found
    exit()


# 2. Function to normalize and check CP type status
def check_status(value):
    """
    Returns:
    - 'PRESENT' if value is a number > 0 or 'X' (case insensitive)
    - 'ABSENT' if value is 0 or '0'
    - 'UNKNOWN' if value is '???', empty, or NaN
    """
    # Convert to string to handle 'X', '???', etc.
    s_val = str(value).strip().upper()

    # Case 1: It is a clear 0
    if s_val == '0' or s_val == '0.0':
        return 'ABSENT'

    # Case 2: It is '???' or empty/NaN
    if s_val == '???' or s_val == 'NAN' or s_val == '':
        return 'UNKNOWN'

    # Case 3: It is an 'X'
    if s_val == 'X':
        return 'PRESENT'

    # Case 4: It is a number (e.g., 12, 5.0)
    # Try converting to float to check if > 0
    try:
        float_val = float(value)
        if float_val > 0:
            return 'PRESENT'
        elif float_val == 0:
            return 'ABSENT'
    except (ValueError, TypeError):
        pass

    # Default: Treat anything else as unknown
    return 'UNKNOWN'


# 3. Apply the logic to the CP columns
# Creating temporary status columns to simplify filtering
target_cols = ['Hemiplegic', 'Diplegic', 'Quadriplegic']

# Check if columns exist in the file
if not all(col in df.columns for col in target_cols):
    print(f"Error: The sheet must contain columns: {target_cols}")
else:
    for col in target_cols:
        df[f'{col}_status'] = df[col].apply(check_status)

    # 4. Create filters for the 4 specific categories

    # Group 1: Hemiplegic AND Diplegic (Without Quadriplegic)
    # Logic: Hemi=PRESENT, Di=PRESENT, Quad=ABSENT
    group1 = df[
        (df['Hemiplegic_status'] == 'PRESENT') &
        (df['Diplegic_status'] == 'PRESENT') &
        (df['Quadriplegic_status'] == 'ABSENT')
        ]

    # Group 2: Diplegic AND Quadriplegic (Without Hemiplegic)
    # Logic: Di=PRESENT, Quad=PRESENT, Hemi=ABSENT
    group2 = df[
        (df['Diplegic_status'] == 'PRESENT') &
        (df['Quadriplegic_status'] == 'PRESENT') &
        (df['Hemiplegic_status'] == 'ABSENT')
        ]

    # Group 3: Hemiplegic AND Quadriplegic (Without Diplegic)
    # Logic: Hemi=PRESENT, Quad=PRESENT, Di=ABSENT
    group3 = df[
        (df['Hemiplegic_status'] == 'PRESENT') &
        (df['Quadriplegic_status'] == 'PRESENT') &
        (df['Diplegic_status'] == 'ABSENT')
        ]

    # Group 4: All Three (Hemiplegic AND Diplegic AND Quadriplegic)
    # Logic: Hemi=PRESENT, Di=PRESENT, Quad=PRESENT
    group4 = df[
        (df['Hemiplegic_status'] == 'PRESENT') &
        (df['Diplegic_status'] == 'PRESENT') &
        (df['Quadriplegic_status'] == 'PRESENT')
        ]

    # 5. Display the results
    cols_to_show = ['ArtNb', 'ref', 'title', 'Hemiplegic', 'Diplegic', 'Quadriplegic']

    print(f"--- 1. Hemiplegic + Diplegic (ONLY) : {len(group1)} articles ---")
    if not group1.empty:
        print(group1[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    print(f"--- 2. Diplegic + Quadriplegic (ONLY) : {len(group2)} articles ---")
    if not group2.empty:
        print(group2[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    print(f"--- 3. Hemiplegic + Quadriplegic (ONLY) : {len(group3)} articles ---")
    if not group3.empty:
        print(group3[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    print(f"--- 4. All Three (Hemi + Di + Quad) : {len(group4)} articles ---")
    if not group4.empty:
        print(group4[cols_to_show].to_string(index=False))
    else:
        print("None found.")
    print("\n")

    # Optional: Save results to a new Excel file
    # output_file = 'CP_Analysis_Results.xlsx'
    # with pd.ExcelWriter(output_file) as writer:
    #     group1[cols_to_show].to_excel(writer, sheet_name='Hemi_Di_Only', index=False)
    #     group2[cols_to_show].to_excel(writer, sheet_name='Di_Quad_Only', index=False)
    #     group3[cols_to_show].to_excel(writer, sheet_name='Hemi_Quad_Only', index=False)
    #     group4[cols_to_show].to_excel(writer, sheet_name='All_Three', index=False)
    # print(f"Results saved to {output_file}")