import pandas as pd

# 1. Load the Excel file
# Replace with your actual file path
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
        return False  # Absent
    if s_val == '???' or s_val == 'NAN' or s_val == '':
        return False  # Considered absent for grouping, or handle as unknown
    if s_val == 'X':
        return True  # Present
    try:
        float_val = float(value)
        if float_val > 0:
            return True  # Present
    except (ValueError, TypeError):
        pass
    return False


# 3. Define the columns to check
subtype_cols = ['Spastic', 'Ataxic', 'Dyskinetic', 'Mixed']

# Check if columns exist
if not all(col in df.columns for col in subtype_cols):
    print(f"Error: The sheet must contain columns: {subtype_cols}")
else:
    # 4. Create the 'Combination' column
    def get_combination(row):
        present_types = []
        for col in subtype_cols:
            if check_status(row[col]):
                present_types.append(col)

        if not present_types:
            return "Unspecified / None"
        return " + ".join(present_types)


    df['Subtype_Combination'] = df.apply(get_combination, axis=1)

    # 5. Group by combination and display
    # We count how many articles per combination
    combination_counts = df['Subtype_Combination'].value_counts()

    print("--- SUMMARY OF ALL FOUND COMBINATIONS ---\n")
    print(combination_counts)
    print("\n" + "=" * 50 + "\n")

    # 6. Detail for each combination
    cols_to_show = ['ArtNb', 'ref', 'title'] + subtype_cols

    # Get unique combinations found in the file
    unique_combinations = df['Subtype_Combination'].unique()

    for combo in unique_combinations:
        # Filter data for this specific combination
        subset = df[df['Subtype_Combination'] == combo]

        print(f"### GROUP: {combo} (Count: {len(subset)})")
        print(subset[cols_to_show].to_string(index=False))
        print("\n" + "-" * 50 + "\n")