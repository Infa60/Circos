import pandas as pd
import numpy as np
from sklearn.metrics import cohen_kappa_score

# --- CONFIGURATION ---
# Replace with the actual path to your Excel file
excel_file_path = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Quality_assessment_kappa.xlsx"

# Exact names of your Excel sheets
sheet_rater1 = 'QA_MB_v2'
sheet_rater2 = 'QA_NH_v2'

# Exact list of columns (must match Excel headers perfectly)
columns_of_interest = [
    "1 Aims and hypothesis clearly stated",
    "2 Ethics and consent",
    "3 Description of the participant recruitment",
    "4 Description of the sample",
    "5 Sample size calculation",
    "6 Instrumented measure description",
    "7 Description of the movement tasks",
    "8 Data analysis",
    "9 Main outcomes of the study",
    "10 Statistical analysis",
    "11 Interpretable results",
    "12 Description of study limits",
    "13 Key findings answer the initial objectives"
]


def analyze_reliability():
    try:
        print("Loading data...")
        df_mb = pd.read_excel(excel_file_path, sheet_name=sheet_rater1)
        df_nh = pd.read_excel(excel_file_path, sheet_name=sheet_rater2)

        # Check for missing columns
        missing = [c for c in columns_of_interest if c not in df_mb.columns]
        if missing:
            print(f"ERROR: The following columns are missing in the Excel file: {missing}")
            return

        # Storage for global pooled calculation
        global_mb = []
        global_nh = []

        # Print Table Header
        print("\n" + "=" * 105)
        print(f"{'ITEM':<55} | {'WEIGHTED KAPPA'} | {'AGREEMENT (%)'} | {'INTERPRETATION'}")
        print("=" * 105)

        # --- LOOP THROUGH EACH ITEM ---
        for col in columns_of_interest:
            # Clean data: drop empty cells (NaNs)
            s1 = df_mb[col].dropna()
            s2 = df_nh[col].dropna()

            # Align indices to ensure we compare the same rows
            common_idx = s1.index.intersection(s2.index)

            if len(common_idx) > 0:
                val1 = s1.loc[common_idx].values
                val2 = s2.loc[common_idx].values

                # 1. Calculate Linear Weighted Kappa
                # Handle cases where Kappa is impossible (e.g., only one score used)
                try:
                    kappa = cohen_kappa_score(val1, val2, weights='linear')
                except:
                    kappa = 0.0

                # 2. Calculate Percent Agreement
                agreement = np.mean(val1 == val2) * 100

                # Add to global lists for pooled calculation later
                global_mb.extend(val1)
                global_nh.extend(val2)

                # Logic for text interpretation
                if kappa < 0.40:
                    interp = "Weak"
                elif kappa < 0.60:
                    interp = "Moderate"
                elif kappa < 0.80:
                    interp = "Good"
                else:
                    interp = "Excellent"

                # Special Case: "Kappa Paradox" (Low Kappa but High Agreement)
                if kappa < 0.40 and agreement > 85:
                    interp = "Paradox*"

                # Print row
                print(f"{col[:53]:<55} | {kappa:.3f}          | {agreement:.1f}%         | {interp}")
            else:
                print(f"{col[:53]:<55} | ---            | ---           | No data")

        print("-" * 105)

        # --- GLOBAL CALCULATIONS (POOLED) ---
        if len(global_mb) > 0:
            # Convert lists to numpy arrays
            pool_mb = np.array(global_mb)
            pool_nh = np.array(global_nh)

            # Global Weighted Kappa
            k_global = cohen_kappa_score(pool_mb, pool_nh, weights='linear')

            # Global Percent Agreement
            a_global = np.mean(pool_mb == pool_nh) * 100

            print(
                f"{'GLOBAL RESULT (Pooled across all items)':<55} | {k_global:.3f}          | {a_global:.1f}%         | GLOBAL")
            print("=" * 105)
            print(
                f"*Paradox: Kappa is low due to a lack of variance (ceiling effect), but the actual agreement is high.")
            print(f"Calculation base: {len(pool_mb)} total observations.")

    except FileNotFoundError:
        print(f"Error: Could not find the file '{excel_file_path}'.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


analyze_reliability()
