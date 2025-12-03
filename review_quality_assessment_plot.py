import pandas as pd
import matplotlib.pyplot as plt

EXCEL_PATH = r"C:\Users\bourgema\OneDrive - Université de Genève\Documents\ENABLE\Review\Full_text_inclusion_v1.xlsx"

# 1. Define exact column names
noms_colonnes = [
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

# --- LOAD DATA ---
df = pd.read_excel(EXCEL_PATH, sheet_name='QA_NH_v2')

# ---------------------------------------------

# 2. Count occurrences of 0, 1, and 2 per column
counts = df[noms_colonnes].apply(pd.Series.value_counts).fillna(0)

# 3. Ensure index order (0, 1, 2) for consistent coloring
counts = counts.reindex([0, 1, 2])

# 4. Create Plot
plt.figure(figsize=(7, 4))

# Create stacked bar chart
ax = counts.T.plot(kind='bar',
                   stacked=True,
                   color=['#d9534f', '#f0ad4e', '#5cb85c'], # Red, Orange, Green
                   edgecolor='black',
                   width=0.8,
                   figsize=(7, 4))

# 5. Formatting (English Labels)
plt.title("Methodological Quality per Criteria (Score Distribution)", fontsize=12)
plt.ylabel("Frequency / Number of Studies", fontsize=10)
plt.xlabel("Evaluation Criteria", fontsize=10)

# Rotate x-axis labels
plt.xticks(rotation=45, ha='right', fontsize=8)

# Custom legend configuration in English
plt.legend(title="Score",
           labels=['0 (No)', '1 (Partial)', '2 (Yes)'],
           loc='upper center',
           bbox_to_anchor=(0.5, -0.35),
           ncol=3)

plt.subplots_adjust(bottom=0.5, top=0.85, left=0.2)
# 6. Display plot
plt.show()