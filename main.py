import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

print("=== Importation et affichage de la feuille Excel ===\n")

df1 = pd.read_excel('student_grades_2027-2028.xlsx', sheet_name='Data')
df2 = pd.read_excel('student_grades_2027-2028.xlsx', sheet_name='Finance')
df3 = pd.read_excel('student_grades_2027-2028.xlsx', sheet_name='BM')

df1["Track"] = "Data"
df2["Track"] = "Finance"
df3["Track"] = "BM"

unified_df = pd.concat([df1, df2, df3], ignore_index=True)

unified_df.columns = unified_df.columns.str.strip()
subjects = ["Math", "English", "Science", "History"]

for col in subjects:
    unified_df[col] = unified_df[col].replace("WAIVE", np.nan)

unified_df["Attendance (%)"] = unified_df["Attendance (%)"].replace("", np.nan)

unified_df["Passed (Y/N)"] = unified_df["Passed (Y/N)"].replace("no", "N")
unified_df["Passed (Y/N)"] = unified_df["Passed (Y/N)"].str.upper()

# Ensure numeric types for numeric columns
for col in subjects + ["Attendance (%)"]:
    unified_df[col] = pd.to_numeric(unified_df[col], errors="coerce")

if "ProjectScore" in unified_df.columns:
    unified_df["ProjectScore"] = pd.to_numeric(unified_df["ProjectScore"],
                                               errors="coerce")

unified_df.to_csv("cleaned_file.csv", index=False, encoding="utf-8")


# Average scores function
def moyenne(colonne):
    colonne_numerique = pd.to_numeric(colonne, errors='coerce')
    return colonne_numerique.mean()


df_math = moyenne(unified_df["Math"])
df_english = moyenne(unified_df["English"])
df_science = moyenne(unified_df["Science"])
df_history = moyenne(unified_df["History"])
df_attendance = moyenne(unified_df["Attendance (%)"])

print("\n=== Track-Level Summary Statistics ===")
print("\nAverage subject scores")
print(f"Math: {df_math:.2f}")
print(f"English: {df_english:.2f}")
print(f"Science: {df_science:.2f}")
print(f"History: {df_history:.2f}")
print(f"Attendance (%): {df_attendance:.2f}")

# Track counts
track_counts = unified_df["Track"].value_counts()
data_students = track_counts.get("Data", 0)
finance_students = track_counts.get("Finance", 0)
bm_students = track_counts.get("BM", 0)

print(f"\nStudents in Data: {data_students}")
print(f"Students in Finance: {finance_students}")
print(f"Students in Business Management: {bm_students}")

# Pass rate
pass_counts = unified_df["Passed (Y/N)"].value_counts()
pass_rate = pass_counts.get("Y", 0) / pass_counts.sum() * 100
print(f"\nPass rate: {pass_rate:.2f}%")

# Cross-Track Comparative Analysis
print("\n=== Cross-Track Comparative Analysis ===")
comparison_table = pd.DataFrame({
    "Average Math Score":
    unified_df.groupby("Track")["Math"].mean().round(2)
})

print(comparison_table)

# Boxplots: History grade distribution by track
plt.figure(figsize=(8, 6))

unified_df.boxplot(column="History", by="Track", grid=False, patch_artist=True)

plt.title("Boxplot des notes d'Histoire par track")
plt.suptitle("")

plt.xlabel("Track")
plt.ylabel("Note d'Histoire")
plt.tight_layout()
plt.savefig("boxplot_history.png")
plt.close()

# Bar chart for average Math per track
plt.figure(figsize=(7, 5))
comparison_table["Average Math Score"].plot(kind="bar", edgecolor="black")
plt.title("Average Math Score per Track")
plt.xlabel("Track")
plt.ylabel("Average Math Score")
plt.grid(axis="y", alpha=0.3)
plt.tight_layout()
plt.savefig("avg_math_per_track.png")
plt.close()

# Correlation between Attendance and Project Score
print('\nCorrelation between Attendance and Project Score')

if "ProjectScore" in unified_df.columns:
    if unified_df["ProjectScore"].notna().sum(
    ) > 0 and unified_df["Attendance (%)"].notna().sum() > 0:
        correlation = unified_df["Attendance (%)"].corr(
            unified_df["ProjectScore"])
        print(f"Correlation: {correlation:.4f}")
    else:
        print("Not enough numeric data to compute correlation.")
else:
    print("Column 'ProjectScore' not found in the dataset.")

# Cohort-Level Analysis
print("\n=== Cohort-Level Analysis ===")

avg_by_cohort = unified_df.groupby("Cohort")[[
    "Math", "English", "Science", "History"
]].mean().round(2)
print(avg_by_cohort)

print("\nPass Rate by Cohort:")
pass_rate_by_cohort = (unified_df.groupby("Cohort")["Passed (Y/N)"].apply(
    lambda x: (x == "Y").mean() * 100).round(2))
print(pass_rate_by_cohort)

avg_by_income = unified_df.groupby("IncomeStudent")[[
    "Math", "English", "Science", "History"
]].mean()

unified_df["IncomeStudent"] = unified_df["IncomeStudent"].replace({
    0: "No",
    1: "Yes"
})
pass_rate_income = (unified_df.groupby("IncomeStudent")["Passed (Y/N)"].apply(lambda x: (x == "Y").mean() * 100).round(2).reset_index(name="Pass Rate (%)"))

print("\nPass Rate by Income Status:")
print(pass_rate_income.to_string(index=False))

#Final Export
summary_data = {
    "Average Math": [df_math],
    "Average English": [df_english],
    "Average Science": [df_science],
    "Average History": [df_history],
    "Average Attendance (%)": [df_attendance],
    "Total Students Data": [data_students],
    "Total Students Finance": [finance_students],
    "Total Students BM": [bm_students],
    "Overall Pass Rate (%)": [pass_rate],
}

summary_df = pd.DataFrame(summary_data)
summary_df.to_csv("summary_statistics.csv", index=False)

print("\nSummary report exported as summary_statistics.csv")

plt.figure(figsize=(8, 6))
unified_df.boxplot(column="History", by="Track", grid=False, patch_artist=True)
plt.title("Boxplot des notes d'Histoire par track")
plt.suptitle("")
plt.xlabel("Track")
plt.ylabel("Note d'Histoire")
plt.tight_layout()
plt.savefig("boxplot_history.png")
plt.close()

plt.figure(figsize=(7, 5))
comparison_table["Average Math Score"].plot(kind="bar", edgecolor="black")
plt.title("Average Math Score per Track")
plt.xlabel("Track")
plt.ylabel("Average Math Score")
plt.grid(axis="y", alpha=0.3)
plt.tight_layout()
plt.savefig("avg_math_per_track.png")
plt.close()
