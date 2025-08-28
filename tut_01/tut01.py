import pandas as pd
import os

# Configuration
input_file = "students.xlsx"   # Input Excel file
output_dir = "output_files"    # Output directory for all generated files

os.makedirs(output_dir, exist_ok=True)

# --- Read Excel File ---
try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"Error: File '{input_file}' not found!")
    exit()

if 'Roll' not in df.columns:
    print("Error: 'Roll' column not found in Excel file.")
    exit()

# Extract branch
df['Branch'] = df['Roll'].astype(str).str[4:6]

# --- 1. Branch-wise segregation ---
branch_output_dir = os.path.join(output_dir, "branch_wise")
os.makedirs(branch_output_dir, exist_ok=True)
branches = df['Branch'].unique()

for branch in branches:
    branch_df = df[df['Branch'] == branch].drop(columns=['Branch'])
    branch_df.to_csv(os.path.join(branch_output_dir, f"{branch}.csv"), index=False)

print(f"Branch-wise segregation done! Files created in '{branch_output_dir}'")

# --- 2. Round Robin Mix ---
n = int(input("Enter the number of files for round robin mix (n): "))
round_robin_output_dir = os.path.join(output_dir, "branch_wise_mix")
os.makedirs(round_robin_output_dir, exist_ok=True)

branch_dict = {b: df[df['Branch'] == b].reset_index(drop=True) for b in branches}
max_len = max(len(v) for v in branch_dict.values())

round_robin_files = [[] for _ in range(n)]
index = 0
for i in range(max_len):
    for branch in branches:
        if i < len(branch_dict[branch]):
            round_robin_files[index % n].append(branch_dict[branch].iloc[i])
            index += 1

round_robin_counts = []
for idx, students in enumerate(round_robin_files):
    mix_df = pd.DataFrame(students)
    counts = mix_df['Branch'].value_counts().reindex(branches, fill_value=0)
    round_robin_counts.append(counts)
    mix_df.drop(columns=['Branch']).to_csv(os.path.join(round_robin_output_dir, f"round_robin_{idx+1}.csv"), index=False)

round_robin_summary = pd.DataFrame(round_robin_counts)
round_robin_summary.index = [f"round_robin_{i+1}" for i in range(n)]
round_robin_summary.to_excel(os.path.join(output_dir, "branch_wise_mix_summary.xlsx"))

print(f"Round robin mix done! Files and summary created.")

# --- 3. Uniform Mix ---
uniform_mix_output_dir = os.path.join(output_dir, "uniform_mix")
os.makedirs(uniform_mix_output_dir, exist_ok=True)

sorted_branches = sorted(branch_dict.items(), key=lambda x: len(x[1]), reverse=True)

uniform_files = [[] for _ in range(n)]
index = 0
for branch, students_df in sorted_branches:
    for i in range(len(students_df)):
        uniform_files[index % n].append(students_df.iloc[i])
        index += 1

uniform_counts = []
for idx, students in enumerate(uniform_files):
    uniform_df = pd.DataFrame(students)
    counts = uniform_df['Branch'].value_counts().reindex(branches, fill_value=0)
    uniform_counts.append(counts)
    uniform_df.drop(columns=['Branch']).to_csv(os.path.join(uniform_mix_output_dir, f"uniform_mix_{idx+1}.csv"), index=False)

uniform_summary = pd.DataFrame(uniform_counts)
uniform_summary.index = [f"uniform_mix_{i+1}" for i in range(n)]
uniform_summary.to_excel(os.path.join(output_dir, "uniform_mix_summary.xlsx"))

print(f"Uniform mix done! Files and summary created.")
