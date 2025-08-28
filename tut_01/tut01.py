import streamlit as st
import pandas as pd
import os
import shutil

OUTPUT_DIR = "output_files"

def clean_output_dir():
    """Delete old output folder and recreate"""
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

def rows_from_students_list(students):
    """Convert list of pd.Series/dicts into rows for DataFrame"""
    rows = []
    for s in students:
        if isinstance(s, pd.Series):
            rows.append(s.to_dict())
        elif isinstance(s, dict):
            rows.append(s)
    return rows

def process_file(uploaded_file, n):
    clean_output_dir()
    df = pd.read_excel(uploaded_file)

    if 'Roll' not in df.columns:
        st.error("Error: 'Roll' column not found in Excel file.")
        return None

    df['Branch'] = df['Roll'].astype(str).str[4:6]
    branches = list(df['Branch'].unique())

    created = {
        'branch_files': [],
        'round_robin_files': [],
        'uniform_files': [],
        'round_robin_summary': None,
        'uniform_summary': None
    }

    # 1. Branch-wise segregation
    branch_out = os.path.join(OUTPUT_DIR, "branch_wise")
    os.makedirs(branch_out, exist_ok=True)
    for b in branches:
        branch_df = df[df['Branch'] == b].drop(columns=['Branch'])
        path = os.path.join(branch_out, f"{b}.csv")
        branch_df.to_csv(path, index=False)
        created['branch_files'].append(path)

    # 2. Round robin mix
    branch_dict = {b: df[df['Branch'] == b].reset_index(drop=True) for b in branches}
    max_len = max(len(v) for v in branch_dict.values())
    rr_out = os.path.join(OUTPUT_DIR, "branch_wise_mix")
    os.makedirs(rr_out, exist_ok=True)

    rr_files = [[] for _ in range(n)]
    idx = 0
    for i in range(max_len):
        for b in branches:
            if i < len(branch_dict[b]):
                rr_files[idx % n].append(branch_dict[b].iloc[i])
                idx += 1

    rr_counts = []
    for i, students in enumerate(rr_files):
        rows = rows_from_students_list(students)
        df_mix = pd.DataFrame(rows)
        counts = df_mix['Branch'].value_counts().reindex(branches, fill_value=0)
        rr_counts.append(counts)
        df_mix.drop(columns=['Branch']).to_csv(os.path.join(rr_out, f"round_robin_{i+1}.csv"), index=False)
        created['round_robin_files'].append(os.path.join(rr_out, f"round_robin_{i+1}.csv"))

    rr_summary = pd.DataFrame(rr_counts)
    rr_summary.index = [f"round_robin_{i+1}" for i in range(n)]
    rr_summary_path = os.path.join(OUTPUT_DIR, "branch_wise_mix_summary.xlsx")
    rr_summary.to_excel(rr_summary_path)
    created['round_robin_summary'] = rr_summary_path

    # 3. Uniform mix
    u_out = os.path.join(OUTPUT_DIR, "uniform_mix")
    os.makedirs(u_out, exist_ok=True)
    sorted_branches = sorted(branch_dict.items(), key=lambda x: len(x[1]), reverse=True)

    u_files = [[] for _ in range(n)]
    idx = 0
    for b, bdf in sorted_branches:
        for i in range(len(bdf)):
            u_files[idx % n].append(bdf.iloc[i])
            idx += 1

    u_counts = []
    for i, students in enumerate(u_files):
        rows = rows_from_students_list(students)
        u_df = pd.DataFrame(rows)
        counts = u_df['Branch'].value_counts().reindex(branches, fill_value=0)
        u_counts.append(counts)
        u_df.drop(columns=['Branch']).to_csv(os.path.join(u_out, f"uniform_mix_{i+1}.csv"), index=False)
        created['uniform_files'].append(os.path.join(u_out, f"uniform_mix_{i+1}.csv"))

    u_summary = pd.DataFrame(u_counts)
    u_summary.index = [f"uniform_mix_{i+1}" for i in range(n)]
    u_summary_path = os.path.join(OUTPUT_DIR, "uniform_mix_summary.xlsx")
    u_summary.to_excel(u_summary_path)
    created['uniform_summary'] = u_summary_path

    return created

# ---------------- Streamlit UI ----------------
st.title("Student Grouping Tool")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
n = st.number_input("Enter number of groups", min_value=1, step=1, value=5)

if "created" not in st.session_state:
    st.session_state.created = None

if uploaded_file:
    if st.button("Create Groups"):
        created = process_file(uploaded_file, int(n))
        st.session_state.created = created
        st.success("Files generated successfully!")

if st.session_state.created:
    created = st.session_state.created
    option = st.radio("Choose output type", ["All Branch", "Mix Branch (Round Robin)", "Uniform Branch"])

    if option == "All Branch":
        st.write("Branch-wise files:")
        for f in created['branch_files']:
            with open(f, "rb") as file:
                st.download_button(f"Download {os.path.basename(f)}", file.read(), file_name=os.path.basename(f))

    elif option == "Mix Branch (Round Robin)":
        st.write("Round-robin files:")
        for f in created['round_robin_files']:
            with open(f, "rb") as file:
                st.download_button(f"Download {os.path.basename(f)}", file.read(), file_name=os.path.basename(f))
        # ✅ Summary file download
        with open(created['round_robin_summary'], "rb") as file:
            st.download_button("Download Round Robin Summary (Excel)", file.read(), file_name="branch_wise_mix_summary.xlsx")

    elif option == "Uniform Branch":
        st.write("Uniform mix files:")
        for f in created['uniform_files']:
            with open(f, "rb") as file:
                st.download_button(f"Download {os.path.basename(f)}", file.read(), file_name=os.path.basename(f))
        # ✅ Summary file download
        with open(created['uniform_summary'], "rb") as file:
            st.download_button("Download Uniform Mix Summary (Excel)", file.read(), file_name="uniform_mix_summary.xlsx")
