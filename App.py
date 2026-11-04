import streamlit as st
import re
import xlsxwriter
import io
from decimal import Decimal
import pandas as pd

HARTREE_TO_KCAL = Decimal('627.509')

def extract_energy(uploaded_file):
    if uploaded_file is None: return None
    content = uploaded_file.read().decode('utf-8')
    stitched = re.sub(r'[\r\n]+ ', '', content)
    matches = re.findall(r'HF=([-+]?\d*\.\d+|\d+)', stitched)
    return Decimal(matches[-1]) if matches else None

st.title("Interaction Energy Calculator")

# --- UI: FILE UPLOADS ---
col1, col2 = st.columns(2)
with col1:
    m1_file = st.file_uploader("Upload Monomer 1 (Amino Acid)", type=["log", "out"])
with col2:
    m2_file = st.file_uploader("Upload Monomer 2 (Solvent)", type=["log", "out"])

st.subheader("Upload Dimer / Complex Files")
# accept_multiple_files=True replaces the need for a manual "Add more" button
dimer_files = st.file_uploader("Upload all Dimer files", type=["log", "out"], accept_multiple_files=True)

if st.button("Calculate Energies"):
    if not (m1_file and m2_file and dimer_files):
        st.error("Please upload Monomer 1, Monomer 2, and at least one Dimer.")
    else:
        e_m1 = extract_energy(m1_file)
        e_m2 = extract_energy(m2_file)
        
        results = []
        for d_file in dimer_files:
            e_dimer = extract_energy(d_file)
            if e_dimer and e_m1 and e_m2:
                delta_au = e_dimer - (e_m1 + e_m2)
                delta_kcal = delta_au * HARTREE_TO_KCAL
                results.append({
                    "Complex": d_file.name,
                    "E_Complex (a.u.)": float(e_dimer),
                    "ΔE (a.u.)": float(delta_au),
                    "ΔE (kcal/mol)": float(delta_kcal)
                })
        
        # --- UI: TABS FOR TABULATION & DOWNLOAD ---
        tab1, tab2 = st.tabs(["Tabulation", "Download Excel"])
        
        with tab1:
            df = pd.DataFrame(results)
            st.dataframe(df, use_container_width=True)
            
        with tab2:
            # Write Excel to memory buffer instead of hard drive
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet("Energies")
            
            # Write headers and data (simplified for example)
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name)
            for row_num, row in enumerate(results, start=1):
                worksheet.write(row_num, 0, row["Complex"])
                worksheet.write(row_num, 1, row["E_Complex (a.u.)"])
                worksheet.write(row_num, 2, row["ΔE (a.u.)"])
                worksheet.write(row_num, 3, row["ΔE (kcal/mol)"])
            workbook.close()
            
            st.download_button(
                label="📥 Download Excel File",
                data=output.getvalue(),
                file_name="Interaction_Energies.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
