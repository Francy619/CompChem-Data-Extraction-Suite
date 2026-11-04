import streamlit as st
import os
import re
import xlsxwriter
import io
from decimal import Decimal
import pandas as pd

HARTREE_TO_KCAL = Decimal('627.509')

def extract_energy(uploaded_file):
    if uploaded_file is None: return None
    uploaded_file.seek(0)
    content = uploaded_file.read().decode('utf-8')
    stitched = re.sub(r'[\r\n]+ ', '', content)
    matches = re.findall(r'HF=([-+]?\d*\.\d+|\d+)', stitched)
    return Decimal(matches[-1]) if matches else None

st.set_page_config(page_title="Interaction Energy Calculator", layout="wide")
st.title("🔬 Universal Interaction Energy Calculator")
st.markdown("Calculate $\Delta E$ for any system. Adjust the **Quantity** to calculate complexes with multiple solvent molecules (e.g., 2 Waters, 3 Methanols).")

# --- UI: COMMON MONOMER ---
st.markdown("### 1. Main System (Common to all calculations)")
m1_file = st.file_uploader("Upload Monomer 1 (e.g., Main Amino Acid, Catalyst, etc.)", type=["log", "out"])

st.markdown("---")

# --- UI: DYNAMIC GROUPS WITH MULTIPLIERS ---
st.markdown("### 2. Secondary Systems (Solvents, Ligands, Bases, etc.)")
num_systems = st.number_input("How many different secondary systems are you calculating against?", min_value=1, max_value=20, value=1)

systems_data = []

for i in range(num_systems):
    with st.container():
        st.markdown(f"#### System Type {i+1}")
        col1, col2, col3 = st.columns([2, 1, 3])
        
        with col1:
            m2 = st.file_uploader(f"Upload Monomer 2", type=["log", "out"], key=f"m2_{i}")
        with col2:
            n_m2 = st.number_input(f"Quantity (n)", min_value=1, max_value=50, value=1, key=f"qty_{i}", 
                                   help="How many of this monomer are in the complex? The energy will be multiplied by this number.")
        with col3:
            comps = st.file_uploader(f"Upload all Complexes formed with Monomer 2", type=["log", "out"], accept_multiple_files=True, key=f"comps_{i}")
            
        systems_data.append({"m2": m2, "multiplier": Decimal(n_m2), "complexes": comps})
        st.write("") # Spacer

st.markdown("---")

# --- EXECUTE CALCULATIONS ---
if st.button("Calculate Interaction Energies", type="primary", use_container_width=True):
    if not m1_file:
        st.error("⚠️ Please upload Monomer 1.")
        st.stop()
        
    with st.spinner("Extracting energies and generating tabulations..."):
        e_m1 = extract_energy(m1_file)
        if e_m1 is None:
            st.error("Could not extract energy from Monomer 1.")
            st.stop()
            
        results = []
        is_first_in_sheet = True
        
        for idx, sys in enumerate(systems_data):
            m2_file = sys["m2"]
            n_m2 = sys["multiplier"]
            dimer_files = sys["complexes"]
            
            if not m2_file or not dimer_files:
                continue
                
            e_m2_single = extract_energy(m2_file)
            if e_m2_single is None:
                st.warning(f"Could not extract energy from {m2_file.name}. Skipping this group.")
                continue
            
            # Multiply the solvent energy by the quantity provided by the user
            e_m2_total = e_m2_single * n_m2
                
            group_results = []
            
            for d_file in dimer_files:
                e_dimer = extract_energy(d_file)
                if e_dimer is not None:
                    delta_au = e_dimer - (e_m1 + e_m2_total)
                    delta_kcal = delta_au * HARTREE_TO_KCAL
                    
                    group_results.append({
                        "Complex File": d_file.name,
                        "E Complex (a.u.)": e_dimer,
                        "Amino Acid File": m1_file.name,
                        "E Amino Acid (a.u.)": e_m1,
                        "Solvent File": m2_file.name,
                        "Qty": int(n_m2),
                        "E Solvent (single a.u.)": e_m2_single,
                        "ΔE (a.u.)": delta_au,
                        "ΔE (kcal/mol)": delta_kcal
                    })
            
            if group_results:
                min_kcal = min(r["ΔE (kcal/mol)"] for r in group_results)
                for i, r in enumerate(group_results):
                    r["Is_Best"] = (r["ΔE (kcal/mol)"] == min_kcal)
                    r["Is_First_In_Group"] = (i == 0)
                    r["Is_First_In_Sheet"] = is_first_in_sheet
                    results.append(r)
                    is_first_in_sheet = False

        if not results:
            st.error("No valid calculations could be completed. Please ensure you uploaded matching complex files for your secondary monomers.")
            st.stop()

        # --- UI: TABS FOR TABULATION & DOWNLOAD ---
        tab1, tab2 = st.tabs(["📊 Data Tabulation", "📥 Download Report"])
        
        with tab1:
            st.success("Calculations complete!")
            display_df = pd.DataFrame(results).drop(columns=["Is_Best", "Is_First_In_Sheet", "Is_First_In_Group"])
            for col in display_df.columns:
                if "E " in col or "ΔE" in col:
                    display_df[col] = display_df[col].astype(float)
            
            def highlight_best(row):
                idx = row.name
                if results[idx]["Is_Best"]:
                    return ['background-color: #FFFF00; color: black'] * len(row)
                return [''] * len(row)
            
            styled_df = display_df.style.apply(highlight_best, axis=1).format(precision=6)
            st.dataframe(styled_df, use_container_width=True)
            
        with tab2:
            st.markdown("### Generate Styled Excel File")
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet("Interaction Energies")
            
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            highlight_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00', 'font_name': 'Times New Roman', 'font_size': 12})
            
            columns = [
                "Complex File", "E Complex (a.u.)", 
                "Monomer 1 File", "E Monomer 1 (a.u.)", 
                "Monomer 2 File", "Qty", "E Monomer 2 (single a.u.)", 
                "ΔE (a.u.)", "ΔE (kcal/mol)"
            ]
            
            worksheet.set_row(0, 24)
            for col_num, header in enumerate(columns):
                worksheet.write(0, col_num, header, header_format)
                
            worksheet.set_column('A:B', 22)
            worksheet.set_column('C:D', 22)
            worksheet.set_column('E:E', 20)
            worksheet.set_column('F:F', 8)  # Qty column
            worksheet.set_column('G:G', 25) # E Monomer 2 column
            worksheet.set_column('H:I', 18)
            
            for row_num, row_data in enumerate(results, start=1):
                worksheet.set_row(row_num, 20)
                
                worksheet.write(row_num, 0, row_data["Complex File"], left_format)
                worksheet.write_number(row_num, 1, float(row_data["E Complex (a.u.)"]), center_format)
                
                if row_data["Is_First_In_Sheet"]:
                    worksheet.write(row_num, 2, row_data["Amino Acid File"], left_format)
                    worksheet.write_number(row_num, 3, float(row_data["E Amino Acid (a.u.)"]), center_format)
                else:
                    worksheet.write(row_num, 2, "", left_format)
                    worksheet.write(row_num, 3, "", center_format)
                    
                if row_data["Is_First_In_Group"]:
                    worksheet.write(row_num, 4, row_data["Solvent File"], left_format)
                    worksheet.write_number(row_num, 5, row_data["Qty"], center_format)
                    worksheet.write_number(row_num, 6, float(row_data["E Solvent (single a.u.)"]), center_format)
                else:
                    worksheet.write(row_num, 4, "", left_format)
                    worksheet.write(row_num, 5, "", center_format)
                    worksheet.write(row_num, 6, "", center_format)
                    
                worksheet.write_number(row_num, 7, float(row_data["ΔE (a.u.)"]), center_format)
                
                fmt = highlight_format if row_data["Is_Best"] else center_format
                worksheet.write_number(row_num, 8, float(row_data["ΔE (kcal/mol)"]), fmt)
                        
            workbook.close()
            
            base_name = m1_file.name.split('.')[0]
            download_name = f"{base_name}_Universal_Energies.xlsx"
            
            st.download_button(
                label=f"📥 Download {download_name}",
                data=output.getvalue(),
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
