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

st.set_page_config(page_title="Interaction Energy Calculator", layout="wide", initial_sidebar_state="collapsed")

# --- CUSTOM CSS FOR PROFESSIONAL LOOK & BUTTONS ---
st.markdown("""
<style>
    /* Centered Main Title */
    .main-header-container {
        text-align: center;
        margin-bottom: 2rem;
        padding-bottom: 1rem;
        border-bottom: 2px solid #E2E8F0;
    }
    .main-title {
        font-size: 3.2rem;
        color: #1E3A8A;
        font-weight: 800;
        margin-bottom: 5px;
        font-family: 'Segoe UI', sans-serif;
    }
    .sub-title {
        font-size: 1.1rem;
        color: #4B5563;
        font-family: 'Segoe UI', sans-serif;
    }
    
    /* Section Headers inside frames */
    .section-header { 
        font-size: 1.1rem; 
        color: #1E3A8A; 
        font-weight: 700; 
        margin-bottom: 10px; 
    }
    
    /* Custom Button Styling (Blue with Hover) */
    div.stButton > button, div.stDownloadButton > button {
        background-color: #1E3A8A !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 10px 24px !important;
        transition: all 0.3s ease-in-out !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
    }
    
    /* Hover Effect for Buttons */
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        background-color: #2563EB !important; /* Lighter vivid blue on hover */
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2) !important;
        color: #FFFFFF !important;
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

# --- HEADER SECTION ---
st.markdown("""
<div class="main-header-container">
    <div class="main-title">🔬 Interaction Energy Calculator</div>
    <div class="sub-title">Automated extraction and computation of ΔE from Gaussian log files.</div>
</div>
""", unsafe_allow_html=True)

with st.expander("View Quick Guide"):
    st.markdown("""
    **Workflow Instructions:**
    1. **Primary System:** Upload your central monomer (e.g., Amino Acid).
    2. **Secondary Systems:** Specify how many different solvents or ligands you are testing.
    3. **Upload & Map:** For each secondary system, upload the pure Monomer 2 file, set its stoichiometric quantity ($n$), and upload the associated complex files.
    4. **Calculate:** The tool applies the formula: $\Delta E = E_{\text{Complex}} - (E_{\text{Monomer1}} + n \\times E_{\text{Monomer2}})$.
    """)

# --- TOP ROW: FRAMED CONFIGURATION ---
col_left, col_right = st.columns(2)

with col_left:
    with st.container(border=True):
        st.markdown('<p class="section-header">1. Primary Reference System</p>', unsafe_allow_html=True)
        m1_file = st.file_uploader("Upload Monomer 1 (e.g., Amino Acid)", type=["log", "out"])

with col_right:
    with st.container(border=True):
        st.markdown('<p class="section-header">2. Secondary Systems (Solvents, Ligands, Bases)</p>', unsafe_allow_html=True)
        num_systems = st.number_input("Number of secondary systems to evaluate:", min_value=1, max_value=20, value=1)

# --- DYNAMIC GROUPS WITH MULTIPLIERS (Also Framed) ---
st.write("") # Spacer
systems_data = []

for i in range(num_systems):
    with st.container(border=True):
        st.markdown(f"**System Type {i+1}**")
        col1, col2, col3 = st.columns([2, 1, 3])
        
        with col1:
            m2 = st.file_uploader(f"Upload Monomer 2", type=["log", "out"], key=f"m2_{i}")
        with col2:
            n_m2 = st.number_input(f"Quantity (n)", min_value=1, max_value=50, value=1, key=f"qty_{i}")
        with col3:
            comps = st.file_uploader(f"Upload corresponding Complexes", type=["log", "out"], accept_multiple_files=True, key=f"comps_{i}")
            
        systems_data.append({"m2": m2, "multiplier": Decimal(n_m2), "complexes": comps})

st.divider()

# --- INITIALIZE SESSION STATE ---
if 'calc_done' not in st.session_state:
    st.session_state.calc_done = False
    st.session_state.results = []
    st.session_state.excel_data = None
    st.session_state.download_name = ""

# --- BUTTON LAYOUT (CENTERED) ---
col_spacer1, col_calc, col_down, col_spacer2 = st.columns([2, 3, 3, 2])

with col_calc:
    calc_btn = st.button("Calculate Energies", use_container_width=True)

# --- EXECUTE CALCULATIONS ---
if calc_btn:
    if not m1_file:
        st.error("Please upload Monomer 1 to proceed.")
        st.session_state.calc_done = False
    else:
        with st.spinner("Processing Gaussian outputs..."):
            e_m1 = extract_energy(m1_file)
            if e_m1 is None:
                st.error("Failed to extract HF energy from Monomer 1.")
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
                    st.warning(f"Failed to extract HF energy from {m2_file.name}. Skipping.")
                    continue
                
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
                            "Monomer 1 File": m1_file.name,
                            "E Monomer 1 (a.u.)": e_m1,
                            "Monomer 2 File": m2_file.name,
                            "Qty": int(n_m2),
                            "E Monomer 2 (single a.u.)": e_m2_single,
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
                st.error("Calculations aborted. No valid interaction energies could be extracted.")
                st.session_state.calc_done = False
                st.stop()

            # --- GENERATE EXCEL BUFFER ---
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
            worksheet.set_column('F:F', 8)  
            worksheet.set_column('G:G', 25) 
            worksheet.set_column('H:I', 18)
            
            for row_num, row_data in enumerate(results, start=1):
                worksheet.set_row(row_num, 20)
                
                worksheet.write(row_num, 0, row_data["Complex File"], left_format)
                worksheet.write_number(row_num, 1, float(row_data["E Complex (a.u.)"]), center_format)
                
                if row_data["Is_First_In_Sheet"]:
                    worksheet.write(row_num, 2, row_data["Monomer 1 File"], left_format)
                    worksheet.write_number(row_num, 3, float(row_data["E Monomer 1 (a.u.)"]), center_format)
                else:
                    worksheet.write(row_num, 2, "", left_format)
                    worksheet.write(row_num, 3, "", center_format)
                    
                if row_data["Is_First_In_Group"]:
                    worksheet.write(row_num, 4, row_data["Monomer 2 File"], left_format)
                    worksheet.write_number(row_num, 5, row_data["Qty"], center_format)
                    worksheet.write_number(row_num, 6, float(row_data["E Monomer 2 (single a.u.)"]), center_format)
                else:
                    worksheet.write(row_num, 4, "", left_format)
                    worksheet.write(row_num, 5, "", center_format)
                    worksheet.write(row_num, 6, "", center_format)
                    
                worksheet.write_number(row_num, 7, float(row_data["ΔE (a.u.)"]), center_format)
                
                fmt = highlight_format if row_data["Is_Best"] else center_format
                worksheet.write_number(row_num, 8, float(row_data["ΔE (kcal/mol)"]), fmt)
                        
            workbook.close()
            
            st.session_state.results = results
            st.session_state.excel_data = output.getvalue()
            
            base_name = m1_file.name.split('.')[0]
            st.session_state.download_name = f"{base_name}_Universal_Energies.xlsx"
            st.session_state.calc_done = True

# --- SHOW TABLE AND DOWNLOAD BUTTON ---
if st.session_state.calc_done:
    with col_down:
        st.download_button(
            label="Download Data (.xlsx)",
            data=st.session_state.excel_data,
            file_name=st.session_state.download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    st.write("") 
    st.markdown('<p class="section-header" style="text-align: center; border: none; margin-top: 20px;">Tabulated Results</p>', unsafe_allow_html=True)
    
    display_df = pd.DataFrame(st.session_state.results).drop(columns=["Is_Best", "Is_First_In_Sheet", "Is_First_In_Group"])
    for col in display_df.columns:
        if "E " in col or "ΔE" in col:
            display_df[col] = display_df[col].astype(float)
    
    def highlight_best(row):
        idx = row.name
        if st.session_state.results[idx]["Is_Best"]:
            return ['background-color: #FEF08A; color: #1F2937; font-weight: bold'] * len(row)
        return [''] * len(row)
    
    styled_df = display_df.style.apply(highlight_best, axis=1).format(precision=6)
    st.dataframe(styled_df, use_container_width=True)
