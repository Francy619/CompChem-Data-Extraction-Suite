import streamlit as st
import re
import xlsxwriter
import io
from decimal import Decimal
import pandas as pd

HARTREE_TO_KCAL = Decimal('627.509')

def extract_energy(uploaded_file):
    if uploaded_file is None: return None
    # Reset file pointer just in case it was read before
    uploaded_file.seek(0)
    content = uploaded_file.read().decode('utf-8')
    stitched = re.sub(r'[\r\n]+ ', '', content)
    matches = re.findall(r'HF=([-+]?\d*\.\d+|\d+)', stitched)
    return Decimal(matches[-1]) if matches else None

st.set_page_config(page_title="Interaction Energy Calculator", layout="wide")
st.title("🔬 Interaction Energy Calculator")
st.markdown("Calculate $\Delta E$ for complexes. Upload your Gaussian output (`.log` or `.out`) files below.")

# --- UI: FILE UPLOADS ---
st.markdown("### 1. Reference Monomers")
col1, col2 = st.columns(2)
with col1:
    m1_file = st.file_uploader("Upload Monomer 1 (e.g., Amino Acid)", type=["log", "out"])
with col2:
    m2_file = st.file_uploader("Upload Monomer 2 (e.g., Solvent)", type=["log", "out"])

st.markdown("### 2. Complex / Dimer Files")
dimer_files = st.file_uploader(
    "Upload all Complex files (You can select multiple files at once)", 
    type=["log", "out"], 
    accept_multiple_files=True
)

st.markdown("---")

if st.button("Calculate Interaction Energies", type="primary", use_container_width=True):
    if not (m1_file and m2_file and dimer_files):
        st.error("⚠️ Please upload Monomer 1, Monomer 2, and at least one Dimer file to proceed.")
    else:
        with st.spinner("Extracting energies and calculating..."):
            e_m1 = extract_energy(m1_file)
            e_m2 = extract_energy(m2_file)
            
            if e_m1 is None or e_m2 is None:
                st.error("Could not extract energy from one of the reference monomers. Please check the files.")
                st.stop()
            
            results = []
            for d_file in dimer_files:
                e_dimer = extract_energy(d_file)
                if e_dimer is not None:
                    delta_au = e_dimer - (e_m1 + e_m2)
                    delta_kcal = delta_au * HARTREE_TO_KCAL
                    results.append({
                        "Complex File": d_file.name,
                        "E Complex (a.u.)": e_dimer,
                        "Amino Acid File": m1_file.name,
                        "E Amino Acid (a.u.)": e_m1,
                        "Solvent File": m2_file.name,
                        "E Solvent (a.u.)": e_m2,
                        "ΔE (a.u.)": delta_au,
                        "ΔE (kcal/mol)": delta_kcal
                    })
            
            if not results:
                st.warning("No valid energies were extracted from the uploaded complex files.")
                st.stop()
                
            # Determine the lowest (best) interaction energy for highlighting
            min_kcal = min(r["ΔE (kcal/mol)"] for r in results)
            for i, r in enumerate(results):
                r["Is_Best"] = (r["ΔE (kcal/mol)"] == min_kcal)
                r["Is_First_In_Sheet"] = (i == 0)
                r["Is_First_In_Group"] = (i == 0)

            # --- UI: TABS FOR TABULATION & DOWNLOAD ---
            tab1, tab2 = st.tabs(["📊 Data Tabulation", "📥 Download Report"])
            
            with tab1:
                st.success("Calculations complete! Review the results below.")
                # Prepare a simplified dataframe for web display
                display_df = pd.DataFrame(results).drop(columns=["Is_Best", "Is_First_In_Sheet", "Is_First_In_Group"])
                # Convert Decimals to floats for streamlit display
                for col in display_df.columns:
                    if "E" in col:
                        display_df[col] = display_df[col].astype(float)
                
                # Apply basic highlight to the web dataframe
                def highlight_min(s):
                    is_min = s == s.min()
                    return ['background-color: #FFFF00; color: black' if v else '' for v in is_min]
                
                styled_df = display_df.style.apply(highlight_min, subset=['ΔE (kcal/mol)']).format(precision=6)
                st.dataframe(styled_df, use_container_width=True)
                
            with tab2:
                st.markdown("### Generate Styled Excel File")
                st.info("The downloaded file includes professional formatting (Times New Roman, borders, aligned cells) and highlights the most stable conformer.")
                
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                worksheet = workbook.add_worksheet("Interaction Energies")
                
                # --- EXCEL FORMATTING ---
                header_format = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1,
                    'font_name': 'Times New Roman', 'font_size': 12
                })
                center_format = workbook.add_format({
                    'align': 'center', 'valign': 'vcenter', 'border': 1,
                    'font_name': 'Times New Roman', 'font_size': 12
                })
                left_format = workbook.add_format({
                    'align': 'left', 'valign': 'vcenter', 'border': 1,
                    'font_name': 'Times New Roman', 'font_size': 12
                })
                highlight_format = workbook.add_format({
                    'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1,
                    'bg_color': '#FFFF00',
                    'font_name': 'Times New Roman', 'font_size': 12
                })
                
                columns = [
                    "Complex File", "E Complex (a.u.)", 
                    "Amino Acid File", "E Amino Acid (a.u.)", 
                    "Solvent File", "E Solvent (a.u.)", 
                    "ΔE (a.u.)", "ΔE (kcal/mol)"
                ]
                
                worksheet.set_row(0, 24)
                for col_num, header in enumerate(columns):
                    worksheet.write(0, col_num, header, header_format)
                    
                worksheet.set_column('A:A', 22)
                worksheet.set_column('B:B', 22)
                worksheet.set_column('C:C', 20)
                worksheet.set_column('D:D', 22)
                worksheet.set_column('E:E', 20)
                worksheet.set_column('F:F', 22)
                worksheet.set_column('G:H', 18)
                
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
                        worksheet.write_number(row_num, 5, float(row_data["E Solvent (a.u.)"]), center_format)
                    else:
                        worksheet.write(row_num, 4, "", left_format)
                        worksheet.write(row_num, 5, "", center_format)
                        
                    worksheet.write_number(row_num, 6, float(row_data["ΔE (a.u.)"]), center_format)
                    
                    fmt = highlight_format if row_data["Is_Best"] else center_format
                    worksheet.write_number(row_num, 7, float(row_data["ΔE (kcal/mol)"]), fmt)
                            
                workbook.close()
                
                # Setup dynamic download filename
                base_name = m1_file.name.split('.')[0]
                solv_name = m2_file.name.split('.')[0]
                download_name = f"{base_name}_{solv_name}_Energies.xlsx"
                
                st.download_button(
                    label=f"📥 Download {download_name}",
                    data=output.getvalue(),
                    file_name=download_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
