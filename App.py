import streamlit as st
import os
import re
import xlsxwriter
import io
from decimal import Decimal
import pandas as pd

HARTREE_TO_KCAL = Decimal('627.509')

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================
def extract_energy(uploaded_file, is_bsse=False):
    if uploaded_file is None: return None
    uploaded_file.seek(0)
    content = uploaded_file.read().decode('utf-8')
    
    if is_bsse:
        matches = re.findall(r'Counterpoise:\s*corrected\s*energy\s*=\s*([-+]?\d*\.\d+)', content, re.IGNORECASE)
        return Decimal(matches[-1]) if matches else None
    else:
        stitched = re.sub(r'[\r\n]+ ', '', content)
        matches = re.findall(r'HF=([-+]?\d*\.\d+|\d+)', stitched)
        return Decimal(matches[-1]) if matches else None

def format_nbo(text):
    text = re.sub(r'^\s*\d+\.\s*', '', text)
    text = re.sub(r'\(\s+', '(', text)
    text = re.sub(r'\s+\)', ')', text)
    text = re.sub(r'([A-Za-z\*]+)\s+(\d+)', r'\1\2', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()

# ==========================================
# 2. PAGE CONFIG & CSS
# ==========================================
st.set_page_config(page_title="Computational Chemistry Dashboard", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    .main-header-container { text-align: center; margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 2px solid #E2E8F0; }
    .main-title { font-size: 3.2rem; color: #1E3A8A; font-weight: 800; margin-bottom: 5px; font-family: 'Segoe UI', sans-serif; }
    .sub-title { font-size: 1.1rem; color: #4B5563; font-family: 'Segoe UI', sans-serif; }
    .section-header { font-size: 1.1rem; color: #1E3A8A; font-weight: 700; margin-bottom: 10px; }
    .sidebar-title { font-size: 2.2rem; font-weight: 800; color: #1E3A8A; text-align: center; margin-top: -10px; margin-bottom: 20px; font-family: 'Segoe UI', sans-serif; }
    
    button[kind="primary"] { background-color: #1E3A8A !important; color: #FFFFFF !important; border: none !important; border-radius: 8px !important; font-weight: 700 !important; padding: 14px 24px !important; transition: all 0.3s ease-in-out !important; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important; }
    button[kind="primary"]:hover { background-color: #2563EB !important; transform: translateY(-2px); }
    button[kind="secondary"] { background-color: #F8FAFC !important; color: #475569 !important; border: 1px solid #CBD5E1 !important; border-radius: 8px !important; font-weight: 700 !important; padding: 14px 24px !important; transition: all 0.3s ease-in-out !important; }
    button[kind="secondary"]:hover { background-color: #E2E8F0 !important; color: #1E293B !important; border-color: #94A3B8 !important; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. SIDEBAR NAVIGATION
# ==========================================
col1, col2 = st.sidebar.columns(2)
with col1:
    try: st.image("logo1.png", use_container_width=True)
    except Exception: pass
with col2:
    try: st.image("logo2.png", width=60)
    except Exception: pass

st.sidebar.markdown("---")
st.sidebar.markdown('<p class="sidebar-title">📊 Dashboard</p>', unsafe_allow_html=True)

if 'app_mode' not in st.session_state: st.session_state.app_mode = "Interaction Energy Calculator"

btn_ie = st.sidebar.button("🔬 Interaction Energy", use_container_width=True, type="primary" if st.session_state.app_mode == "Interaction Energy Calculator" else "secondary")
btn_nbo = st.sidebar.button("⚛️ NBO E(2) Analysis", use_container_width=True, type="primary" if st.session_state.app_mode == "NBO E(2) Analysis" else "secondary")
btn_qtaim = st.sidebar.button("⭕ QTAIM Topology", use_container_width=True, type="primary" if st.session_state.app_mode == "QTAIM Analysis" else "secondary")

if btn_ie: st.session_state.app_mode = "Interaction Energy Calculator"; st.rerun()
if btn_nbo: st.session_state.app_mode = "NBO E(2) Analysis"; st.rerun()
if btn_qtaim: st.session_state.app_mode = "QTAIM Analysis"; st.rerun()

app_mode = st.session_state.app_mode
st.sidebar.markdown("---")

with st.sidebar.expander("ℹ️ About & Instructions"):
    st.markdown("""
    **Developed by:** Francis Thomas  
    **Institution:** St. Dominic's College, Kanjirapally  
    """)

# ==========================================
# 4. APP MODE 1: INTERACTION ENERGY
# ==========================================
if app_mode == "Interaction Energy Calculator":
    st.markdown("""
    <div class="main-header-container">
        <div class="main-title">🔬 Interaction Energy Calculator</div>
        <div class="sub-title">Automated extraction and computation of ΔE from Gaussian log files.</div>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns(2)
    with col_left:
        with st.container(border=True):
            st.markdown('<p class="section-header">1. Primary Reference System</p>', unsafe_allow_html=True)
            m1_file = st.file_uploader("Upload Monomer 1 (e.g., Amino Acid)", type=["log", "out"])

    with col_right:
        with st.container(border=True):
            st.markdown('<p class="section-header">2. Secondary Systems (Solvents, Ligands)</p>', unsafe_allow_html=True)
            num_systems = st.number_input("Number of secondary systems to evaluate:", min_value=1, max_value=20, value=1)
            bsse_choice = st.radio("BSSE Calculated?", ["No", "Yes"], horizontal=True)
            is_bsse = (bsse_choice == "Yes")

    st.write("") 
    systems_data = []

    for i in range(num_systems):
        with st.container(border=True):
            st.markdown(f"**Secondary System Group {i+1}**")
            col1, col2, col3 = st.columns([2, 1, 3])
            with col1: m2 = st.file_uploader(f"Upload Monomer 2", type=["log", "out"], key=f"m2_{i}")
            with col2: n_m2 = st.number_input(f"Quantity (n)", min_value=1, max_value=50, value=1, key=f"qty_{i}")
            with col3: comps = st.file_uploader(f"Upload corresponding Complexes", type=["log", "out"], accept_multiple_files=True, key=f"comps_{i}")
            systems_data.append({"m2": m2, "multiplier": Decimal(n_m2), "complexes": comps})

    st.divider()

    if 'calc_done_ie' not in st.session_state:
        st.session_state.calc_done_ie = False
        st.session_state.results_ie = []
        st.session_state.excel_data_ie = None
        st.session_state.download_name_ie = ""

    btn_placeholder = st.empty()
    do_calc = False

    with btn_placeholder.container():
        if not st.session_state.calc_done_ie:
            col_s1, col_btn, col_s2 = st.columns([1, 1, 1])
            with col_btn:
                if st.button("Calculate Energies", type="primary", use_container_width=True): do_calc = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Calculate Energies", type="primary", use_container_width=True): do_calc = True
            with col_down:
                st.download_button("Download Data (.xlsx)", data=st.session_state.excel_data_ie, file_name=st.session_state.download_name_ie, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_calc:
        if not m1_file:
            st.error("Please upload Monomer 1 to proceed.")
            st.session_state.calc_done_ie = False
        else:
            with st.spinner("Processing Gaussian outputs..."):
                e_m1 = extract_energy(m1_file, is_bsse=False)
                if e_m1 is None: st.stop()
                    
                results = []
                is_first_in_sheet = True
                
                for idx, sys in enumerate(systems_data):
                    m2_file = sys["m2"]
                    n_m2 = sys["multiplier"]
                    dimer_files = sys["complexes"]
                    
                    if not m2_file or not dimer_files: continue
                        
                    e_m2_single = extract_energy(m2_file, is_bsse=False) 
                    if e_m2_single is None: continue
                    
                    e_m2_total = e_m2_single * n_m2
                    group_results = []
                    
                    for d_file in dimer_files:
                        e_dimer = extract_energy(d_file, is_bsse=is_bsse)
                        if e_dimer is not None:
                            delta_au = e_dimer - (e_m1 + e_m2_total)
                            delta_kcal = delta_au * HARTREE_TO_KCAL
                            group_results.append({
                                "Complex File": d_file.name, "E Complex (a.u.)": e_dimer, "Monomer 1 File": m1_file.name,
                                "E Monomer 1 (a.u.)": e_m1, "Monomer 2 File": m2_file.name, "Qty": int(n_m2),
                                "E Monomer 2 (single a.u.)": e_m2_single, "ΔE (a.u.)": delta_au, "ΔE (kcal/mol)": delta_kcal
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
                    st.session_state.calc_done_ie = False
                    st.stop()

                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                worksheet = workbook.add_worksheet("Interaction Energies")
                
                header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
                center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
                left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
                highlight_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00', 'font_name': 'Times New Roman', 'font_size': 12})
                
                columns = ["Complex File", "E Complex (a.u.)", "Monomer 1 File", "E Monomer 1 (a.u.)", "Monomer 2 File", "Qty", "E Monomer 2 (single a.u.)", "ΔE (a.u.)", "ΔE (kcal/mol)"]
                worksheet.set_row(0, 24)
                for col_num, header in enumerate(columns): worksheet.write(0, col_num, header, header_format)
                worksheet.set_column('A:B', 22); worksheet.set_column('C:D', 22); worksheet.set_column('E:E', 20); worksheet.set_column('F:F', 8); worksheet.set_column('G:G', 25); worksheet.set_column('H:I', 18)
                
                for row_num, row_data in enumerate(results, start=1):
                    worksheet.set_row(row_num, 20)
                    worksheet.write(row_num, 0, row_data["Complex File"], left_format)
                    worksheet.write_number(row_num, 1, float(row_data["E Complex (a.u.)"]), center_format)
                    if row_data["Is_First_In_Sheet"]:
                        worksheet.write(row_num, 2, row_data["Monomer 1 File"], left_format)
                        worksheet.write_number(row_num, 3, float(row_data["E Monomer 1 (a.u.)"]), center_format)
                    else:
                        worksheet.write(row_num, 2, "", left_format); worksheet.write(row_num, 3, "", center_format)
                    if row_data["Is_First_In_Group"]:
                        worksheet.write(row_num, 4, row_data["Monomer 2 File"], left_format)
                        worksheet.write_number(row_num, 5, row_data["Qty"], center_format)
                        worksheet.write_number(row_num, 6, float(row_data["E Monomer 2 (single a.u.)"]), center_format)
                    else:
                        worksheet.write(row_num, 4, "", left_format); worksheet.write(row_num, 5, "", center_format); worksheet.write(row_num, 6, "", center_format)
                    worksheet.write_number(row_num, 7, float(row_data["ΔE (a.u.)"]), center_format)
                    fmt = highlight_format if row_data["Is_Best"] else center_format
                    worksheet.write_number(row_num, 8, float(row_data["ΔE (kcal/mol)"]), fmt)
                            
                workbook.close()
                st.session_state.results_ie = results
                st.session_state.excel_data_ie = output.getvalue()
                st.session_state.download_name_ie = f"{m1_file.name.split('.')[0]}_Energies.xlsx"
                st.session_state.calc_done_ie = True
                st.rerun()

    if st.session_state.calc_done_ie:
        st.write("") 
        st.markdown('<p class="section-header" style="text-align: center; border: none; margin-top: 20px;">Tabulated Results</p>', unsafe_allow_html=True)
        display_df = pd.DataFrame(st.session_state.results_ie).drop(columns=["Is_Best", "Is_First_In_Sheet", "Is_First_In_Group"])
        def highlight_best(row):
            idx = row.name
            if st.session_state.results_ie[idx]["Is_Best"]: return ['background-color: #FEF08A; color: #1F2937; font-weight: bold'] * len(row)
            return [''] * len(row)
        styled_df = display_df.style.apply(highlight_best, axis=1).format(precision=6)
        st.dataframe(styled_df, use_container_width=True)

# ==========================================
# 5. APP MODE 2: NBO E(2) ANALYSIS
# ==========================================
elif app_mode == "NBO E(2) Analysis":
    st.markdown("""
    <div class="main-header-container">
        <div class="main-title">⚛️ NBO E(2) Perturbation Analysis</div>
        <div class="sub-title">Extract and format Natural Bond Orbital donor-acceptor interactions.</div>
    </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown('<p class="section-header">Analysis Configuration</p>', unsafe_allow_html=True)
        col1, col2 = st.columns([1, 2])
        with col1: top_n = st.number_input("How many highest E(2) values needed per file?", min_value=1, max_value=50, value=1)
        with col2:
            allowed_sections = st.multiselect("Interaction types to consider:", ["within unit 1", "from unit 1 to unit 2", "from unit 2 to unit 1", "within unit 2"], default=["from unit 1 to unit 2", "from unit 2 to unit 1"])
        nbo_files = st.file_uploader("Upload NBO log files (.log, .out)", accept_multiple_files=True)

    st.divider()
    
    if 'calc_done_nbo' not in st.session_state:
        st.session_state.calc_done_nbo = False
        st.session_state.excel_data_nbo = None
        st.session_state.all_top_n = []
        st.session_state.file_data_dict = {}

    btn_placeholder = st.empty()
    do_nbo_run = False

    with btn_placeholder.container():
        if not st.session_state.calc_done_nbo:
            col_s1, col_btn, col_s2 = st.columns([1, 1, 1])
            with col_btn:
                if st.button("Run NBO Extraction", type="primary", use_container_width=True): do_nbo_run = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Run NBO Extraction", type="primary", use_container_width=True): do_nbo_run = True
            with col_down:
                st.download_button("Download NBO Report (.xlsx)", data=st.session_state.excel_data_nbo, file_name="NBO_Batch_Results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_nbo_run:
        if not nbo_files:
            st.error("Please upload at least one NBO file.")
            st.session_state.calc_done_nbo = False
        elif not allowed_sections:
            st.error("Please select at least one interaction type.")
            st.session_state.calc_done_nbo = False
        else:
            with st.spinner("Parsing NBO files..."):
                file_data_dict = {}
                all_top_n_data = []
                
                for uploaded_file in nbo_files:
                    uploaded_file.seek(0)
                    lines = uploaded_file.read().decode('utf-8').splitlines()
                    basename = os.path.splitext(uploaded_file.name)[0]
                    
                    file_extracted_data = []
                    in_target_section = False
                    current_section = "Unknown Section"
                    
                    for line in lines:
                        if "Second Order Perturbation Theory Analysis of Fock Matrix in NBO Basis" in line:
                            in_target_section = True
                            continue
                        if in_target_section:
                            if "Natural Bond Orbitals (Summary):" in line: break
                            line_stripped = line.strip()
                            if line_stripped.startswith("within unit") or line_stripped.startswith("from unit"):
                                current_section = re.sub(r'\s+', ' ', line_stripped)
                                continue
                            if '/' in line:
                                parts = line.split('/')
                                donor_raw = parts[0].strip()
                                right_part = parts[1].strip()
                                right_tokens = right_part.split()
                                if len(right_tokens) >= 4:
                                    e2_val = right_tokens[-3]
                                    e_diff_val = right_tokens[-2]
                                    f_ij_val = right_tokens[-1]
                                    acceptor_raw = " ".join(right_tokens[1:-3])
                                    donor_fmt = format_nbo(donor_raw)
                                    acceptor_fmt = format_nbo(acceptor_raw)
                                    
                                    if not any(term in donor_fmt for term in ["RY*", "CR"]) and not any(term in acceptor_fmt for term in ["RY*", "CR"]):
                                        file_extracted_data.append([basename, current_section, donor_fmt, acceptor_fmt, float(e2_val), float(e_diff_val), float(f_ij_val)])
                    
                    if file_extracted_data:
                        file_data_dict[basename] = file_extracted_data
                        filtered = [r for r in file_extracted_data if r[1] in allowed_sections]
                        top_n_rows = sorted(filtered, key=lambda r: r[4], reverse=True)[:top_n]
                        all_top_n_data.extend(top_n_rows)

                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
                center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
                left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
                section_format = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})

                if all_top_n_data:
                    ws_top = workbook.add_worksheet(f'Top {top_n} Interactions')
                    top_n_columns = ["System", "Interaction Section", "Donor NBO (i)", "Acceptor NBO (j)", "E(2) kcal/mol", "E(j)-E(i) a.u.", "F(i,j) a.u."]
                    ws_top.set_row(0, 20)
                    for col_num, header in enumerate(top_n_columns): ws_top.write(0, col_num, header, header_format)
                    ws_top.set_column('A:A', 20); ws_top.set_column('B:B', 30); ws_top.set_column('C:D', 35); ws_top.set_column('E:G', 15)
                    for row_num, row_data in enumerate(all_top_n_data, start=1):
                        ws_top.set_row(row_num, 20)
                        for col_num in range(4): ws_top.write(row_num, col_num, row_data[col_num], left_format)
                        for col_num in range(4, 7): ws_top.write_number(row_num, col_num, row_data[col_num], center_format)

                indiv_columns = ["Donor NBO (i)", "Acceptor NBO (j)", "E(2) kcal/mol", "E(j)-E(i) a.u.", "F(i,j) a.u."]
                for basename, data_rows in file_data_dict.items():
                    sheet_name = basename[:31] 
                    ws_file = workbook.add_worksheet(sheet_name)
                    ws_file.set_row(0, 20)
                    for col_num, header in enumerate(indiv_columns): ws_file.write(0, col_num, header, header_format)
                    ws_file.set_column('A:B', 35); ws_file.set_column('C:E', 15)
                    current_excel_row = 1
                    last_section = None
                    for row_data in data_rows:
                        current_section = row_data[1]
                        if current_section != last_section:
                            ws_file.set_row(current_excel_row, 20)
                            ws_file.write(current_excel_row, 0, current_section, section_format)
                            current_excel_row += 1
                            last_section = current_section
                        ws_file.set_row(current_excel_row, 20)
                        ws_file.write(current_excel_row, 0, row_data[2], left_format) 
                        ws_file.write(current_excel_row, 1, row_data[3], left_format) 
                        ws_file.write_number(current_excel_row, 2, row_data[4], center_format)
                        ws_file.write_number(current_excel_row, 3, row_data[5], center_format)
                        ws_file.write_number(current_excel_row, 4, row_data[6], center_format)
                        current_excel_row += 1

                workbook.close()
                st.session_state.all_top_n = all_top_n_data
                st.session_state.excel_data_nbo = output.getvalue()
                st.session_state.calc_done_nbo = True
                st.rerun()

    if st.session_state.calc_done_nbo:
        st.write("") 
        st.markdown(f'<p class="section-header" style="text-align: center; border: none;">Top {top_n} Selected Interactions</p>', unsafe_allow_html=True)
        if st.session_state.all_top_n:
            df = pd.DataFrame(st.session_state.all_top_n, columns=["System", "Interaction Section", "Donor (i)", "Acceptor (j)", "E(2) kcal/mol", "E(j)-E(i) a.u.", "F(i,j) a.u."])
            st.dataframe(df, use_container_width=True)

# ==========================================
# 6. APP MODE 3: QTAIM ANALYSIS
# ==========================================
elif app_mode == "QTAIM Analysis":
    st.markdown("""
    <div class="main-header-container">
        <div class="main-title">⭕ QTAIM Topology Analysis</div>
        <div class="sub-title">Parse Multiwfn CPprop.txt files to extract density properties at Bond Critical Points.</div>
    </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown('<p class="section-header">QTAIM Configuration</p>', unsafe_allow_html=True)
        num_qtaim_files = st.number_input("How many CPprop output files are you analyzing?", min_value=1, max_value=20, value=1)
        
    st.write("")
    qtaim_inputs = []
    
    for i in range(num_qtaim_files):
        with st.container(border=True):
            st.markdown(f"**MultiWFN Output {i+1}**")
            col1, col2 = st.columns([1, 2])
            with col1:
                q_file = st.file_uploader(f"Upload CPprop.txt", type=["txt"], key=f"q_file_{i}")
            with col2:
                cp_string = st.text_input("Enter Target Critical Points (comma-separated, e.g. 15, 23, 44)", key=f"q_cps_{i}")
            
            qtaim_inputs.append({"file": q_file, "cp_str": cp_string})

    st.divider()

    if 'calc_done_qtaim' not in st.session_state:
        st.session_state.calc_done_qtaim = False
        st.session_state.qtaim_results = []
        st.session_state.excel_data_qtaim = None

    btn_placeholder = st.empty()
    do_qtaim_run = False

    with btn_placeholder.container():
        if not st.session_state.calc_done_qtaim:
            col_s1, col_btn, col_s2 = st.columns([1, 1, 1])
            with col_btn:
                if st.button("Run QTAIM Extraction", type="primary", use_container_width=True): do_qtaim_run = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Run QTAIM Extraction", type="primary", use_container_width=True): do_qtaim_run = True
            with col_down:
                st.download_button("Download QTAIM Report (.xlsx)", data=st.session_state.excel_data_qtaim, file_name="QTAIM_Results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_qtaim_run:
        search_terms = {
            "Density of all electrons": "Density of all electrons:",
            "Laplacian of electron density": "Laplacian of electron density:",
            "Potential energy density V(r)": "Potential energy density V(r):",
            "Lagrangian kinetic energy G(r)": "Lagrangian kinetic energy G(r):",
            "Energy density E(r) or H(r)": "Energy density E(r) or H(r):"
        }
        
        all_extracted_data = []
        has_files = False
        
        with st.spinner("Parsing Multiwfn Outputs..."):
            for item in qtaim_inputs:
                uploaded_file = item["file"]
                cp_str = item["cp_str"]
                
                if not uploaded_file: continue
                has_files = True
                
                try:
                    target_cps = [int(x.strip()) for x in cp_str.split(',') if x.strip().isdigit()]
                except Exception:
                    st.error(f"Invalid CP format for {uploaded_file.name}. Ensure you use numbers separated by commas.")
                    st.session_state.calc_done_qtaim = False
                    st.stop()
                
                if not target_cps:
                    st.warning(f"No valid CPs entered for {uploaded_file.name}. Skipping.")
                    continue
                
                uploaded_file.seek(0)
                lines = uploaded_file.read().decode('utf-8').splitlines()
                
                basename = os.path.splitext(uploaded_file.name)[0]
                system_name = re.sub(r'^[cC][pP][pP][rR][oO][pP]', '', basename)
                system_name = re.sub(r'[\d,\s]+$', '', system_name).strip()
                if not system_name: system_name = basename
                
                extracted_data = []
                current_cp = None
                cp_data = {}
                
                for line in lines:
                    if "----------------   CP" in line:
                        if current_cp in target_cps and cp_data:
                            extracted_data.append(cp_data)
                        
                        match = re.search(r'CP\s+(\d+)', line)
                        if match:
                            current_cp = int(match.group(1))
                            cp_data = {"System": system_name, "CP Number": current_cp}
                        else:
                            current_cp = None
                        continue
                        
                    if current_cp in target_cps:
                        for key, term in search_terms.items():
                            if line.lstrip().startswith(term):
                                try:
                                    val_str = line.split(term)[1].strip()
                                    cp_data[key] = float(val_str)
                                except Exception:
                                    cp_data[key] = None

                if current_cp in target_cps and cp_data:
                    extracted_data.append(cp_data)
                
                for data in extracted_data:
                    density = data.get("Density of all electrons")
                    if density is not None:
                        data["Binding Energy"] = (-223.08 * density) + 0.7423
                    else:
                        data["Binding Energy"] = None

                if extracted_data:
                    extracted_data.sort(key=lambda x: x["CP Number"])
                    all_extracted_data.extend(extracted_data)

            if not has_files:
                st.error("Please upload at least one CPprop.txt file.")
                st.session_state.calc_done_qtaim = False
                st.stop()
            elif not all_extracted_data:
                st.error("No targeted CPs were found in the uploaded files.")
                st.session_state.calc_done_qtaim = False
                st.stop()

            # --- EXCEL GENERATION (QTAIM) ---
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet("QTAIM Data")

            header_base = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            prop_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12, 'num_format': '0.0000'})
            be_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12, 'num_format': '0.00'})

            sup = workbook.add_format({'bold': True, 'font_script': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            sub = workbook.add_format({'bold': True, 'font_script': 2, 'font_name': 'Times New Roman', 'font_size': 12})
            bold_fmt = workbook.add_format({'bold': True, 'font_name': 'Times New Roman', 'font_size': 12})

            worksheet.set_row(0, 20)
            worksheet.write(0, 0, "System", header_format)
            worksheet.write(0, 1, "CP", header_format)
            worksheet.write(0, 2, "ρ(r)", header_format)
            worksheet.write_rich_string(0, 3, bold_fmt, "Δ", sup, "2", bold_fmt, "ρ", header_base)
            worksheet.write_rich_string(0, 4, bold_fmt, "V(r", sub, "bcp", bold_fmt, ")", header_base)
            worksheet.write(0, 5, "G(rcp)", header_format)
            worksheet.write(0, 6, "H(r)", header_format)
            worksheet.write(0, 7, "B. E (kcal/Mol)", header_format)

            worksheet.set_column('A:A', 22); worksheet.set_column('B:B', 8); worksheet.set_column('C:G', 14); worksheet.set_column('H:H', 18)

            keys_in_order = [
                "Density of all electrons", "Laplacian of electron density", 
                "Potential energy density V(r)", "Lagrangian kinetic energy G(r)", "Energy density E(r) or H(r)"
            ]

            start_row = 1
            current_system = all_extracted_data[0]["System"]

            for row_num, data in enumerate(all_extracted_data, start=1):
                worksheet.set_row(row_num, 20)
                if data["System"] != current_system:
                    if row_num - 1 > start_row: worksheet.merge_range(start_row, 0, row_num - 1, 0, current_system, left_format)
                    else: worksheet.write(start_row, 0, current_system, left_format)
                    start_row = row_num
                    current_system = data["System"]
                    
                worksheet.write_number(row_num, 1, data["CP Number"], center_format)
                for col_offset, key in enumerate(keys_in_order):
                    val = data.get(key)
                    if val is not None: worksheet.write_number(row_num, 2 + col_offset, val, prop_format)
                    else: worksheet.write(row_num, 2 + col_offset, "N/A", center_format)
                        
                be = data.get("Binding Energy")
                if be is not None: worksheet.write_number(row_num, 7, be, be_format)
                else: worksheet.write(row_num, 7, "N/A", be_format)

            last_row = len(all_extracted_data)
            if last_row > start_row: worksheet.merge_range(start_row, 0, last_row, 0, current_system, left_format)
            else: worksheet.write(start_row, 0, current_system, left_format)

            workbook.close()
            st.session_state.qtaim_results = all_extracted_data
            st.session_state.excel_data_qtaim = output.getvalue()
            st.session_state.calc_done_qtaim = True
            st.rerun()

    if st.session_state.calc_done_qtaim:
        st.write("") 
        st.markdown('<p class="section-header" style="text-align: center; border: none; margin-top: 20px;">Extracted QTAIM Properties</p>', unsafe_allow_html=True)
        
        # Display simplified data in web UI
        display_data = []
        for d in st.session_state.qtaim_results:
            row = {
                "System": d["System"], "CP": d["CP Number"],
                "ρ(r)": d.get("Density of all electrons"),
                "Δ²ρ": d.get("Laplacian of electron density"),
                "V(r_bcp)": d.get("Potential energy density V(r)"),
                "G(rcp)": d.get("Lagrangian kinetic energy G(r)"),
                "H(r)": d.get("Energy density E(r) or H(r)"),
                "Binding Energy": d.get("Binding Energy")
            }
            display_data.append(row)
            
        st.dataframe(pd.DataFrame(display_data), use_container_width=True)
