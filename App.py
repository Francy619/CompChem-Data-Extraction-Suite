import streamlit as st
import os
import re
import xlsxwriter
import io
from decimal import Decimal
import pandas as pd
import streamlit.components.v1 as components

HARTREE_TO_KCAL = Decimal('627.509')

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================
def extract_energy(uploaded_file, is_bsse=False):
    if uploaded_file is None: return None
    uploaded_file.seek(0)
    content = uploaded_file.read().decode('utf-8')
    if is_bsse:
        matches = re.findall(r'Counterpoise\s*corrected\s*energy\s*=\s*([-+]?\d*\.\d+)', content, re.IGNORECASE)
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

def extract_led_values(uploaded_file, target_strings):
    extracted = {}
    if uploaded_file is None: return extracted
    uploaded_file.seek(0)
    lines = uploaded_file.read().decode('utf-8').splitlines()
    for line in lines:
        for term in target_strings:
            if term.lower() in line.lower():
                if term.lower().startswith("dispersion") and "non" in line.lower(): continue
                parts = line.strip().split()
                try:
                    if term == "Intra REF. energy":
                        extracted["Intra REF. energy of Fragment 1"] = Decimal(parts[3])
                        extracted["Intra REF. energy of Fragment 2"] = Decimal(parts[4])
                    else: extracted[term] = Decimal(parts[-1])
                except (ValueError, IndexError): continue
    return extracted

# ==========================================
# 2. PAGE CONFIG & CSS (Global)
# ==========================================
st.set_page_config(page_title="Computational Chemistry Dashboard", layout="wide", initial_sidebar_state="collapsed")

# ---------------------------------------------------------
# CRITICAL FIX: Initialize Session State before rendering
# ---------------------------------------------------------
if 'app_mode' not in st.session_state: 
    st.session_state.app_mode = "Home"
# ---------------------------------------------------------

st.markdown("""
<style>
    /* --- Hide Streamlit Branding & Overlays --- */
    footer, [data-testid="stFooter"], .stStatusWidget {
        display: none !important;
        visibility: hidden !important;
        height: 0px !important;
        max-height: 0px !important;
    }
    button[title="View fullscreen"] { display: none !important; }
    header { visibility: hidden !important; display: none !important; }
    
    /* Hide sidebar completely for full screen feel */
    [data-testid="collapsedControl"] { display: none; }
    
    .main-header-container { text-align: center; margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 2px solid #E2E8F0; }
    .main-title { font-size: 3.2rem; color: #1E3A8A; font-weight: 800; margin-bottom: 5px; font-family: 'Segoe UI', sans-serif; }
    .sub-title { font-size: 1.1rem; color: #4B5563; font-family: 'Segoe UI', sans-serif; }
    .section-header { font-size: 1.1rem; color: #1E3A8A; font-weight: 700; margin-bottom: 10px; }
    
    /* Global Action Buttons (Primary) - NORMAL SIZE */
    button[kind="primary"] { 
        background-color: #1E3A8A !important; 
        color: #FFFFFF !important; 
        border: none !important; 
        border-radius: 6px !important; 
        font-weight: 700 !important; 
        padding: 8px 16px !important; 
        transition: all 0.3s ease-in-out !important; 
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1) !important; 
    }
    button[kind="primary"]:hover { 
        background-color: #2563EB !important; 
        transform: translateY(-2px); 
    }
    
    /* Global Secondary Buttons (Upload, Back to Home) - NORMAL SIZE */
    button[kind="secondary"] {
        border-radius: 6px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease-in-out !important;
    }
    button[kind="secondary"]:hover {
        border-color: #1E3A8A !important;
        color: #1E3A8A !important;
    }
    
    /* Force File Uploader Elements into a Single Horizontal Row */
    [data-testid="stFileUploaderDropzone"] > div {
        display: flex !important;
        flex-direction: row !important;
        align-items: center !important;
        flex-wrap: nowrap !important;
        gap: 10px !important;
        padding: 10px !important;
    }
    
    /* Prevent the helper text from breaking into multiple lines */
    [data-testid="stFileUploaderDropzone"] small {
        white-space: nowrap !important;
        margin-top: 0 !important;
        font-size: 0.85rem !important;
    }
</style>
""", unsafe_allow_html=True)

# JavaScript injection to aggressively remove iframe footers
components.html("""
    <script>
        const removeFooter = setInterval(() => {
            const footer = window.parent.document.querySelectorAll('footer, [data-testid="stFooter"]');
            footer.forEach(el => el.style.display = 'none');
        }, 500);
    </script>
""", height=0)


# ==========================================
# 3. HOME DASHBOARD (FULL SCREEN)
# ==========================================
if st.session_state.app_mode == "Home":
    
    # Inject MASSIVE button CSS *only* for the Home Page
    st.markdown("""
    <style>
        .menu-title { 
            font-size: 4.8rem; font-weight: 600; color: #1E3A8A; text-align: center; 
            margin-top: 20px; margin-bottom: 40px; 
            font-family: 'Inter', 'Roboto', 'Helvetica Neue', sans-serif; 
            line-height: 1.1; letter-spacing: 2px; 
        }
        .footer-text { text-align: center; color: #64748B; font-size: 1rem; margin-top: 80px; border-top: 1px solid #E2E8F0; padding-top: 30px;}
        
        /* Override secondary buttons ONLY on this page to make them giant blue blocks */
        button[kind="secondary"] { 
            background-color: #1E3A8A !important; 
            color: #FFFFFF !important; 
            border: none !important; 
            border-radius: 12px !important; 
            font-weight: 800 !important; 
            font-size: 1.4rem !important; 
            min-height: 140px !important;
            padding: 30px 20px !important; 
            box-shadow: 0 6px 12px rgba(30, 58, 138, 0.2) !important;
        }
        button[kind="secondary"]:hover { 
            background-color: #2563EB !important; 
            color: #FFFFFF !important; 
            transform: translateY(-6px) !important;
            box-shadow: 0 12px 20px rgba(37, 99, 235, 0.3) !important;
        }
    </style>
    """, unsafe_allow_html=True)
        
    st.markdown('<p class="menu-title">CompChem Data Extraction Suite</p>', unsafe_allow_html=True)
    st.write("")
    
    # Centered 2x2 Grid for Buttons
    c_space_l, c_btn_1, c_btn_2, c_space_r = st.columns([1, 3, 3, 1])
    
    with c_btn_1:
        if st.button("🔬 Interaction Energy Extractor", use_container_width=True, type="secondary"): 
            st.session_state.app_mode = "Interaction Energy"
            st.rerun()
        st.write("")
        if st.button("⚛️ NBO E(2) Extractor", use_container_width=True, type="secondary"): 
            st.session_state.app_mode = "NBO E(2) Analysis"
            st.rerun()
            
    with c_btn_2:
        if st.button("📊 LED Data Parser", use_container_width=True, type="secondary"): 
            st.session_state.app_mode = "LED Analysis"
            st.rerun()
        st.write("")
        if st.button("⭕ QTAIM Properties Tabulator", use_container_width=True, type="secondary"): 
            st.session_state.app_mode = "QTAIM Analysis"
            st.rerun()
            
    st.markdown("""
    <div class="footer-text">
        <strong>Developed by:</strong> Francis Thomas & Dr. Renjith Thomas <br>
        <strong>Research Group:</strong> Centre for Theoretical and Computational Chemistry (CTCC) <br>
        <strong>Institutions:</strong> St. Dominic's College, Kanjirapally & St. Berchmans College, Changanassery <br>
        <br>
        Automated tools for parsing outputs from Gaussian, ORCA, and Multiwfn.
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# 4. APP MODE 1: INTERACTION ENERGY
# ==========================================
elif st.session_state.app_mode == "Interaction Energy":
    st.markdown('<div class="home-btn-container">', unsafe_allow_html=True)
    col_home, _ = st.columns([1, 6])
    with col_home:
        if st.button("🏠 Back to Home", type="secondary", key="home_ie"): 
            st.session_state.app_mode = "Home"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="main-header-container"><div class="main-title">🔬 Interaction Energy Extractor</div><div class="sub-title">Automated extraction and computation of ΔE from Gaussian log files.</div></div>', unsafe_allow_html=True)

    col_left, col_right = st.columns(2)
    with col_left:
        with st.container(border=True):
            st.markdown('<p class="section-header">1. Primary Reference System</p>', unsafe_allow_html=True)
            m1_file = st.file_uploader("Upload Monomer 1 (e.g., Amino Acid)", type=["log", "out"])
    with col_right:
        with st.container(border=True):
            st.markdown('<p class="section-header">2. Secondary Systems</p>', unsafe_allow_html=True)
            num_systems = st.number_input("Number of secondary systems:", min_value=1, max_value=20, value=1)
            bsse_choice = st.radio("BSSE Calculated?", ["No", "Yes"], horizontal=True)
            is_bsse = (bsse_choice == "Yes")

    st.write("") 
    systems_data = []
    for i in range(num_systems):
        with st.container(border=True):
            st.markdown(f"**System Type {i+1}**")
            col1, col2, col3 = st.columns([2, 1, 3])
            with col1: m2 = st.file_uploader(f"Upload Monomer 2", type=["log", "out"], key=f"ie_m2_{i}")
            with col2: n_m2 = st.number_input(f"Quantity (n)", min_value=1, max_value=50, value=1, key=f"ie_qty_{i}")
            with col3: comps = st.file_uploader(f"Upload corresponding Complexes", type=["log", "out"], accept_multiple_files=True, key=f"ie_comps_{i}")
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
                if st.button("Extract Data & Tabulate", type="primary", use_container_width=True): do_calc = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Extract Data & Tabulate", type="primary", use_container_width=True): do_calc = True
            with col_down:
                st.download_button("Download Data (.xlsx)", data=st.session_state.excel_data_ie, file_name=st.session_state.download_name_ie, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_calc:
        if not m1_file:
            st.error("Please upload Monomer 1.")
        else:
            with st.spinner("Processing..."):
                e_m1 = extract_energy(m1_file, is_bsse=False)
                if e_m1 is None: 
                    st.error(f"HF= energy is not found in {m1_file.name}.")
                    st.stop()
                    
                results = []
                is_first_in_sheet = True
                
                for idx, sys in enumerate(systems_data):
                    if not sys["m2"] or not sys["complexes"]: continue
                    
                    e_m2_single = extract_energy(sys["m2"], is_bsse=False) 
                    if e_m2_single is None: 
                        st.error(f"HF= energy is not found in {sys['m2'].name}. Skipping system.")
                        continue
                        
                    e_m2_total = e_m2_single * sys["multiplier"]
                    group_results = []
                    
                    for d_file in sys["complexes"]:
                        e_dimer = extract_energy(d_file, is_bsse=is_bsse)
                        if e_dimer is not None:
                            delta_au = e_dimer - (e_m1 + e_m2_total)
                            group_results.append({
                                "Complex File": d_file.name, "E Complex (a.u.)": e_dimer, "Monomer 1 File": m1_file.name,
                                "E Monomer 1 (a.u.)": e_m1, "Monomer 2 File": sys["m2"].name, "Qty": int(sys["multiplier"]),
                                "E Monomer 2 (single a.u.)": e_m2_single, "ΔE (a.u.)": delta_au, "ΔE (kcal/mol)": delta_au * HARTREE_TO_KCAL
                            })
                        else:
                            if is_bsse:
                                st.warning(f"Counterpoise corrected energy is not found in {d_file.name}.")
                            else:
                                st.warning(f"HF= energy is not found in {d_file.name}.")
                            
                    if group_results:
                        min_kcal = min(r["ΔE (kcal/mol)"] for r in group_results)
                        for i, r in enumerate(group_results):
                            r["Is_Best"] = (r["ΔE (kcal/mol)"] == min_kcal)
                            r["Is_First_In_Group"] = (i == 0)
                            r["Is_First_In_Sheet"] = is_first_in_sheet
                            results.append(r)
                            is_first_in_sheet = False

                if not results:
                    st.error("Calculations aborted. No valid interaction energies could be extracted from the uploaded files.")
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
        st.markdown('<p class="section-header" style="text-align: center; border: none;">Tabulated Results</p>', unsafe_allow_html=True)
        
        display_df = pd.DataFrame(st.session_state.results_ie).drop(columns=["Is_Best", "Is_First_In_Sheet", "Is_First_In_Group"])
        
        cols_to_float = ["E Complex (a.u.)", "E Monomer 1 (a.u.)", "E Monomer 2 (single a.u.)", "ΔE (a.u.)", "ΔE (kcal/mol)"]
        for col in cols_to_float:
            if col in display_df.columns:
                display_df[col] = pd.to_numeric(display_df[col], errors='coerce')
                
        def highlight_best(row):
            if st.session_state.results_ie[row.name]["Is_Best"]: return ['background-color: #FEF08A; color: #1F2937; font-weight: bold'] * len(row)
            return [''] * len(row)
            
        styled_df = display_df.style.apply(highlight_best, axis=1).format(precision=6)
        st.dataframe(styled_df, use_container_width=True)

# ==========================================
# 5. APP MODE 2: LED ANALYSIS
# ==========================================
elif st.session_state.app_mode == "LED Analysis":
    st.markdown('<div class="home-btn-container">', unsafe_allow_html=True)
    col_home, _ = st.columns([1, 6])
    with col_home:
        if st.button("🏠 Back to Home", type="secondary", key="home_led"): 
            st.session_state.app_mode = "Home"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
            
    st.markdown('<div class="main-header-container"><div class="main-title">📊 Local Energy Decomposition (LED) Parser</div><div class="sub-title">Parse Orca DLPNO-CCSD(T) output to decompose interaction energies.</div></div>', unsafe_allow_html=True)
    
    num_led_systems = st.number_input("How many systems to evaluate?", min_value=1, max_value=10, value=1)
    
    led_inputs = []
    global_m1 = None
    
    for i in range(num_led_systems):
        with st.container(border=True):
            header_placeholder = st.empty()
            d_file = st.file_uploader(f"Upload Dimer/Complex File", type=["log", "out"], key=f"led_dimer_{i}")
            
            if d_file:
                sys_name = os.path.splitext(d_file.name)[0]
                header_placeholder.markdown(f'<p class="section-header">System {i+1} Configuration: <span style="color:#2563EB;">{sys_name}</span></p>', unsafe_allow_html=True)
            else:
                sys_name = f"System_{i+1}"
                header_placeholder.markdown(f'<p class="section-header">System {i+1} Configuration</p>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                if i == 0:
                    m1_file = st.file_uploader(f"Upload Monomer 1", type=["log", "out"], key=f"led_m1_{i}")
                    global_m1 = m1_file
                else:
                    use_global_m1 = st.checkbox(f"Use same Monomer 1 as System 1", value=True, key=f"led_use_m1_{i}")
                    if use_global_m1:
                        m1_file = global_m1
                        st.info("Using Monomer 1 from System 1")
                    else:
                        m1_file = st.file_uploader(f"Upload Monomer 1 for this system", type=["log", "out"], key=f"led_m1_{i}")
                
                f1_file = st.file_uploader(f"Upload Fragment 1 (Dimer geometry)", type=["log", "out"], key=f"led_f1_{i}")
                
            with col2:
                m2_file = st.file_uploader(f"Upload Monomer 2", type=["log", "out"], key=f"led_m2_{i}")
                f2_file = st.file_uploader(f"Upload Fragment 2 (Dimer geometry)", type=["log", "out"], key=f"led_f2_{i}")
            
            led_inputs.append({"System_Name": sys_name, "m1": m1_file, "m2": m2_file, "f1": f1_file, "f2": f2_file, "d": d_file})

    st.divider()
    
    if 'calc_done_led' not in st.session_state:
        st.session_state.calc_done_led = False
        st.session_state.excel_data_led = None

    btn_placeholder = st.empty()
    do_led_calc = False

    with btn_placeholder.container():
        if not st.session_state.calc_done_led:
            col_s1, col_btn, col_s2 = st.columns([1, 1, 1])
            with col_btn:
                if st.button("Parse LED Data", type="primary", use_container_width=True): do_led_calc = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Parse LED Data", type="primary", use_container_width=True): do_led_calc = True
            with col_down:
                st.download_button("Download LED Report (.xlsx)", data=st.session_state.excel_data_led, file_name="LED_Batch_Results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_led_calc:
        with st.spinner("Parsing Orca outputs..."):
            base_energy = ["FINAL SINGLE POINT ENERGY"]
            triples_term = ["Triples Correction (T)"] 
            fragment_terms = ["E(CORR)(corrected)"] 
            monomer_terms = ["E(0)"] 
            dimer_terms = [
                "Intra REF. energy", "Electrostatics (REF.)", "Exchange (REF.)",
                "Non dispersion (strong pairs)", "Non dispersion (weak pairs)",
                "Dispersion (strong pairs)", "Dispersion (weak pairs)"
            ]
            
            unique_extracted_data = {}
            all_calculations = []
            
            for sys in led_inputs:
                d_file = sys["d"]
                f1_file = sys["f1"]
                f2_file = sys["f2"]
                m1_file = sys["m1"]
                m2_file = sys["m2"]
                
                if not all([d_file, f1_file, f2_file, m1_file, m2_file]):
                    st.error(f"Missing files for System '{sys['System_Name']}'. Skipping.")
                    continue
                
                file_mapping = [
                    (d_file, "Dimer", base_energy + triples_term + dimer_terms),
                    (f1_file, "Fragment 1", base_energy + fragment_terms),
                    (f2_file, "Fragment 2", base_energy + fragment_terms),
                    (m1_file, "True Monomer 1", base_energy + triples_term + monomer_terms),
                    (m2_file, "True Monomer 2", base_energy + triples_term + monomer_terms)
                ]
                
                sys_dict_refs = {}
                for fileobj, ftype, terms in file_mapping:
                    if fileobj.name not in unique_extracted_data:
                        data = {"Filename": fileobj.name, "Type": ftype}
                        extracted = extract_led_values(fileobj, terms)
                        
                        # LED Error validation for ALL expected terms dynamically
                        for term in terms:
                            if term == "Intra REF. energy":
                                if "Intra REF. energy of Fragment 1" not in extracted or "Intra REF. energy of Fragment 2" not in extracted:
                                    st.warning(f"'{term}' is missing or incomplete in {fileobj.name}.")
                            else:
                                if term not in extracted:
                                    st.warning(f"'{term}' is not found in {fileobj.name}.")
                                    
                        data.update(extracted)
                        unique_extracted_data[fileobj.name] = data
                    sys_dict_refs[ftype] = unique_extracted_data[fileobj.name]
                
                calc = {"System": sys["System_Name"]}
                res_d = sys_dict_refs["Dimer"]
                res_f1 = sys_dict_refs["Fragment 1"]
                res_f2 = sys_dict_refs["Fragment 2"]
                res_m1 = sys_dict_refs["True Monomer 1"]
                res_m2 = sys_dict_refs["True Monomer 2"]

                prep_energy = el_prep = electrostatics = exchange = delta_e_no_disp = e_disp = delta_t_int = Decimal('0.0')

                try:
                    prep_energy = (res_f1["FINAL SINGLE POINT ENERGY"] + res_f2["FINAL SINGLE POINT ENERGY"]) - (res_m1["FINAL SINGLE POINT ENERGY"] + res_m2["FINAL SINGLE POINT ENERGY"])
                    calc["Delta E geo-prep"] = prep_energy * HARTREE_TO_KCAL
                except KeyError: pass
                try:
                    delta_t_int = res_d["Triples Correction (T)"] - res_m1["Triples Correction (T)"] - res_m2["Triples Correction (T)"]
                    calc["Delta E C-(T) int"] = delta_t_int * HARTREE_TO_KCAL
                except KeyError: pass
                try:
                    el_prep = (res_d["Intra REF. energy of Fragment 1"] - res_m1["E(0)"]) + (res_d["Intra REF. energy of Fragment 2"] - res_m2["E(0)"])
                    calc["Delta E ref el-prep"] = el_prep * HARTREE_TO_KCAL
                except KeyError: pass
                try:
                    electrostatics = res_d["Electrostatics (REF.)"]
                    calc["Electrostatics (REF.)"] = electrostatics * HARTREE_TO_KCAL
                except KeyError: pass
                try:
                    exchange = res_d["Exchange (REF.)"]
                    calc["Exchange (REF.)"] = exchange * HARTREE_TO_KCAL
                except KeyError: pass
                try:
                    delta_e_no_disp = (res_d["Non dispersion (strong pairs)"] + res_d["Non dispersion (weak pairs)"]) - res_f1["E(CORR)(corrected)"] - res_f2["E(CORR)(corrected)"]
                    calc["Delta E C-CCSD no-disp"] = delta_e_no_disp * HARTREE_TO_KCAL
                except KeyError: pass
                try:
                    e_disp = res_d["Dispersion (strong pairs)"] + res_d["Dispersion (weak pairs)"]
                    calc["E C-CCSD disp"] = e_disp * HARTREE_TO_KCAL
                except KeyError: pass
                
                calc["Total Delta E"] = (prep_energy + el_prep + electrostatics + exchange + delta_e_no_disp + e_disp + delta_t_int) * HARTREE_TO_KCAL
                all_calculations.append(calc)

            if not all_calculations:
                st.error("No valid LED data could be parsed. Check warnings above.")
                st.session_state.calc_done_led = False
                st.stop()

            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            
            # RAW VALUES TAB
            ws_values = workbook.add_worksheet('Extracted Values')
            all_val_keys = set()
            for data in unique_extracted_data.values(): all_val_keys.update(data.keys())
            val_columns = ["Filename", "Type"] + sorted([k for k in all_val_keys if k not in ["Filename", "Type"]])
            ws_values.set_row(0, 20)
            for col_num, header in enumerate(val_columns): ws_values.write(0, col_num, header, header_format)
            ws_values.set_column('A:B', 25)
            ws_values.set_column(2, len(val_columns)-1, 20)
            for row_num, data in enumerate(unique_extracted_data.values(), start=1):
                ws_values.set_row(row_num, 20)
                for col_num, col_name in enumerate(val_columns):
                    val = data.get(col_name, "")
                    cell_fmt = left_format if col_num < 2 else center_format
                    if isinstance(val, Decimal): ws_values.write_number(row_num, col_num, float(val), cell_fmt)
                    else: ws_values.write(row_num, col_num, str(val), cell_fmt)

            # CALCULATIONS TAB
            ws_calcs = workbook.add_worksheet('Tabulated Data (kcal per mol)')
            calc_columns = list(all_calculations[0].keys())
            sup = workbook.add_format({'bold': True, 'font_script': 1, 'font_name': 'Times New Roman', 'font_size': 12})
            sub = workbook.add_format({'bold': True, 'font_script': 2, 'font_name': 'Times New Roman', 'font_size': 12})
            bold = workbook.add_format({'bold': True, 'font_name': 'Times New Roman', 'font_size': 12})
            header_base = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            
            rich_headers = {
                "System": [bold, "System"],
                "Delta E geo-prep": [bold, "ΔE", sub, "geo-prep"],
                "Delta E C-(T) int": [bold, "ΔE", sup, "C-(T)", sub, "int"],
                "Delta E ref el-prep": [bold, "ΔE", sup, "ref", sub, "el-prep"],
                "Electrostatics (REF.)": [bold, "E", sup, "ref", sub, "elstat"],
                "Exchange (REF.)": [bold, "E", sup, "ref", sub, "exch"],
                "Delta E C-CCSD no-disp": [bold, "ΔE", sup, "C-CCSD", sub, "no-disp"],
                "E C-CCSD disp": [bold, "E", sup, "C-CCSD", sub, "disp"],
                "Total Delta E": [bold, "ΔE"]
            }
            ws_calcs.set_row(0, 20)
            for col_num, header_key in enumerate(calc_columns):
                if header_key in rich_headers:
                    parts = rich_headers[header_key]
                    if len(parts) == 2: ws_calcs.write(0, col_num, parts[1], header_format)
                    else: ws_calcs.write_rich_string(0, col_num, *parts, header_base)
                else:
                    ws_calcs.write(0, col_num, header_key, header_format)
            ws_calcs.set_column(0, 0, 25)
            if len(calc_columns) > 1: ws_calcs.set_column(1, len(calc_columns)-1, 20)
            for row_num, calc_data in enumerate(all_calculations, start=1):
                ws_calcs.set_row(row_num, 20)
                for col_num, col_name in enumerate(calc_columns):
                    val = calc_data.get(col_name, "")
                    cell_fmt = left_format if col_num == 0 else center_format
                    if isinstance(val, Decimal): ws_calcs.write_number(row_num, col_num, float(val), cell_fmt)
                    else: ws_calcs.write(row_num, col_num, str(val), cell_fmt)

            workbook.close()
            st.session_state.excel_data_led = output.getvalue()
            st.session_state.calc_done_led = True
            st.success("LED Data Parsing Complete! Click Download to view the report.")
            st.rerun()

# ==========================================
# 6. APP MODE 3: NBO E(2) EXTRACTOR
# ==========================================
elif st.session_state.app_mode == "NBO E(2) Analysis":
    st.markdown('<div class="home-btn-container">', unsafe_allow_html=True)
    col_home, _ = st.columns([1, 6])
    with col_home:
        if st.button("🏠 Back to Home", type="secondary", key="home_nbo"): 
            st.session_state.app_mode = "Home"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="main-header-container"><div class="main-title">⚛️ NBO E(2) Extractor</div><div class="sub-title">Parse and tabulate Natural Bond Orbital donor-acceptor interactions from log files.</div></div>', unsafe_allow_html=True)
    with st.container(border=True):
        col1, col2 = st.columns([1, 2])
        with col1: top_n = st.number_input("How many highest E(2) values needed per file?", min_value=1, max_value=50, value=1)
        with col2: allowed_sections = st.multiselect("Interaction types to consider (for highest values):", ["within unit 1", "from unit 1 to unit 2", "from unit 2 to unit 1", "within unit 2"], default=["from unit 1 to unit 2", "from unit 2 to unit 1"])
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
                if st.button("Extract NBO Data", type="primary", use_container_width=True): do_nbo_run = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Extract NBO Data", type="primary", use_container_width=True): do_nbo_run = True
            with col_down:
                st.download_button("Download NBO Report (.xlsx)", data=st.session_state.excel_data_nbo, file_name="NBO_Batch_Results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_nbo_run:
        if not nbo_files:
            st.error("Please upload NBO files.")
        else:
            file_data_dict = {}
            all_top_n_data = []
            for uploaded_file in nbo_files:
                uploaded_file.seek(0)
                lines = uploaded_file.read().decode('utf-8').splitlines()
                basename = os.path.splitext(uploaded_file.name)[0]
                file_extracted_data = []
                in_target_section = False
                found_perturbation = False
                
                for line in lines:
                    if "Second Order Perturbation Theory Analysis of Fock Matrix in NBO Basis" in line:
                        in_target_section = True
                        found_perturbation = True
                        continue
                    if in_target_section:
                        if "Natural Bond Orbitals (Summary):" in line: break
                        line_stripped = line.strip()
                        if line_stripped.startswith("within unit") or line_stripped.startswith("from unit"):
                            current_section = re.sub(r'\s+', ' ', line_stripped); continue
                        if '/' in line:
                            parts = line.split('/')
                            donor_raw = parts[0].strip()
                            right_tokens = parts[1].strip().split()
                            if len(right_tokens) >= 4:
                                donor_fmt = format_nbo(donor_raw)
                                acceptor_fmt = format_nbo(" ".join(right_tokens[1:-3]))
                                if not any(term in donor_fmt for term in ["RY*", "CR"]) and not any(term in acceptor_fmt for term in ["RY*", "CR"]):
                                    file_extracted_data.append([basename, current_section, donor_fmt, acceptor_fmt, float(right_tokens[-3]), float(right_tokens[-2]), float(right_tokens[-1])])
                
                # NBO Error validation
                if not found_perturbation:
                    st.warning(f"Perturbation table is not found in {uploaded_file.name}.")
                elif not file_extracted_data:
                    st.warning(f"No valid donor-acceptor interactions extracted from {uploaded_file.name}.")
                else:
                    file_data_dict[basename] = file_extracted_data
                    filtered = [r for r in file_extracted_data if r[1] in allowed_sections]
                    all_top_n_data.extend(sorted(filtered, key=lambda r: r[4], reverse=True)[:top_n])
                    
            if not all_top_n_data:
                st.error("NBO extraction failed. No valid data found in the provided files.")
                st.stop()
                
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
            workbook.close()
            st.session_state.all_top_n = all_top_n_data
            st.session_state.excel_data_nbo = output.getvalue()
            st.session_state.calc_done_nbo = True
            st.rerun()
    if st.session_state.calc_done_nbo:
        st.write("") 
        if st.session_state.all_top_n:
            st.dataframe(pd.DataFrame(st.session_state.all_top_n, columns=["System", "Interaction Section", "Donor (i)", "Acceptor (j)", "E(2) kcal/mol", "E(j)-E(i) a.u.", "F(i,j) a.u."]), use_container_width=True)

# ==========================================
# 7. APP MODE 4: QTAIM PROPERTIES TABULATOR
# ==========================================
elif st.session_state.app_mode == "QTAIM Analysis":
    st.markdown('<div class="home-btn-container">', unsafe_allow_html=True)
    col_home, _ = st.columns([1, 6])
    with col_home:
        if st.button("🏠 Back to Home", type="secondary", key="home_qtaim"): 
            st.session_state.app_mode = "Home"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="main-header-container"><div class="main-title">⭕ QTAIM Properties Tabulator</div><div class="sub-title">Parse Multiwfn CPprop.txt files to extract density properties at Bond Critical Points.</div></div>', unsafe_allow_html=True)
    with st.container(border=True):
        num_qtaim_files = st.number_input("How many CPprop output files are you analyzing?", min_value=1, max_value=20, value=1)
    st.write("")
    qtaim_inputs = []
    for i in range(num_qtaim_files):
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1: q_file = st.file_uploader(f"Upload CPprop.txt {i+1}", type=["txt"], key=f"q_file_{i}")
            with col2: cp_string = st.text_input("Target Critical Points (comma-separated, e.g. 15, 23)", key=f"q_cps_{i}")
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
                if st.button("Parse QTAIM Data", type="primary", use_container_width=True): do_qtaim_run = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Parse QTAIM Data", type="primary", use_container_width=True): do_qtaim_run = True
            with col_down:
                st.download_button("Download QTAIM Report (.xlsx)", data=st.session_state.excel_data_qtaim, file_name="QTAIM_Results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_qtaim_run:
        search_terms = {"Density of all electrons": "Density of all electrons:", "Laplacian of electron density": "Laplacian of electron density:", "Potential energy density V(r)": "Potential energy density V(r):", "Lagrangian kinetic energy G(r)": "Lagrangian kinetic energy G(r):", "Energy density E(r) or H(r)": "Energy density E(r) or H(r):"}
        all_extracted_data = []
        for item in qtaim_inputs:
            uploaded_file = item["file"]
            if not uploaded_file: continue
            try: target_cps = [int(x.strip()) for x in item["cp_str"].split(',') if x.strip().isdigit()]
            except: continue
            if not target_cps: continue
            
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
                    if current_cp in target_cps and cp_data: extracted_data.append(cp_data)
                    match = re.search(r'CP\s+(\d+)', line)
                    if match:
                        current_cp = int(match.group(1))
                        cp_data = {"System": system_name, "CP Number": current_cp}
                    else: current_cp = None
                    continue
                if current_cp in target_cps:
                    for key, term in search_terms.items():
                        if line.lstrip().startswith(term):
                            try: cp_data[key] = float(line.split(term)[1].strip())
                            except: cp_data[key] = None
            if current_cp in target_cps and cp_data: extracted_data.append(cp_data)
            
            # QTAIM Error validation
            if not extracted_data:
                st.warning(f"Target Critical Points ({item['cp_str']}) and their properties are not found in {uploaded_file.name}.")
            else:
                for data in extracted_data:
                    density = data.get("Density of all electrons")
                    data["Binding Energy"] = (-223.08 * density) + 0.7423 if density is not None else None
                extracted_data.sort(key=lambda x: x["CP Number"])
                all_extracted_data.extend(extracted_data)

        if not all_extracted_data:
            st.error("QTAIM extraction failed. Check warnings above for missing data.")
            st.stop()

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
        worksheet.write(0, 0, "System", header_format); worksheet.write(0, 1, "CP", header_format); worksheet.write(0, 2, "ρ(r)", header_format)
        worksheet.write_rich_string(0, 3, bold_fmt, "Δ", sup, "2", bold_fmt, "ρ", header_base)
        worksheet.write_rich_string(0, 4, bold_fmt, "V(r", sub, "bcp", bold_fmt, ")", header_base)
        worksheet.write(0, 5, "G(rcp)", header_format); worksheet.write(0, 6, "H(r)", header_format); worksheet.write(0, 7, "B. E (kcal/Mol)", header_format)
        worksheet.set_column('A:A', 22); worksheet.set_column('B:B', 8); worksheet.set_column('C:G', 14); worksheet.set_column('H:H', 18)
        
        if all_extracted_data:
            keys_in_order = ["Density of all electrons", "Laplacian of electron density", "Potential energy density V(r)", "Lagrangian kinetic energy G(r)", "Energy density E(r) or H(r)"]
            start_row = 1; current_system = all_extracted_data[0]["System"]
            for row_num, data in enumerate(all_extracted_data, start=1):
                worksheet.set_row(row_num, 20)
                if data["System"] != current_system:
                    if row_num - 1 > start_row: worksheet.merge_range(start_row, 0, row_num - 1, 0, current_system, left_format)
                    else: worksheet.write(start_row, 0, current_system, left_format)
                    start_row = row_num; current_system = data["System"]
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
