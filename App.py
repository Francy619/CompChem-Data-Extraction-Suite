import streamlit as st
import os
import re
import xlsxwriter
import io
from decimal import Decimal
import pandas as pd
import streamlit.components.v1 as components

# --- IMPORTS FOR 3D RENDERING ---
try:
    import py3Dmol
    from stmol import showmol
    HAS_3D = True
except ImportError:
    HAS_3D = False

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
        
        [data-testid="stAppViewContainer"] { overflow-y: hidden !important; }
        
        /* Override secondary buttons ONLY on this page to make them giant blue blocks */
        button[kind="secondary"] { 
            background-color: #1E3A8A !important; 
            color: #FFFFFF !important; 
            border: none !important; 
            border-radius: 8px !important; 
            font-weight: 800 !important; 
            font-size: 2.4rem !important; 
            min-height: 100px !important;
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
        st.write("")
        if st.button("📏 Vibrational Analysis", use_container_width=True, type="secondary"): 
            st.session_state.app_mode = "Structural Analysis"
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
        st.session_state.warnings_ie = []

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
                st.session_state.warnings_ie = [] 
                
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
                                st.session_state.warnings_ie.append(f"Counterpoise corrected energy is not found in {d_file.name}.")
                            else:
                                st.session_state.warnings_ie.append(f"HF= energy is not found in {d_file.name}.")
                            
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
        for w in st.session_state.warnings_ie:
            st.warning(w)
            
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
        st.session_state.warnings_led = []

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
            st.session_state.warnings_led = []
            
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
                        
                        for term in terms:
                            if term == "Intra REF. energy":
                                if "Intra REF. energy of Fragment 1" not in extracted or "Intra REF. energy of Fragment 2" not in extracted:
                                    st.session_state.warnings_led.append(f"'{term}' is missing or incomplete in {fileobj.name}.")
                            else:
                                if term not in extracted:
                                    st.session_state.warnings_led.append(f"'{term}' is not found in {fileobj.name}.")
                                    
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
            st.rerun()
            
    if st.session_state.calc_done_led:
        for w in st.session_state.warnings_led:
            st.warning(w)

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
        st.session_state.warnings_nbo = []

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
            st.session_state.warnings_nbo = [] 
            
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
                
                if not found_perturbation:
                    st.session_state.warnings_nbo.append(f"Perturbation table is not found in {uploaded_file.name}.")
                elif not file_extracted_data:
                    st.session_state.warnings_nbo.append(f"No valid donor-acceptor interactions extracted from {uploaded_file.name}.")
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
        for w in st.session_state.warnings_nbo:
            st.warning(w)
            
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
        st.session_state.warnings_qtaim = []

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
        st.session_state.warnings_qtaim = [] 
        
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
            
            if not extracted_data:
                st.session_state.warnings_qtaim.append(f"Target Critical Points ({item['cp_str']}) and their properties are not found in {uploaded_file.name}.")
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

    if st.session_state.calc_done_qtaim:
        for w in st.session_state.warnings_qtaim:
            st.warning(w)

# ==========================================
# 8. APP MODE 5: STRUCTURAL & VIBRATIONAL ANALYSIS
# ==========================================
elif st.session_state.app_mode == "Structural Analysis":
    import math
    st.markdown('<div class="home-btn-container">', unsafe_allow_html=True)
    col_home, _ = st.columns([1, 6])
    with col_home:
        if st.button("🏠 Back to Home", type="secondary", key="home_struct"): 
            st.session_state.app_mode = "Home"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="main-header-container"><div class="main-title">📏 Structural & Vibrational Data Analysis</div><div class="sub-title">Track bond lengths and stretching frequencies from Gaussian log files.</div></div>', unsafe_allow_html=True)
    
    # --- HELPER FUNCTIONS FOR THIS MODULE ---
    ATOMIC_SYMBOLS = {
        1: 'H', 2: 'He', 3: 'Li', 4: 'Be', 5: 'B', 6: 'C', 7: 'N', 8: 'O', 9: 'F', 10: 'Ne',
        11: 'Na', 12: 'Mg', 13: 'Al', 14: 'Si', 15: 'P', 16: 'S', 17: 'Cl', 18: 'Ar',
        19: 'K', 20: 'Ca', 24: 'Cr', 25: 'Mn', 26: 'Fe', 27: 'Co', 28: 'Ni', 29: 'Cu', 30: 'Zn',
        35: 'Br', 53: 'I'
    }

    def get_optimized_coords(content):
        coords = {}
        lines = content.splitlines()
        for i in range(len(lines)-1, -1, -1):
            if "Standard orientation:" in lines[i] or "Input orientation:" in lines[i]:
                start_idx = i + 5
                for j in range(start_idx, len(lines)):
                    if "---------------------------------------------------------------------" in lines[j]: break
                    parts = lines[j].split()
                    if len(parts) == 6: coords[int(parts[0])] = (float(parts[3]), float(parts[4]), float(parts[5]))
                break
        return coords

    def show_3d_molecule(content):
        lines = content.splitlines()
        coords = []
        for i in range(len(lines)-1, -1, -1):
            if "Standard orientation:" in lines[i] or "Input orientation:" in lines[i]:
                start_idx = i + 5
                for j in range(start_idx, len(lines)):
                    if "---------------------------------------------------------------------" in lines[j]: break
                    parts = lines[j].split()
                    if len(parts) == 6:
                        atom_num = int(parts[1])
                        sym = ATOMIC_SYMBOLS.get(atom_num, 'X')
                        coords.append({"idx": parts[0], "sym": sym, "x": parts[3], "y": parts[4], "z": parts[5]})
                break
        
        if not coords: 
            st.warning("Could not extract molecular coordinates for 3D viewer.")
            return
        
        xyz_str = f"{len(coords)}\n\n"
        for c in coords:
            xyz_str += f"{c['sym']} {c['x']} {c['y']} {c['z']}\n"
            
        view = py3Dmol.view(width=450, height=400)
        view.addModel(xyz_str, 'xyz')
        view.setStyle({'stick':{'radius': 0.15}, 'sphere':{'scale': 0.3}})
        
        for c in coords:
            view.addLabel(f"{c['sym']}{c['idx']}", {
                "position": {"x": float(c['x']), "y": float(c['y']), "z": float(c['z'])},
                "fontColor": "black",
                "backgroundColor": "white",
                "backgroundOpacity": 0.8,
                "fontsize": 14
            })
            
        view.zoomTo()
        showmol(view, height=400, width=450)

    def calculate_distance(coord1, coord2):
        return math.sqrt(sum((c1 - c2) ** 2 for c1, c2 in zip(coord1, coord2)))

    def find_stretching_frequency(content, atomA, atomB, coords):
        if not coords or atomA not in coords or atomB not in coords: return None
        cA, cB = coords[atomA], coords[atomB]
        dist = calculate_distance(cA, cB)
        bond_vector = [(cB[0]-cA[0])/dist, (cB[1]-cA[1])/dist, (cB[2]-cA[2])/dist]
        max_stretch, best_freq = -1.0, None
        current_freqs = []
        current_disps = {}
        lines = content.splitlines()
        for line in lines:
            if "Frequencies --" in line:
                current_freqs = [float(x) for x in line.split()[2:]]
                current_disps = {} 
            elif current_freqs and len(line.split()) >= 11 and "Atom" not in line:
                parts = line.split()
                if not parts[0].isdigit(): continue
                try:
                    atom_idx = int(parts[0])
                    if atom_idx in (atomA, atomB):
                        current_disps[atom_idx] = [(float(parts[2+m*3]), float(parts[3+m*3]), float(parts[4+m*3])) for m in range(len(current_freqs))]
                        if atomA in current_disps and atomB in current_disps:
                            disp_A, disp_B = current_disps[atomA], current_disps[atomB]
                            for m in range(len(current_freqs)):
                                rel = (disp_B[m][0]-disp_A[m][0], disp_B[m][1]-disp_A[m][1], disp_B[m][2]-disp_A[m][2])
                                stretch = abs(rel[0]*bond_vector[0] + rel[1]*bond_vector[1] + rel[2]*bond_vector[2])
                                if stretch > max_stretch:
                                    max_stretch, best_freq = stretch, current_freqs[m]
                except ValueError: continue
        return best_freq

    # --- UI LAYOUT ---
    st.markdown('<p class="section-header">⚙️ General Settings</p>', unsafe_allow_html=True)
    with st.container(border=True):
        col_scale1, col_scale2 = st.columns(2)
        with col_scale1:
            need_scaling = st.radio("Apply Frequency Scaling?", ["No", "Yes"], index=1, horizontal=True)
        with col_scale2:
            scaling_factor = None
            if need_scaling == "Yes":
                scaling_factor = st.number_input("Enter Scaling Factor:", value=0.9594, format="%.4f")
    
    st.write("")
    st.markdown('<p class="section-header">1. Primary System (Monomer 1)</p>', unsafe_allow_html=True)
    with st.container(border=True):
        col_m1, _ = st.columns([1, 1])
        with col_m1:
            m1_file = st.file_uploader("Upload Monomer 1 .log file (e.g., Amino Acid)", type=["log", "out"], key="str_m1")
            
            # --- 3D MOLECULE VIEWER M1 ---
            if m1_file:
                if HAS_3D:
                    st.markdown("**👁️ Monomer 1 Structure & Atoms**")
                    show_3d_molecule(m1_file.getvalue().decode('utf-8'))
                else:
                    st.error("🚨 Missing Required Packages! To see the 3D molecule viewer, you must add `stmol` and `py3Dmol` to your `requirements.txt` file (if on Streamlit Cloud) or run `pip install stmol py3Dmol` on your computer.")
            
            m1_bonds = st.number_input("Number of target bonds in Monomer 1:", min_value=0, max_value=5, value=2)
            m1_bond_data = []
            for b in range(m1_bonds):
                c_name, c_atoms = st.columns([2, 2])
                with c_name: b_name = st.text_input(f"Bond {b+1} Name (e.g., O-H, C=O)", key=f"m1_bname_{b}")
                with c_atoms: b_atoms = st.text_input(f"Atom Numbers (e.g., 1, 17)", key=f"m1_batoms_{b}")
                m1_bond_data.append({"name": b_name, "atoms_str": b_atoms})

    st.write("")
    st.markdown('<p class="section-header">2. Secondary Systems & Complexes</p>', unsafe_allow_html=True)
    num_complexes = st.number_input("Number of secondary systems (e.g., different solvents):", min_value=1, max_value=20, value=1)
    
    systems_input = []
    for i in range(num_complexes):
        with st.container(border=True):
            st.markdown(f"**System {i+1}**")
            c_m2, c_comp = st.columns([1, 1])
            with c_m2:
                m2_file = st.file_uploader(f"Upload Monomer 2 .log file (Solvent {i+1})", type=["log", "out"], key=f"str_m2_{i}")
                
                # --- 3D MOLECULE VIEWER M2 ---
                if m2_file:
                    if HAS_3D:
                        st.markdown(f"**👁️ Monomer 2 Structure & Atoms**")
                        show_3d_molecule(m2_file.getvalue().decode('utf-8'))

                m2_bonds = st.number_input(f"Number of target bonds in Monomer 2 (System {i+1}):", min_value=0, max_value=5, value=1, key=f"m2_b_cnt_{i}")
                m2_bond_data = []
                for b in range(m2_bonds):
                    c_n, c_a = st.columns([2, 2])
                    with c_n: bn = st.text_input(f"Bond {b+1} Name", key=f"m2_n_{i}_{b}")
                    with c_a: ba = st.text_input(f"Atom Numbers", key=f"m2_a_{i}_{b}")
                    m2_bond_data.append({"name": bn, "atoms_str": ba})
                    
            with c_comp:
                c_file = st.file_uploader(f"Upload Complex .log file (M1 + Solvent {i+1})", type=["log", "out"], key=f"str_c_{i}")
                
                # --- 3D MOLECULE VIEWER COMPLEX ---
                if c_file:
                    if HAS_3D:
                        st.markdown(f"**👁️ Complex Structure & Atoms**")
                        show_3d_molecule(c_file.getvalue().decode('utf-8'))
                            
                st.markdown("**Atoms in Complex:**")
                
                m1_d_atoms = []
                if m1_bond_data:
                    st.caption("Monomer 1 Bonds:")
                    for b, bdata in enumerate(m1_bond_data):
                        if bdata["name"]: 
                            d_atm = st.text_input(f"{bdata['name']} atoms:", value=bdata["atoms_str"], key=f"c{i}_m1b_{b}")
                            m1_d_atoms.append({"name": bdata["name"], "atoms_str": d_atm})
                            
                m2_d_atoms = []
                if m2_bond_data:
                    st.caption("Monomer 2 Bonds:")
                    for b, bdata in enumerate(m2_bond_data):
                        if bdata["name"]: 
                            d_atm = st.text_input(f"{bdata['name']} atoms:", value=bdata["atoms_str"], key=f"c{i}_m2b_{b}")
                            m2_d_atoms.append({"name": bdata["name"], "atoms_str": d_atm})
                            
            systems_input.append({"m2_file": m2_file, "m2_bonds": m2_bond_data, "c_file": c_file, "c_m1_bonds": m1_d_atoms, "c_m2_bonds": m2_d_atoms})

    if 'calc_done_struct' not in st.session_state:
        st.session_state.calc_done_struct = False
        st.session_state.excel_data_struct = None
        st.session_state.warnings_struct = []
        st.session_state.extracted_rows_struct = []
        st.session_state.bond_names_struct = []
        
    btn_placeholder = st.empty()
    do_struct_run = False
    with btn_placeholder.container():
        if not st.session_state.calc_done_struct:
            col_s1, col_btn, col_s2 = st.columns([1, 1, 1])
            with col_btn:
                if st.button("Run Vibrational Analysis", type="primary", use_container_width=True): do_struct_run = True
        else:
            col_s1, col_calc, col_down, col_s2 = st.columns([2, 3, 3, 2])
            with col_calc:
                if st.button("Run Vibrational Analysis", type="primary", use_container_width=True): do_struct_run = True
            with col_down:
                st.download_button("Download Analysis (.xlsx)", data=st.session_state.excel_data_struct, file_name="Vibrational_Analysis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="secondary")

    if do_struct_run:
        st.session_state.warnings_struct = [] 
        extracted_rows = []
        
        all_bond_names = []
        for bd in m1_bond_data:
            if bd["name"] and bd["name"] not in all_bond_names: all_bond_names.append(bd["name"])
        for sys in systems_input:
            for bd in sys["m2_bonds"]:
                if bd["name"] and bd["name"] not in all_bond_names: all_bond_names.append(bd["name"])
            
        def process_bond(content, name, atom_str, file_name):
            try: atoms = tuple([int(x.strip()) for x in atom_str.split(",")])
            except: return None, None
            if len(atoms) != 2: return None, None
            coords = get_optimized_coords(content)
            if not coords:
                st.session_state.warnings_struct.append(f"Standard/Input orientation not found in {file_name}.")
                return None, None
            if atoms[0] not in coords or atoms[1] not in coords:
                st.session_state.warnings_struct.append(f"Atoms {atoms} for bond '{name}' not found in coordinates of {file_name}.")
                return None, None
            dist = calculate_distance(coords[atoms[0]], coords[atoms[1]])
            freq = find_stretching_frequency(content, atoms[0], atoms[1], coords)
            if freq is None: st.session_state.warnings_struct.append(f"Stretching frequency for '{name}' not found in {file_name}.")
            return dist, freq

        m1_name = "Monomer_1"
        if m1_file:
            m1_file.seek(0)
            m1_content = m1_file.read().decode('utf-8')
            m1_name = os.path.splitext(m1_file.name)[0]
            row_data = {"System": m1_name}
            for bd in m1_bond_data:
                if bd["name"] and bd["atoms_str"]:
                    dist, freq = process_bond(m1_content, bd["name"], bd["atoms_str"], m1_file.name)
                    row_data[bd["name"]] = {"dist": dist, "freq": freq}
            extracted_rows.append(row_data)

        base_m1 = re.split(r'[_ -]|site', m1_name, flags=re.IGNORECASE)[0]
        
        for idx, comp in enumerate(systems_input):
            c_file = comp["c_file"]
            m2_file = comp["m2_file"]
            
            c_name = os.path.splitext(c_file.name)[0] if c_file else f"Complex_{idx+1}"
            m2_name = os.path.splitext(m2_file.name)[0] if m2_file else f"M2_{idx+1}"
            
            if c_file:
                c_file.seek(0)
                c_content = c_file.read().decode('utf-8')
                c_system_name = f"{base_m1} ({c_name})"
                row_data_comp = {"System": c_system_name}
                for bd in comp["c_m1_bonds"]:
                    if bd["name"] and bd["atoms_str"]:
                        dist, freq = process_bond(c_content, bd["name"], bd["atoms_str"], c_file.name)
                        row_data_comp[bd["name"]] = {"dist": dist, "freq": freq}
                extracted_rows.append(row_data_comp)
            
            if m2_file:
                m2_file.seek(0)
                m2_content = m2_file.read().decode('utf-8')
                row_data_m2 = {"System": m2_name}
                for bd in comp["m2_bonds"]:
                    if bd["name"] and bd["atoms_str"]:
                        dist, freq = process_bond(m2_content, bd["name"], bd["atoms_str"], m2_file.name)
                        row_data_m2[bd["name"]] = {"dist": dist, "freq": freq}
                extracted_rows.append(row_data_m2)
                
            if c_file and comp["c_m2_bonds"]:
                row_data_c_solv = {"System": f"{m2_name} ({c_name})"}
                has_m2_bonds = False
                for bd in comp["c_m2_bonds"]:
                    if bd["name"] and bd["atoms_str"]:
                        has_m2_bonds = True
                        dist, freq = process_bond(c_content, bd["name"], bd["atoms_str"], c_file.name)
                        row_data_c_solv[bd["name"]] = {"dist": dist, "freq": freq}
                if has_m2_bonds: extracted_rows.append(row_data_c_solv)

        if not extracted_rows:
            st.error("No valid data could be extracted.")
            st.stop()
            
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Frequencies')
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
        center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
        left_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'font_name': 'Times New Roman', 'font_size': 12})
        
        columns = ["System"]
        for b_name in all_bond_names:
            columns.extend([f"Bond Length (A) {b_name}", f"{b_name} Freq"])
            if scaling_factor is not None:
                columns.append(f"Scaled {b_name}")
            
        worksheet.set_row(0, 20)
        for col_num, header in enumerate(columns): worksheet.write(0, col_num, header, header_format)
        worksheet.set_column('A:A', 30, left_format)
        if len(columns) > 1: worksheet.set_column(1, len(columns)-1, 20, center_format)
        
        for row_num, row_dict in enumerate(extracted_rows, start=1):
            worksheet.set_row(row_num, 20)
            worksheet.write(row_num, 0, row_dict.get("System", ""), left_format)
            
            col_idx = 1
            for b_name in all_bond_names:
                b_data = row_dict.get(b_name, {})
                dist = b_data.get("dist")
                freq = b_data.get("freq")
                
                worksheet.write(row_num, col_idx, float(dist) if dist else "", center_format)
                worksheet.write(row_num, col_idx+1, float(freq) if freq else "", center_format)
                
                if scaling_factor is not None:
                    worksheet.write(row_num, col_idx+2, float(freq * scaling_factor) if freq else "", center_format)
                    col_idx += 3
                else:
                    col_idx += 2
                
        workbook.close()
        
        st.session_state.extracted_rows_struct = extracted_rows
        st.session_state.bond_names_struct = all_bond_names
        st.session_state.excel_data_struct = output.getvalue()
        st.session_state.applied_scaling_factor = scaling_factor 
        st.session_state.calc_done_struct = True
        st.rerun()

    if st.session_state.calc_done_struct:
        for w in st.session_state.warnings_struct: st.warning(w)
        st.markdown('<p class="section-header" style="text-align: center; border: none;">Tabulated Results</p>', unsafe_allow_html=True)
        
        df_data = []
        for row in st.session_state.extracted_rows_struct:
            flat_row = {"System": row["System"]}
            for b_name in st.session_state.bond_names_struct:
                b_data = row.get(b_name, {})
                flat_row[f"Bond Length (A) {b_name}"] = b_data.get("dist")
                flat_row[f"{b_name} Freq"] = b_data.get("freq")
                if st.session_state.applied_scaling_factor is not None:
                    flat_row[f"Scaled {b_name}"] = b_data.get("freq") * st.session_state.applied_scaling_factor if b_data.get("freq") else None
            df_data.append(flat_row)
            
        display_df = pd.DataFrame(df_data)
        st.dataframe(display_df.style.format(precision=4), use_container_width=True)
