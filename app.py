import streamlit as st
import os
import smtplib
from email.message import EmailMessage
from docxtpl import DocxTemplate
from datetime import datetime
import json

# --- Configuration & Helpers ---
SENDER_EMAIL = st.secrets["sender_email"]
SENDER_PASSWORD = st.secrets["sender_password"]
RECEIVER_EMAIL = "shekerlianlaw@gmail.com""phil@barthattorneys.com"

def send_email_with_docx(docx_path, filename):
    msg = EmailMessage()
    msg['Subject'] = f"New EP Intake: {filename}"
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg.set_content("A new client intake questionnaire has been completed. The generated Word document is attached.")

    with open(docx_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.wordprocessingml.document', filename=filename)

    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)

def cb(condition):
    return "☑" if condition else "☐"

# --- Main App Interface ---
st.set_page_config(page_title="BarthCalderon Intake Form", layout="wide")

# --- Sidebar: Draft Restore System ---
with st.sidebar:
    st.header("💾 Resume Progress")
    st.write("If you saved a draft to your computer, upload the `.json` file here to restore your answers.")
    draft_file = st.file_uploader("Upload Draft File", type=["json"], key="draft_uploader")
    
    if draft_file is not None:
        try:
            saved_data = json.load(draft_file)
            for k, v in saved_data.items():
                if k != "draft_uploader":
                    st.session_state[k] = v
            st.success("Draft loaded! Your answers have been fully restored.")
        except Exception as e:
            st.error("Error loading draft. Please ensure it is a valid file.")

st.title("BarthCalderon Estate Planning Questionnaire")
st.markdown("---")

warning_msg = "⚠️️ **CRITICAL WARNING:** Your privacy is our priority, so your answers are NOT stored on a cloud database. **If you close this browser window or lose internet connection, all information will be lost.** To avoid this, go to Tab 6 at any time and click **'Save Draft to Computer'**."

t_client, t_assets, t_health, t_fiduciaries, t_distrib, t_admin = st.tabs([
    "1. Client & Spouse", "2. Assets & Real Estate", "3. Health & Agents", 
    "4. Guardians & Trustees", "5. Beneficiaries & Gifts", "6. Admin & General"
])

with t_client:
    st.warning(warning_msg)
    st.subheader("Client Information")
    c1, c2, c3 = st.columns(3)
    c_id = c1.text_input("Client name as it appears on your ID", key="c_id")
    c_dob = c2.text_input("Client DOB", key="c_dob")
    c_pob = c3.text_input("Birth State/Country", key="c_pob")
    
    c4, c5 = st.columns(2)
    c_address = c4.text_input("Residence Address", key="c_address")
    c_county = c5.text_input("Residence County", key="c_county")
    c_mailing = st.text_input("Mailing Address (if different)", key="c_mailing")
    
    c6, c7 = st.columns(2)
    c_phone = c6.text_input("Primary Phone", key="c_phone")
    c_email = c7.text_input("Email Address", key="c_email")
    
    c8, c9, c10, c11 = st.columns(4)
    c_gender = c8.radio("Gender", ["Male", "Female"], key="c_gen")
    c_cit = c9.radio("US Citizen?", ["Yes", "No"], key="c_cit")
    c_wid = c10.radio("Widowed/Divorced?", ["No", "Yes"], key="c_wid")
    c_63 = c11.radio("Over 63?", ["No", "Yes"], key="c_63")
    
    c_former_spouses = st.text_input("Name of former spouses & date of divorce or death (if applicable)", key="c_former_spouses")
    
    c12, c13 = st.columns(2)
    c_income = c12.text_input("Approx. Income", key="c_income")
    c_obligations = c13.text_input("Court Ordered Obligations", key="c_obligations")

    st.markdown("---")
    st.subheader("Spouse Information")
    s1, s2, s3 = st.columns(3)
    s_id = s1.text_input("Spouse name as it appears on your ID", key="s_id")
    s_dob = s2.text_input("Spouse DOB", key="s_dob")
    s_pob = s3.text_input("Spouse Birth State/Country", key="s_pob")
    
    s4, s5 = st.columns(2)
    s_address = s4.text_input("Spouse Residence Address", key="s_address")
    s_county = s5.text_input("Spouse Residence County", key="s_county")
    s_mailing = st.text_input("Spouse Mailing Address", key="s_mailing")
    
    s6, s7 = st.columns(2)
    s_phone = s6.text_input("Spouse Phone", key="s_phone")
    s_email = s7.text_input("Spouse Email", key="s_email")
    
    s8, s9, s10, s11 = st.columns(4)
    s_gender = s8.radio("Spouse Gender", ["Male", "Female"], key="s_gen")
    s_cit = s9.radio("Spouse US Citizen?", ["Yes", "No"], key="s_cit")
    s_wid = s10.radio("Spouse Widowed/Divorced?", ["No", "Yes"], key="s_wid")
    s_63 = s11.radio("Spouse Over 63?", ["No", "Yes"], key="s_63")
    
    s_former_spouses = st.text_input("Spouse Former Spouses & Dates", key="s_former_spouses")
    
    s12, s13 = st.columns(2)
    s_income = s12.text_input("Spouse Approx. Income", key="s_income")
    s_obligations = s13.text_input("Spouse Court Obligations", key="s_obligations")

with t_assets:
    st.warning(warning_msg)
    st.subheader("Trust History")
    ta1, ta2 = st.columns(2)
    prior_trust = ta1.radio("Do you already have a trust?", ["No", "Yes"], key="prior_trust")
    t_amend_num = ta2.text_input("Amendment/Restatement #", key="t_amend_num")
    t_desired_name = st.text_input("Desired Name of Trust", key="t_desired_name")
    t_reason = st.text_area("Why do you want a trust?", key="t_reason")
    
    st.markdown("---")
    st.subheader("Financial Accounts")
    ta3, ta4 = st.columns(2)
    li_rev = ta3.radio("Have you had your life insurance reviewed recently?", ["Yes", "No", "Not Applicable"], key="li_rev")
    inv_rev = ta4.radio("Do you want a reccomendation for a financial advisor to review your investments for free?", ["Yes", "No", "Not Applicable"], key="inv_rev")
    
    accounts = []
    for i in range(1, 15):
        with st.expander(f"Account {i}"):
            ac1, ac2, ac3, ac4 = st.columns(4)
            accounts.append({
                'type': ac1.text_input("Type", key=f"a{i}_type"),
                'val': ac2.text_input("Value", key=f"a{i}_val"),
                'inst': ac3.text_input("Institution", key=f"a{i}_inst"),
                'own': ac4.text_input("Owner", key=f"a{i}_own")
            })
            
    st.markdown("---")
    st.subheader("Businesses")
    businesses = []
    for i in range(1, 4):
        with st.expander(f"Business {i}"):
            b1, b2 = st.columns(2)
            name = b1.text_input("Name", key=f"b{i}_name")
            type_b = b2.selectbox("Type", ["Sole Proprietor", "Partnership", "LLC", "Corp", ""], key=f"b{i}_type")
            b3, b4, b5 = st.columns(3)
            date = b3.text_input("Date Started", key=f"b{i}_date")
            ent = b4.text_input("Entity #", key=f"b{i}_ent")
            state = b5.text_input("State", key=f"b{i}_state")
            tax = st.selectbox("Tax Election", ["Disregarded", "Partnership", "S", "C", ""], key=f"b{i}_tax")
            pct = st.text_input("% Owned", key=f"b{i}_pct")
            act = st.text_input("Activity", key=f"b{i}_act")
            val = st.text_input("Value", key=f"b{i}_val")
            suc = st.radio("Succession Plan?", ["Yes", "No"], key=f"b{i}_suc")
            dis = st.radio("Disaster Plan?", ["Yes", "No"], key=f"b{i}_dis")
            businesses.append({'name':name, 'type':type_b, 'date':date, 'ent':ent, 'state':state, 'tax':tax, 'pct':pct, 'act':act, 'val':val, 'suc':suc, 'dis':dis})

    st.markdown("---")
    st.subheader("Money Owed to You")
    owed = []
    for i in range(1, 3):
        with st.expander(f"Loan {i}"):
            o1, o2, o3 = st.columns(3)
            who = o1.text_input("Who Owes", key=f"ow{i}_who")
            amt = o2.text_input("Amount", key=f"ow{i}_amt")
            date = o3.text_input("Date", key=f"ow{i}_date")
            sec = st.text_input("Secured", key=f"ow{i}_sec")
            purp = st.text_input("Purpose", key=f"ow{i}_purp")
            pb = st.text_input("Payback at death?", key=f"ow{i}_pb")
            owed.append({'who':who, 'amt':amt, 'date':date, 'sec':sec, 'purp':purp, 'pb':pb})

    st.markdown("---")
    st.subheader("Real Estate")
    hs_decl = st.radio("Do you have a homestead declaration filed for your primary home?", ["Yes", "No"], key="hs_decl")
    re_pri_addr = st.text_input("Primary Residence Address", key="re_pri_addr")
    re1, re2, re3, re4, re5 = st.columns(5)
    re_pri_val = re1.text_input("Value", key="re_pri_val")
    re_pri_cnty = re2.text_input("County", key="re_pri_cnty")
    re_pri_move = re3.text_input("Move In Date", key="re_pri_move")
    re_pri_eq = re4.text_input("Equity %", key="re_pri_eq")
    re_pri_basis = re5.text_input("Basis", key="re_pri_basis")
    re_pri_gain = st.text_input("Gain", key="re_pri_gain")
    
    re_sec_addr = st.text_input("Secondary Residence Address", key="re_sec_addr")
    re6, re7, re8, re9, re10 = st.columns(5)
    re_sec_val = re6.text_input("Sec Value", key="re_sec_val")
    re_sec_cnty = re7.text_input("Sec County", key="re_sec_cnty")
    re_sec_move = re8.text_input("Sec Move Date", key="re_sec_move")
    re_sec_eq = re9.text_input("Sec Equity %", key="re_sec_eq")
    re_sec_basis = re10.text_input("Sec Basis", key="re_sec_basis")
    re_sec_gain = st.text_input("Sec Gain", key="re_sec_gain")
    
    rentals = []
    for i in range(1, 6):
        with st.expander(f"Rental Property {i}"):
            addr = st.text_input("Address", key=f"r{i}_addr")
            r1, r2, r3, r4, r5 = st.columns(5)
            val = r1.text_input("Value", key=f"r{i}_val")
            cnty = r2.text_input("County", key=f"r{i}_cnty")
            move = r3.text_input("Move Date", key=f"r{i}_move")
            eq = r4.text_input("Equity %", key=f"r{i}_eq")
            basis = r5.text_input("Basis", key=f"r{i}_basis")
            gain = st.text_input("Gain", key=f"r{i}_gain")
            rentals.append({'addr':addr, 'val':val, 'cnty':cnty, 'move':move, 'eq':eq, 'basis':basis, 'gain':gain})

    st.markdown("---")
    st.subheader("Other Assets")
    inh = st.radio("Will you inherit any assets from anyone else in a trust?", ["Yes", "No"], key="inh")
    ip = st.radio("Do you own any intellectual property?", ["Yes", "No"], key="ip")
    oth_assets = st.text_area("Other assets", key="oth_assets")

with t_health:
    st.warning(warning_msg)
    st.subheader("Health Directives")
    h1, h2 = st.columns(2)
    c_pn = h1.radio("Client do you have any objection to pain meds?", ["No", "Yes"], key="c_pn")
    s_pn = h2.radio("Spouse do you have any objection to pain meds?", ["No", "Yes"], key="s_pn")
    c_res = h1.radio("Client if you are found unresponsive, do you want EMT's to try to resuscitate you?", ["Yes", "No"], key="c_res")
    s_res = h2.radio("Spouse if you are found unresponsive, do you want EMT's to try to resuscitate you?", ["Yes", "No"], key="s_res")
    c_d = h1.radio("Client where do you prefer to die?", ["Home", "Facility"], key="c_d")
    s_d = h2.radio("Spouse where do you prefer to die?", ["Home", "Facility"], key="s_d")
    c_sp = h1.radio("Client can any Dr. treat you in an emergency, or do they need to consult a specialist first?", ["Yes", "No"], key="c_sp")
    s_sp = h2.radio("Spouse can any Dr. treat you in an emergency, or do they need to consult a specialist first?", ["Yes", "No"], key="s_sp")
    
    hc_spec_info = st.text_input("If you marked that a specialist needs to be consulted, please provide his/her name, phone, email, address:", key="hc_spec_info")
    
    h3, h4 = st.columns(2)
    c_bur = h3.radio("Client do you want to be:", ["Buried", "Cremated"], key="c_bur")
    s_bur = h4.radio("Spouse do you want to be:", ["Buried", "Cremated"], key="s_bur")
    
    hc_prepaid = st.text_input("If you have a prepaid funeral or dispostion, please provide the contract number, organization, and contact info:", key="hc_prepaid")
    hc_no_prepaid = st.text_input("If you do not have prepaid arrangements, how do you want your remains to be disposed of:", key="hc_no_prepaid")
    
    h5, h6 = st.columns(2)
    c_org = h5.radio("Client do you want to donaate organs?", ["All", "None", "Certain"], key="c_org")
    s_org = h6.radio("Spouse do you want to donaate organs?", ["All", "None", "Certain"], key="s_org")
    
    h7, h8 = st.columns(2)
    c_op = h7.radio("Client for what purpose", ["Transplant", "Research", "Stem Cell/Cloning", "Education"], key="c_op")
    s_op = h8.radio("Spouse for what purpose", ["Transplant", "Research", "Stem Cell/Cloning", "Education"], key="s_op")
    
    hc_organs_list = st.text_input("If you have any restrictions on organ donation, please provide details here:", key="hc_organs_list")
    
    h7a, h8a = st.columns(2)
    c_sui = h7a.radio("Client if legal, are you open to assisted suicide?", ["Yes", "No"], key="c_sui")
    s_sui = h8a.radio("Spouse if legal, are you open to assisted suicide?", ["Yes", "No"], key="s_sui")
    c_out = h7a.radio("Client do you want to make sure you spend time outdoors?", ["Yes", "No"], key="c_out")
    s_out = h8a.radio("Spouse do you want to make sure you spend time outdoors?", ["Yes", "No"], key="s_out")
    c_oth = h7a.radio("Client do you want to make sure you have social interaction?", ["Yes", "No"], key="c_oth")
    s_oth = h8a.radio("Spouse do you want to make sure you have social interaction?", ["Yes", "No"], key="s_oth")
    
    hc_allergies = st.text_input("Are you allergic to medications or anything found in a Dr.'s office or hospital:", key="hc_allergies")
    
    st.write("")
    h9, h10 = st.columns(2)
    ls_c = h9.radio("Client Life Support Choice", ["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], key="ls_c")
    ls_s = h10.radio("Spouse Life Support Choice", ["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], key="ls_s")

    st.markdown("---")
    st.subheader("Healthcare Agents")
    hca1, hca2 = st.columns(2)
    c_hca1 = hca1.text_input("Client HC Agent 1", key="c_hca1")
    c_hca2 = hca1.text_input("Client HC Agent 2", key="c_hca2")
    c_hca3 = hca1.text_input("Client HC Agent 3", key="c_hca3")
    c_hca4 = hca1.text_input("Client HC Agent 4", key="c_hca4")
    c_hca_act = hca1.selectbox("Client how do you want the above HC agents to act", ["Separate", "Jointly"], key="c_hca_act")
    c_hca_jt = hca1.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], key="chca_jt")

    s_hca1 = hca2.text_input("Spouse HC Agent 1", key="s_hca1")
    s_hca2 = hca2.text_input("Spouse HC Agent 2", key="s_hca2")
    s_hca3 = hca2.text_input("Spouse HC Agent 3", key="s_hca3")
    s_hca4 = hca2.text_input("Spouse HC Agent 4", key="s_hca4")
    s_hca_act = hca2.selectbox("Spouse how do you want the above HC agents to act", ["Separate", "Jointly"], key="s_hca_act")
    s_hca_jt = hca2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], key="shca_jt")
    
    st.markdown("---")
    st.subheader("Financial POA Agents")
    poa1, poa2 = st.columns(2)
    c_poa1 = poa1.text_input("Client POA 1", key="c_poa1")
    c_poa2 = poa1.text_input("Client POA 2", key="c_poa2")
    c_poa3 = poa1.text_input("Client POA 3", key="c_poa3")
    c_poa4 = poa1.text_input("Client POA 4", key="c_poa4")
    c_poa_act = poa1.selectbox("Client how do you want the above POA agents to act", ["Separate", "Jointly"], key="c_poa_act")
    c_poa_jt = poa1.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], key="cpoa_jt")
    c_poaw = poa1.radio("Client when do you want the power over your assets to take effect", ["Now", "Upon Incapacity"], key="c_poaw")
    c_poag = poa1.radio("Client Gifts", ["Yes", "No"], key="c_poag")

    s_poa1 = poa2.text_input("Spouse POA 1", key="s_poa1")
    s_poa2 = poa2.text_input("Spouse POA 2", key="s_poa2")
    s_poa3 = poa2.text_input("Spouse POA 3", key="s_poa3")
    s_poa4 = poa2.text_input("Spouse POA 4", key="s_poa4")
    s_poa_act = poa2.selectbox("Spouse how do you want the above POA agents to act", ["Separate", "Jointly"], key="s_poa_act")
    s_poa_jt = poa2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], key="spoa_jt")
    s_poaw = poa2.radio("Spouse when do you want the power over your assets to take effect", ["Now", "Upon Incapacity"], key="s_poaw")
    s_poag = poa2.radio("Spouse Gifts", ["Yes", "No"], key="s_poag")
    
with t_fiduciaries:
    st.warning(warning_msg)
    st.subheader("Guardians of the Person")
    gop1 = st.text_input("GOP 1", key="gop1")
    gop2 = st.text_input("GOP 2", key="gop2")
    gop3 = st.text_input("GOP 3", key="gop3")
    gop4 = st.text_input("GOP 4", key="gop4")
    gop_act = st.selectbox("GOP Act", ["Separate", "Jointly"], key="gop_act")
    gop_jt = st.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], key="gopjt")

    st.subheader("Guardians of the Estate")
    goe1 = st.text_input("GOE 1", key="goe1")
    goe2 = st.text_input("GOE 2", key="goe2")
    goe3 = st.text_input("GOE 3", key="goe3")
    goe4 = st.text_input("GOE 4", key="goe4")
    goe_act = st.selectbox("GOE Act", ["Separate", "Jointly"], key="goe_act")
    goe_jt = st.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], key="goejt")

    st.subheader("Trustees")
    tst1 = st.text_input("Trustee 1", key="tst1")
    tst2 = st.text_input("Trustee 2", key="tst2")
    tst3 = st.text_input("Trustee 3", key="tst3")
    tst4 = st.text_input("Trustee 4", key="tst4")
    tst_act = st.selectbox("How should the above trustees act?", ["Separate", "Jointly"], key="tst_act")
    tst_jt = st.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], key="tstjt")
    tst_comp = st.text_input("Do you want the trustee to be compensated; and if so, how (hourly, percentage, fixed amount, reasonable, etc.)?", key="tst_comp")

with t_distrib:
    st.warning(warning_msg)
    st.subheader("Beneficiaries")
    bens = []
    for i in range(1, 8):
        with st.expander(f"Beneficiary {i}"):
            nm = st.text_input("Name", key=f"ben{i}_nm")
            cy = st.radio("US Citizen?", ["Yes", "No"], key=f"ben{i}_c")
            pct = st.text_input("% or $ Amount", key=f"ben{i}_pct")
            dies = st.radio("If Dies, to", ["his/her kids", "Other"], key=f"ben{i}_d")
            dor = st.text_input("If Other, name", key=f"ben{i}_dor")
            td = st.text_input("What if they die?", key=f"ben{i}_td")
            gov = st.radio("Receiving Benefits?", ["Yes", "No"], key=f"ben{i}_g")
            how = st.radio("Inherit", ["In Trust", "Outright"], key=f"ben{i}_i")
            bens.append({'nm':nm, 'cy':cy, 'pct':pct, 'dies':dies, 'dor':dor, 'td':td, 'gov':gov, 'how':how})
            
    pets_instructions = st.text_input("Pets Instructions", key="pets_instructions")

    st.subheader("Specific Gifts")
    gifts = []
    for i in range(1, 7):
        with st.expander(f"Gift {i}"):
            desc = st.text_input("Gift", key=f"sg{i}_desc")
            nm = st.text_input("Beneficiary", key=f"sg{i}_nm")
            death = st.radio("Death of", ["H", "W", "Both"], key=f"sg{i}_d")
            dist = st.radio("Distribute", ["In Trust", "Outright", "Augment Existing SST"], key=f"sg{i}_di")
            unable = st.radio("If unable", ["to Issue", "Lapse", "Someone Else"], key=f"sg{i}_u")
            els = st.text_input("Else Name", key=f"sg{i}_else")
            gifts.append({'desc':desc, 'nm':nm, 'death':death, 'dist':dist, 'unable':unable, 'els':els})

    st.subheader("Disinheritance")
    cd_fail = st.multiselect("Client if everyone listed dies, rather than the assets going to the state, would you rather the assets go to heirs, charity, or both:", ["Family", "Charity"], key="cd_fail")
    sd_fail = st.multiselect("Spouse if everyone listed dies, rather than the assets going to the state, would you rather the assets go to heirs, charity, or both:", ["Family", "Charity"], key="sd_fail")
    
    disinh = []
    for i in range(1, 5):
        d1, d2, d3 = st.columns(3)
        disinh.append({
            'nm': d1.text_input(f"Person to Disinherit {i} Name", key=f"dis{i}_nm"),
            'rel': d2.text_input(f"Relation", key=f"dis{i}_rel"),
            'rsn': d3.text_input(f"Reason", key=f"dis{i}_rsn")
        })
    char_cause = st.text_input("Charity Cause", key="char_cause")

with t_admin:
    st.warning(warning_msg)
    st.subheader("Contacts (Master List)")
    contacts = []
    for i in range(1, 10):
        with st.expander(f"Contact {i}"):
            ct1, ct2, ct3, ct4 = st.columns(4)
            contacts.append({
                'nm': ct1.text_input("Name", key=f"ct{i}_nm"),
                'dob': ct2.text_input("DOB", key=f"ct{i}_dob"),
                'rel': ct3.text_input("Relation", key=f"ct{i}_rel"),
                'sbhw': ct4.selectbox("Related to", ["S/B", "H", "W", ""], key=f"ct{i}_sbhw"),
                'addr': st.text_input("Address", key=f"ct{i}_addr"),
                'ph': st.text_input("Phone", key=f"ct{i}_ph"),
                'em': st.text_input("Email", key=f"ct{i}_em")
            })

    st.subheader("General Information")
    g1, g2 = st.columns(2)
    ref_src = g1.text_input("Referral Source", key="ref_src")
    
    st.write("Documents / Services")
    doc_t = st.checkbox("Trust", key="doc_t")
    doc_w = st.checkbox("Will(s)", key="doc_w")
    doc_a = st.checkbox("AHCD(s)", key="doc_a")
    doc_p = st.checkbox("POA(s)", key="doc_p")
    doc_n = st.checkbox("NOG", key="doc_n")
    doc_ub = st.checkbox("Update Existing Binder", key="doc_ub")
    doc_nb = st.checkbox("Client Pay for New Binder", key="doc_nb")
    doc_h = st.checkbox("Homestead Declaration", key="doc_h")
    doc_hd = st.checkbox("Home Deed", key="doc_hd")
    doc_nhd = st.checkbox("Non-Home Deeds", key="doc_nhd")
    doc_tax = st.checkbox("Tax Planning Services", key="doc_tax")
    doc_biz = st.checkbox("Business Law Services", key="doc_biz")
    doc_ap = st.checkbox("Asset Protection Planning", key="doc_ap")
    doc_oth = st.checkbox("Other", key="doc_oth")
    other_notes = st.text_area("Other Notes", key="other_notes")

    # --- SAVE / SUBMIT BUTTONS ---
    st.markdown("---")
    st.subheader("Finish or Save Progress")
    
    # Generate JSON payload for downloading, excluding Streamlit internal objects
    draft_dict = {
        k: v for k, v in st.session_state.items() 
        if isinstance(v, (str, int, float, bool, list)) 
        and k != "draft_uploader" 
        and not k.startswith("FormSubmitter")
    }
    draft_json = json.dumps(draft_dict, indent=2)
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="💾 Save Draft to Computer",
            data=draft_json,
            file_name=f"EP_Draft_{c_id.replace(' ', '_') if c_id else 'Client'}.json",
            mime="application/json"
        )
    with col2:
        submit = st.button("Submit Final Questionnaire ➔")

    if submit:
        ctx = {}
        
        # Client
        ctx['c_id'] = c_id; ctx['c_address'] = c_address; ctx['c_county'] = c_county
        ctx['c_mailing'] = c_mailing; ctx['c_phone'] = c_phone; ctx['c_email'] = c_email
        ctx['c_dob'] = c_dob; ctx['c_pob'] = c_pob
        ctx['cb_c_male'] = cb(c_gender == "Male"); ctx['cb_c_female'] = cb(c_gender == "Female")
        ctx['cb_c_cit_y'] = cb(c_cit == "Yes"); ctx['cb_c_cit_n'] = cb(c_cit == "No")
        ctx['cb_c_wid_y'] = cb(c_wid == "Yes"); ctx['cb_c_wid_n'] = cb(c_wid == "No")
        ctx['cb_c_63_y'] = cb(c_63 == "Yes"); ctx['cb_c_63_n'] = cb(c_63 == "No")
        ctx['c_former_spouses'] = c_former_spouses; ctx['c_income'] = c_income; ctx['c_obligations'] = c_obligations

        # Spouse
        ctx['s_id'] = s_id; ctx['s_address'] = s_address; ctx['s_county'] = s_county
        ctx['s_mailing'] = s_mailing; ctx['s_phone'] = s_phone; ctx['s_email'] = s_email
        ctx['s_dob'] = s_dob; ctx['s_pob'] = s_pob
        ctx['cb_s_male'] = cb(s_gender == "Male"); ctx['cb_s_female'] = cb(s_gender == "Female")
        ctx['cb_s_cit_y'] = cb(s_cit == "Yes"); ctx['cb_s_cit_n'] = cb(s_cit == "No")
        ctx['cb_s_wid_y'] = cb(s_wid == "Yes"); ctx['cb_s_wid_n'] = cb(s_wid == "No")
        ctx['cb_s_63_y'] = cb(s_63 == "Yes"); ctx['cb_s_63_n'] = cb(s_63 == "No")
        ctx['s_former_spouses'] = s_former_spouses; ctx['s_income'] = s_income; ctx['s_obligations'] = s_obligations

        # Trust Info
        ctx['cb_prior_trust_y'] = cb(prior_trust == "Yes"); ctx['cb_prior_trust_n'] = cb(prior_trust == "No")
        ctx['t_amend_num'] = t_amend_num; ctx['t_desired_name'] = t_desired_name; ctx['t_reason'] = t_reason
        ctx['cb_li_rev_y'] = cb(li_rev == "Yes"); ctx['cb_li_rev_n'] = cb(li_rev in ["No", "Not Applicable"])
        ctx['cb_inv_rev_y'] = cb(inv_rev == "Yes"); ctx['cb_inv_rev_n'] = cb(inv_rev in ["No", "Not Applicable"])

        # Accounts
        for i, a in enumerate(accounts, 1):
            ctx[f'a{i}_type'] = a['type']; ctx[f'a{i}_val'] = a['val']
            ctx[f'a{i}_inst'] = a['inst']; ctx[f'a{i}_own'] = a['own']

        # Businesses
        for i, b in enumerate(businesses, 1):
            ctx[f'b{i}_name'] = b['name']; ctx[f'cb_b{i}_sole'] = cb(b['type']=="Sole Proprietor")
            ctx[f'cb_b{i}_part'] = cb(b['type']=="Partnership"); ctx[f'cb_b{i}_llc'] = cb(b['type']=="LLC")
            ctx[f'cb_b{i}_corp'] = cb(b['type']=="Corp"); ctx[f'b{i}_date'] = b['date']
            ctx[f'b{i}_ent'] = b['ent']; ctx[f'b{i}_state'] = b['state']
            ctx[f'cb_b{i}_td'] = cb(b['tax']=="Disregarded"); ctx[f'cb_b{i}_tp'] = cb(b['tax']=="Partnership")
            ctx[f'cb_b{i}_ts'] = cb(b['tax']=="S"); ctx[f'cb_b{i}_tc'] = cb(b['tax']=="C")
            ctx[f'b{i}_pct'] = b['pct']; ctx[f'b{i}_act'] = b['act']; ctx[f'b{i}_val'] = b['val']
            ctx[f'cb_b{i}_suc_y'] = cb(b['suc']=="Yes"); ctx[f'cb_b{i}_suc_n'] = cb(b['suc']=="No")
            ctx[f'cb_b{i}_dis_y'] = cb(b['dis']=="Yes"); ctx[f'cb_b{i}_dis_n'] = cb(b['dis']=="No")

        # Owed
        for i, o in enumerate(owed, 1):
            ctx[f'ow{i}_who'] = o['who']; ctx[f'ow{i}_amt'] = o['amt']; ctx[f'ow{i}_date'] = o['date']
            ctx[f'ow{i}_sec'] = o['sec']; ctx[f'ow{i}_purp'] = o['purp']; ctx[f'ow{i}_pb'] = o['pb']

        # Real Estate
        ctx['cb_hs_y'] = cb(hs_decl=="Yes"); ctx['cb_hs_n'] = cb(hs_decl=="No")
        ctx['re_pri_addr'] = re_pri_addr; ctx['re_pri_val'] = re_pri_val; ctx['re_pri_cnty'] = re_pri_cnty
        ctx['re_pri_move'] = re_pri_move; ctx['re_pri_eq'] = re_pri_eq; ctx['re_pri_basis'] = re_pri_basis; ctx['re_pri_gain'] = re_pri_gain
        ctx['re_sec_addr'] = re_sec_addr; ctx['re_sec_val'] = re_sec_val; ctx['re_sec_cnty'] = re_sec_cnty
        ctx['re_sec_move'] = re_sec_move; ctx['re_sec_eq'] = re_sec_eq; ctx['re_sec_basis'] = re_sec_basis; ctx['re_sec_gain'] = re_sec_gain
        for i, r in enumerate(rentals, 1):
            ctx[f'r{i}_addr'] = r['addr']; ctx[f'r{i}_val'] = r['val']; ctx[f'r{i}_cnty'] = r['cnty']
            ctx[f'r{i}_move'] = r['move']; ctx[f'r{i}_eq'] = r['eq']; ctx[f'r{i}_basis'] = r['basis']; ctx[f'r{i}_gain'] = r['gain']

        # Assets Other
        ctx['cb_inh_y'] = cb(inh=="Yes"); ctx['cb_inh_n'] = cb(inh=="No")
        ctx['cb_ip_y'] = cb(ip=="Yes"); ctx['cb_ip_n'] = cb(ip=="No"); ctx['oth_assets'] = oth_assets

        # Contacts
        for i, ct in enumerate(contacts, 1):
            ctx[f'ct{i}_nm'] = ct['nm']; ctx[f'ct{i}_dob'] = ct['dob']; ctx[f'ct{i}_rel'] = ct['rel']
            ctx[f'cb_ct{i}_sb'] = cb(ct['sbhw']=="S/B"); ctx[f'cb_ct{i}_h'] = cb(ct['sbhw']=="H"); ctx[f'cb_ct{i}_w'] = cb(ct['sbhw']=="W")
            ctx[f'ct{i}_addr'] = ct['addr']; ctx[f'ct{i}_ph'] = ct['ph']; ctx[f'ct{i}_em'] = ct['em']

        # Health
        ctx['cb_c_pn_y'] = cb(c_pn=="Yes"); ctx['cb_c_pn_n'] = cb(c_pn=="No"); ctx['cb_s_pn_y'] = cb(s_pn=="Yes"); ctx['cb_s_pn_n'] = cb(s_pn=="No")
        ctx['cb_c_res_y'] = cb(c_res=="Yes"); ctx['cb_c_res_n'] = cb(c_res=="No"); ctx['cb_s_res_y'] = cb(s_res=="Yes"); ctx['cb_s_res_n'] = cb(s_res=="No")
        ctx['cb_c_d_h'] = cb(c_d=="Home"); ctx['cb_c_d_f'] = cb(c_d=="Facility"); ctx['cb_s_d_h'] = cb(s_d=="Home"); ctx['cb_s_d_f'] = cb(s_d=="Facility")
        ctx['cb_c_sp_y'] = cb(c_sp=="Yes"); ctx['cb_c_sp_n'] = cb(c_sp=="No"); ctx['cb_s_sp_y'] = cb(s_sp=="Yes"); ctx['cb_s_sp_n'] = cb(s_sp=="No")
        ctx['hc_spec_info'] = hc_spec_info
        ctx['cb_c_bur'] = cb(c_bur=="Buried"); ctx['cb_c_crem'] = cb(c_bur=="Cremated"); ctx['cb_s_bur'] = cb(s_bur=="Buried"); ctx['cb_s_crem'] = cb(s_bur=="Cremated")
        ctx['hc_prepaid'] = hc_prepaid; ctx['hc_no_prepaid'] = hc_no_prepaid
        ctx['cb_c_o_all'] = cb(c_org=="All"); ctx['cb_c_o_non'] = cb(c_org=="None"); ctx['cb_c_o_cer'] = cb(c_org=="Certain")
        ctx['cb_s_o_all'] = cb(s_org=="All"); ctx['cb_s_o_non'] = cb(s_org=="None"); ctx['cb_s_o_cer'] = cb(s_org=="Certain")
        ctx['hc_organs_list'] = hc_organs_list
        ctx['cb_c_op_t'] = cb(c_op=="Transplant"); ctx['cb_c_op_r'] = cb(c_op=="Research"); ctx['cb_c_op_s'] = cb(c_op=="Stem Cell/Cloning"); ctx['cb_c_op_e'] = cb(c_op=="Education")
        ctx['cb_s_op_t'] = cb(s_op=="Transplant"); ctx['cb_s_op_r'] = cb(s_op=="Research"); ctx['cb_s_op_s'] = cb(s_op=="Stem Cell/Cloning"); ctx['cb_s_op_e'] = cb(s_op=="Education")
        ctx['cb_c_sui_y'] = cb(c_sui=="Yes"); ctx['cb_c_sui_n'] = cb(c_sui=="No"); ctx['cb_s_sui_y'] = cb(s_sui=="Yes"); ctx['cb_s_sui_n'] = cb(s_sui=="No")
        ctx['cb_c_out_y'] = cb(c_out=="Yes"); ctx['cb_c_out_n'] = cb(c_out=="No"); ctx['cb_s_out_y'] = cb(s_out=="Yes"); ctx['cb_s_out_n'] = cb(s_out=="No")
        ctx['cb_c_oth_y'] = cb(c_oth=="Yes"); ctx['cb_c_oth_n'] = cb(c_oth=="No"); ctx['cb_s_oth_y'] = cb(s_oth=="Yes"); ctx['cb_s_oth_n'] = cb(s_oth=="No")
        ctx['hc_allergies'] = hc_allergies
        ctx['cb_ls1_c'] = cb(ls_c=="No Support"); ctx['cb_ls2_c'] = cb(ls_c=="Prolong within limits"); ctx['cb_ls3_c'] = cb(ls_c=="Prolong if recover")
        ctx['cb_ls1_s'] = cb(ls_s=="No Support"); ctx['cb_ls2_s'] = cb(ls_s=="Prolong within limits"); ctx['cb_ls3_s'] = cb(ls_s=="Prolong if recover")

        # Agents
        ctx['c_hca1'] = c_hca1; ctx['c_hca2'] = c_hca2; ctx['c_hca3'] = c_hca3; ctx['c_hca4'] = c_hca4
        ctx['cb_c_hca_sep'] = cb(c_hca_act=="Separate"); ctx['cb_c_hca_jt'] = cb(c_hca_act=="Jointly")
        ctx['cb_c_hca_maj'] = cb(c_hca_jt=="Majority"); ctx['cb_c_hca_ind'] = cb(c_hca_jt=="Individual"); ctx['cb_c_hca_un'] = cb(c_hca_jt=="Unanimous")
        
        ctx['s_hca1'] = s_hca1; ctx['s_hca2'] = s_hca2; ctx['s_hca3'] = s_hca3; ctx['s_hca4'] = s_hca4
        ctx['cb_s_hca_sep'] = cb(s_hca_act=="Separate"); ctx['cb_s_hca_jt'] = cb(s_hca_act=="Jointly")
        ctx['cb_s_hca_maj'] = cb(s_hca_jt=="Majority"); ctx['cb_s_hca_ind'] = cb(s_hca_jt=="Individual"); ctx['cb_s_hca_un'] = cb(s_hca_jt=="Unanimous")

        # POA
        ctx['c_poa1'] = c_poa1; ctx['c_poa2'] = c_poa2; ctx['c_poa3'] = c_poa3; ctx['c_poa4'] = c_poa4
        ctx['cb_c_poa_sep'] = cb(c_poa_act=="Separate"); ctx['cb_c_poa_jt'] = cb(c_poa_act=="Jointly")
        ctx['cb_c_poa_maj'] = cb(c_poa_jt=="Majority"); ctx['cb_c_poa_ind'] = cb(c_poa_jt=="Individual"); ctx['cb_c_poa_un'] = cb(c_poa_jt=="Unanimous")
        ctx['cb_c_poaw_n'] = cb(c_poaw=="Now"); ctx['cb_c_poaw_i'] = cb(c_poaw=="Upon Incapacity")
        ctx['cb_c_poag_y'] = cb(c_poag=="Yes"); ctx['cb_c_poag_n'] = cb(c_poag=="No")

        ctx['s_poa1'] = s_poa1; ctx['s_poa2'] = s_poa2; ctx['s_poa3'] = s_poa3; ctx['s_poa4'] = s_poa4
        ctx['cb_s_poa_sep'] = cb(s_poa_act=="Separate"); ctx['cb_s_poa_jt'] = cb(s_poa_act=="Jointly")
        ctx['cb_s_poa_maj'] = cb(s_poa_jt=="Majority"); ctx['cb_s_poa_ind'] = cb(s_poa_jt=="Individual"); ctx['cb_s_poa_un'] = cb(s_poa_jt=="Unanimous")
        ctx['cb_s_poaw_n'] = cb(s_poaw=="Now"); ctx['cb_s_poaw_i'] = cb(s_poaw=="Upon Incapacity")
        ctx['cb_s_poag_y'] = cb(s_poag=="Yes"); ctx['cb_s_poag_n'] = cb(s_poag=="No")

        # Guardians
        ctx['gop1'] = gop1; ctx['gop2'] = gop2; ctx['gop3'] = gop3; ctx['gop4'] = gop4
        ctx['cb_gop_sep'] = cb(gop_act=="Separate"); ctx['cb_gop_jt'] = cb(gop_act=="Jointly")
        ctx['cb_gop_maj'] = cb(gop_jt=="Majority"); ctx['cb_gop_ind'] = cb(gop_jt=="Individual"); ctx['cb_gop_un'] = cb(gop_jt=="Unanimous")
        
        ctx['goe1'] = goe1; ctx['goe2'] = goe2; ctx['goe3'] = goe3; ctx['goe4'] = goe4
        ctx['cb_goe_sep'] = cb(goe_act=="Separate"); ctx['cb_goe_jt'] = cb(goe_act=="Jointly")
        ctx['cb_goe_maj'] = cb(goe_jt=="Majority"); ctx['cb_goe_ind'] = cb(goe_jt=="Individual"); ctx['cb_goe_un'] = cb(goe_jt=="Unanimous")

        # Trustees
        ctx['tst1'] = tst1; ctx['tst2'] = tst2; ctx['tst3'] = tst3; ctx['tst4'] = tst4
        ctx['cb_tst_sep'] = cb(tst_act=="Separate"); ctx['cb_tst_jt'] = cb(tst_act=="Jointly")
        ctx['cb_tst_maj'] = cb(tst_jt=="Majority"); ctx['cb_tst_ind'] = cb(tst_jt=="Individual"); ctx['cb_tst_un'] = cb(tst_jt=="Unanimous")
        ctx['tst_comp'] = tst_comp

        # Beneficiaries
        for i, b in enumerate(bens, 1):
            ctx[f'ben{i}_nm'] = b['nm']; ctx[f'cb_ben{i}_cy'] = cb(b['cy']=="Yes"); ctx[f'cb_ben{i}_cn'] = cb(b['cy']=="No")
            ctx[f'ben{i}_pct'] = b['pct']; ctx[f'cb_ben{i}_dk'] = cb(b['dies']=="his/her kids"); ctx[f'cb_ben{i}_do'] = cb(b['dies']=="Other")
            ctx[f'ben{i}_dor'] = b['dor']; ctx[f'ben{i}_td'] = b['td']
            ctx[f'cb_ben{i}_gy'] = cb(b['gov']=="Yes"); ctx[f'cb_ben{i}_gn'] = cb(b['gov']=="No")
            ctx[f'cb_ben{i}_it'] = cb(b['how']=="In Trust"); ctx[f'cb_ben{i}_out'] = cb(b['how']=="Outright")
        
        ctx['pets_instructions'] = pets_instructions

        # Gifts
        for i, g in enumerate(gifts, 1):
            ctx[f'sg{i}_desc'] = g['desc']; ctx[f'sg{i}_nm'] = g['nm']
            ctx[f'cb_sg{i}_h'] = cb(g['death']=="H"); ctx[f'cb_sg{i}_w'] = cb(g['death']=="W"); ctx[f'cb_sg{i}_b'] = cb(g['death']=="Both")
            ctx[f'cb_sg{i}_it'] = cb(g['dist']=="In Trust"); ctx[f'cb_sg{i}_out'] = cb(g['dist']=="Outright"); ctx[f'cb_sg{i}_aug'] = cb(g['dist']=="Augment Existing SST")
            ctx[f'cb_sg{i}_iss'] = cb(g['unable']=="to Issue"); ctx[f'cb_sg{i}_lap'] = cb(g['unable']=="Lapse"); ctx[f'cb_sg{i}_els'] = cb(g['unable']=="Someone Else")
            ctx[f'sg{i}_else'] = g['els']

        # Disinheritance
        ctx['cb_cd_fam'] = cb("Family" in cd_fail); ctx['cb_cd_char'] = cb("Charity" in cd_fail)
        ctx['cb_sd_fam'] = cb("Family" in sd_fail); ctx['cb_sd_char'] = cb("Charity" in sd_fail)
        for i, d in enumerate(disinh, 1):
            ctx[f'dis{i}_nm'] = d['nm']; ctx[f'dis{i}_rel'] = d['rel']; ctx[f'dis{i}_rsn'] = d['rsn']
        ctx['char_cause'] = char_cause

        # General
        ctx['ref_src'] = ref_src
        
        # Hidden Internal Fields
        ctx['sst_age'] = ""; ctx['cb_am1_y'] = cb(False); ctx['cb_am1_n'] = cb(False)
        ctx['cost_fee'] = ""; ctx['rel_plan'] = ""; ctx['cons_date'] = ""; ctx['sign_dc'] = ""
        ctx['trust_type'] = ""; ctx['cb_buy_y'] = cb(False); ctx['cb_buy_n'] = cb(False)
        ctx['cb_doc_t'] = cb(doc_t); ctx['cb_doc_w'] = cb(doc_w); ctx['cb_doc_a'] = cb(doc_a); ctx['cb_doc_p'] = cb(doc_p)
        ctx['cb_doc_n'] = cb(doc_n); ctx['cb_doc_ub'] = cb(doc_ub); ctx['cb_doc_nb'] = cb(doc_nb); ctx['cb_doc_h'] = cb(doc_h)
        ctx['cb_doc_hd'] = cb(doc_hd); ctx['cb_doc_nhd'] = cb(doc_nhd); ctx['cb_doc_tax'] = cb(doc_tax)
        ctx['cb_doc_biz'] = cb(doc_biz); ctx['cb_doc_ap'] = cb(doc_ap); ctx['cb_doc_oth'] = cb(doc_oth)
        ctx['other_notes'] = other_notes

        # File Processing
        input_template = "EP Q (with Fields) - Rev 7-17-26.docx"
        safe_name = c_id.replace(" ", "_") if c_id else "Client"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_filename = f"{safe_name}_EP_Questionnaire_{timestamp}.docx"
        temp_path = f"/tmp/{output_filename}"

        try:
            doc = DocxTemplate(input_template)
            doc.render(context=ctx)
            doc.save(temp_path)
            
            send_email_with_docx(temp_path, output_filename)
            os.remove(temp_path)
            
            st.success("Success! The completed questionnaire has been securely submitted.")
            
        except Exception as e:
            st.error(f"Error compiling document: {e}")
