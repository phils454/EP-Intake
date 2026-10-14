import streamlit as st
import os
import smtplib
from email.message import EmailMessage
from docxtpl import DocxTemplate
from datetime import datetime
import json
import io
import zipfile

# --- Configuration & Credentials ---
SENDER_EMAIL = "shekerlianlaw@gmail.com"
SENDER_PASSWORD = "vanwlgrbuojpnbrm"
RECEIVER_EMAILS = ["edwin@barthattorneys.com", "anu@barthattorneys.com"]

# --- INITIALIZE SESSION STATE FOR DYNAMIC TABLES ---
default_keys = {
    "c_id": "", "c_dob": "", "c_pob": "", "c_address": "", "c_county": "", "c_mailing": "",
    "c_phone": "", "c_email": "", "c_gen": "Male", "c_cit": "Yes", "c_wid": "No", "c_63": "No",
    "c_former_spouses": "", "c_income": "", "c_obligations": "",
    "s_id": "", "s_dob": "", "s_pob": "", "s_address": "", "s_county": "", "s_mailing": "",
    "s_phone": "", "s_email": "", "s_gen": "Female", "s_cit": "Yes", "s_wid": "No", "s_63": "No",
    "s_former_spouses": "", "s_income": "", "s_obligations": "",
    "prior_trust": "No", "t_amend_num": "", "t_desired_name": "", "t_reason": "",
    "li_rev": "Not Applicable", "inv_rev": "Not Applicable",
    "hs_decl": "No", "re_pri_addr": "", "re_pri_val": "", "re_pri_cnty": "", "re_pri_move": "", "re_pri_eq": "", "re_pri_basis": "", "re_pri_gain": "",
    "re_sec_addr": "", "re_sec_val": "", "re_sec_cnty": "", "re_sec_move": "", "re_sec_eq": "", "re_sec_basis": "", "re_sec_gain": "",
    "inh": "No", "ip": "No", "oth_assets": "",
    "c_pn": "No", "s_pn": "No", "c_res": "Yes", "s_res": "Yes", "c_d": "Home", "s_d": "Home", "c_sp": "Yes", "s_sp": "Yes",
    "hc_spec_info": "", "c_bur": "Buried", "s_bur": "Buried", "hc_prepaid": "", "hc_no_prepaid": "",
    "c_org": "All", "s_org": "All", "c_op": "Transplant", "s_op": "Transplant", "hc_organs_list": "",
    "c_sui": "No", "s_sui": "No", "c_out": "Yes", "s_out": "Yes", "c_oth": "Yes", "s_oth": "Yes",
    "hc_allergies": "", "ls_c": "No life support for any reason", "ls_s": "No life support for any reason",
    "c_hca_act": "Separate", "chca_jt": "Majority", "s_hca_act": "Separate", "shca_jt": "Majority",
    "c_poa_act": "Separate", "cpoa_jt": "Majority", "c_poaw": "Now", "c_poag": "Yes",
    "s_poa_act": "Separate", "spoa_jt": "Majority", "s_poaw": "Now", "s_poag": "Yes",
    "gop_act": "Separate", "gopjt": "Majority", "goe_act": "Separate", "goejt": "Majority",
    "tst_act": "Separate", "tstjt": "Majority", "tst_comp": "",
    "pets_instructions": "", "cd_fail": [], "sd_fail": [], "char_cause": "", "ref_src": "",
    "doc_t": False, "doc_w": False, "doc_a": False, "doc_p": False, "doc_n": False, "doc_ub": False, "doc_nb": False,
    "doc_h": False, "doc_hd": False, "doc_nhd": False, "doc_tax": False, "doc_biz": False, "doc_ap": False, "doc_oth": False, "other_notes": "",
    
    # Dynamic Tables Data
    "accounts_data": [{"Type": "", "Value": "", "Institution": "", "Owner": ""}],
    "businesses_data": [{"Name": "", "Type": "", "Date Started": "", "Entity #": "", "State": "", "Tax Election": "", "% Owned": "", "Activity": "", "Value": "", "Succession Plan?": "", "Disaster Plan?": ""}],
    "owed_data": [{"Who Owes": "", "Amount": "", "Date": "", "Secured": "", "Purpose": "", "Payback at death?": ""}],
    "rentals_data": [{"Address": "", "Value": "", "County": "", "Move Date": "", "Equity %": "", "Basis": "", "Gain": ""}],
    "c_hc_agents_data": [{"Name": ""}],
    "s_hc_agents_data": [{"Name": ""}],
    "c_poa_agents_data": [{"Name": ""}],
    "s_poa_agents_data": [{"Name": ""}],
    "gop_data": [{"Name": ""}],
    "goe_data": [{"Name": ""}],
    "trustees_data": [{"Name": ""}],
    "beneficiaries_data": [{"Name": "", "US Citizen?": "", "% or $ Amount": "", "If Dies, to": "", "If Other, name": "", "What if they die?": "", "Gov Benefits?": "", "Inherit": ""}],
    "gifts_data": [{"Gift": "", "Beneficiary Name": "", "Death of (H/W/Both)": "", "Distribute": "", "If unable": "", "Else Name": ""}],
    "disinherited_data": [{"Name": "", "Relation": "", "Reason": ""}],
    "contacts_data": [{"Name": "", "DOB": "", "Relation": "", "Related to (S/B/H/W)": "", "Address": "", "Phone": "", "Email": ""}]
}

for k, v in default_keys.items():
    if k not in st.session_state:
        st.session_state[k] = v

def cb(condition):
    return "☑" if condition else "☐"

# --- Main App Interface ---
st.set_page_config(page_title="BarthCalderon Intake Form", layout="wide")

# --- Sidebar: Draft Restore System ---
with st.sidebar:
    st.header("💾 Save & Resume Progress")
    st.write("Upload a `.json` file here to instantly restore your answers.")
    draft_file = st.file_uploader("Upload Draft File", type=["json"], key="draft_uploader")
    
    if draft_file is not None:
        try:
            saved_data = json.load(draft_file)
            for k, v in saved_data.items():
                if k != "draft_uploader":
                    st.session_state[k] = v
            st.success("✅ Draft loaded! Your answers have been fully restored.")
        except Exception as e:
            st.error("Error loading draft. Please ensure it is a valid file.")
            
    st.markdown("---")
    current_data = {k: st.session_state[k] for k in default_keys.keys() if k in st.session_state}
    json_data = json.dumps(current_data, indent=4)
    client_fname = st.session_state["c_id"].replace(" ", "_") if st.session_state["c_id"] else "Client"
    
    st.download_button(
        label="⬇️ Save Draft to Computer",
        data=json_data,
        file_name=f"EP_Draft_{client_fname}.json",
        mime="application/json",
        type="primary"
    )

st.title("BarthCalderon Estate Planning Questionnaire")
st.markdown("<mark style='background-color: yellow; font-weight: bold; padding: 2px 5px;'>PROGRESS IS NOT AUTOMATICALLY SAVED. YOU MUST SAVE IT VIA THE LEFT SIDEBAR OR ALL YOUR PROGRESS WILL BE LOST</mark>", unsafe_allow_html=True)
st.markdown("---")

t_client, t_assets, t_health, t_fiduciaries, t_distrib, t_admin = st.tabs([
    "1. Client & Spouse", "2. Assets & Real Estate", "3. Health & Agents", 
    "4. Guardians & Trustees", "5. Beneficiaries & Gifts", "6. Admin & General"
])

with t_client:
    st.subheader("Client Information")
    c1, c2, c3 = st.columns(3)
    st.session_state['c_id'] = c1.text_input("Client name as it appears on your ID", value=st.session_state['c_id'])
    st.session_state['c_dob'] = c2.text_input("Client DOB", value=st.session_state['c_dob'])
    st.session_state['c_pob'] = c3.text_input("Birth State/Country", value=st.session_state['c_pob'])
    
    c4, c5 = st.columns(2)
    st.session_state['c_address'] = c4.text_input("Residence Address", value=st.session_state['c_address'])
    st.session_state['c_county'] = c5.text_input("Residence County", value=st.session_state['c_county'])
    st.session_state['c_mailing'] = st.text_input("Mailing Address (if different)", value=st.session_state['c_mailing'])
    
    c6, c7 = st.columns(2)
    st.session_state['c_phone'] = c6.text_input("Primary Phone", value=st.session_state['c_phone'])
    st.session_state['c_email'] = c7.text_input("Email Address", value=st.session_state['c_email'])
    
    c8, c9, c10, c11 = st.columns(4)
    st.session_state['c_gen'] = c8.radio("Gender", ["Male", "Female"], index=["Male", "Female"].index(st.session_state['c_gen']))
    st.session_state['c_cit'] = c9.radio("US Citizen?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['c_cit']))
    st.session_state['c_wid'] = c10.radio("Widowed/Divorced?", ["No", "Yes"], index=["No", "Yes"].index(st.session_state['c_wid']))
    st.session_state['c_63'] = c11.radio("Over 63?", ["No", "Yes"], index=["No", "Yes"].index(st.session_state['c_63']))
    
    st.session_state['c_former_spouses'] = st.text_input("Name of former spouses & date of divorce or death (if applicable)", value=st.session_state['c_former_spouses'])
    
    c12, c13 = st.columns(2)
    st.session_state['c_income'] = c12.text_input("Approx. Income", value=st.session_state['c_income'])
    st.session_state['c_obligations'] = c13.text_input("Court Ordered Obligations", value=st.session_state['c_obligations'])

    st.markdown("---")
    st.subheader("Spouse Information")
    s1, s2, s3 = st.columns(3)
    st.session_state['s_id'] = s1.text_input("Spouse name as it appears on your ID", value=st.session_state['s_id'])
    st.session_state['s_dob'] = s2.text_input("Spouse DOB", value=st.session_state['s_dob'])
    st.session_state['s_pob'] = s3.text_input("Spouse Birth State/Country", value=st.session_state['s_pob'])
    
    s4, s5 = st.columns(2)
    st.session_state['s_address'] = s4.text_input("Spouse Residence Address", value=st.session_state['s_address'])
    st.session_state['s_county'] = s5.text_input("Spouse Residence County", value=st.session_state['s_county'])
    st.session_state['s_mailing'] = st.text_input("Spouse Mailing Address", value=st.session_state['s_mailing'])
    
    s6, s7 = st.columns(2)
    st.session_state['s_phone'] = s6.text_input("Spouse Phone", value=st.session_state['s_phone'])
    st.session_state['s_email'] = s7.text_input("Spouse Email", value=st.session_state['s_email'])
    
    s8, s9, s10, s11 = st.columns(4)
    st.session_state['s_gen'] = s8.radio("Spouse Gender", ["Male", "Female"], index=["Male", "Female"].index(st.session_state['s_gen']))
    st.session_state['s_cit'] = s9.radio("Spouse US Citizen?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['s_cit']))
    st.session_state['s_wid'] = s10.radio("Spouse Widowed/Divorced?", ["No", "Yes"], index=["No", "Yes"].index(st.session_state['s_wid']))
    st.session_state['s_63'] = s11.radio("Spouse Over 63?", ["No", "Yes"], index=["No", "Yes"].index(st.session_state['s_63']))
    
    st.session_state['s_former_spouses'] = st.text_input("Spouse Former Spouses & Dates", value=st.session_state['s_former_spouses'])
    
    s12, s13 = st.columns(2)
    st.session_state['s_income'] = s12.text_input("Spouse Approx. Income", value=st.session_state['s_income'])
    st.session_state['s_obligations'] = s13.text_input("Spouse Court Obligations", value=st.session_state['s_obligations'])

with t_assets:
    st.subheader("Trust History")
    ta1, ta2 = st.columns(2)
    st.session_state['prior_trust'] = ta1.radio("Do you already have a trust?", ["No", "Yes"], index=["No", "Yes"].index(st.session_state['prior_trust']))
    st.session_state['t_amend_num'] = ta2.text_input("Amendment/Restatement #", value=st.session_state['t_amend_num'])
    st.session_state['t_desired_name'] = st.text_input("Desired Name of Trust", value=st.session_state['t_desired_name'])
    st.session_state['t_reason'] = st.text_area("Why do you want a trust?", value=st.session_state['t_reason'])
    
    st.markdown("---")
    st.subheader("Financial Accounts")
    ta3, ta4 = st.columns(2)
    st.session_state['li_rev'] = ta3.radio("Have you had your life insurance reviewed recently?", ["Yes", "No", "Not Applicable"], index=["Yes", "No", "Not Applicable"].index(st.session_state['li_rev']))
    st.session_state['inv_rev'] = ta4.radio("Do you want a recommendation for a financial advisor?", ["Yes", "No", "Not Applicable"], index=["Yes", "No", "Not Applicable"].index(st.session_state['inv_rev']))
    
    st.session_state["accounts_data"] = st.data_editor(st.session_state["accounts_data"], num_rows="dynamic", use_container_width=True, key="accounts_editor")
            
    st.markdown("---")
    st.subheader("Businesses")
    st.session_state["businesses_data"] = st.data_editor(st.session_state["businesses_data"], num_rows="dynamic", use_container_width=True, key="biz_editor")

    st.markdown("---")
    st.subheader("Money Owed to You")
    st.session_state["owed_data"] = st.data_editor(st.session_state["owed_data"], num_rows="dynamic", use_container_width=True, key="owed_editor")

    st.markdown("---")
    st.subheader("Real Estate")
    st.session_state['hs_decl'] = st.radio("Do you have a homestead declaration filed for your primary home?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['hs_decl']))
    
    st.write("**Primary Residence**")
    st.session_state['re_pri_addr'] = st.text_input("Primary Residence Address", value=st.session_state['re_pri_addr'])
    re1, re2, re3, re4, re5 = st.columns(5)
    st.session_state['re_pri_val'] = re1.text_input("Value", value=st.session_state['re_pri_val'])
    st.session_state['re_pri_cnty'] = re2.text_input("County", value=st.session_state['re_pri_cnty'])
    st.session_state['re_pri_move'] = re3.text_input("Move In Date", value=st.session_state['re_pri_move'])
    st.session_state['re_pri_eq'] = re4.text_input("Equity %", value=st.session_state['re_pri_eq'])
    st.session_state['re_pri_basis'] = re5.text_input("Basis", value=st.session_state['re_pri_basis'])
    st.session_state['re_pri_gain'] = st.text_input("Gain", value=st.session_state['re_pri_gain'])
    
    st.write("**Secondary Residence**")
    st.session_state['re_sec_addr'] = st.text_input("Secondary Residence Address", value=st.session_state['re_sec_addr'])
    re6, re7, re8, re9, re10 = st.columns(5)
    st.session_state['re_sec_val'] = re6.text_input("Sec Value", value=st.session_state['re_sec_val'])
    st.session_state['re_sec_cnty'] = re7.text_input("Sec County", value=st.session_state['re_sec_cnty'])
    st.session_state['re_sec_move'] = re8.text_input("Sec Move Date", value=st.session_state['re_sec_move'])
    st.session_state['re_sec_eq'] = re9.text_input("Sec Equity %", value=st.session_state['re_sec_eq'])
    st.session_state['re_sec_basis'] = re10.text_input("Sec Basis", value=st.session_state['re_sec_basis'])
    st.session_state['re_sec_gain'] = st.text_input("Sec Gain", value=st.session_state['re_sec_gain'])
    
    st.write("**Rental Properties**")
    st.session_state["rentals_data"] = st.data_editor(st.session_state["rentals_data"], num_rows="dynamic", use_container_width=True, key="rentals_editor")

    st.markdown("---")
    st.subheader("Other Assets")
    st.session_state['inh'] = st.radio("Will you inherit any assets from anyone else in a trust?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['inh']))
    st.session_state['ip'] = st.radio("Do you own any intellectual property?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['ip']))
    st.session_state['oth_assets'] = st.text_area("Other assets", value=st.session_state['oth_assets'])

with t_health:
    st.subheader("Health Directives")
    h1, h2 = st.columns(2)
    st.session_state['c_pn'] = h1.radio("Client do you have any objection to pain meds?", ["No", "Yes"], index=["No", "Yes"].index(st.session_state['c_pn']))
    st.session_state['s_pn'] = h2.radio("Spouse do you have any objection to pain meds?", ["No", "Yes"], index=["No", "Yes"].index(st.session_state['s_pn']))
    st.session_state['c_res'] = h1.radio("Client if you are found unresponsive, do you want EMT's to try to resuscitate you?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['c_res']))
    st.session_state['s_res'] = h2.radio("Spouse if you are found unresponsive, do you want EMT's to try to resuscitate you?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['s_res']))
    st.session_state['c_d'] = h1.radio("Client where do you prefer to die?", ["Home", "Facility"], index=["Home", "Facility"].index(st.session_state['c_d']))
    st.session_state['s_d'] = h2.radio("Spouse where do you prefer to die?", ["Home", "Facility"], index=["Home", "Facility"].index(st.session_state['s_d']))
    st.session_state['c_sp'] = h1.radio("Client can any Dr. treat you in an emergency, or do they need to consult a specialist first?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['c_sp']))
    st.session_state['s_sp'] = h2.radio("Spouse can any Dr. treat you in an emergency, or do they need to consult a specialist first?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['s_sp']))
    
    st.session_state['hc_spec_info'] = st.text_input("If you marked that a specialist needs to be consulted, please provide his/her name, phone, email, address:", value=st.session_state['hc_spec_info'])
    
    h3, h4 = st.columns(2)
    st.session_state['c_bur'] = h3.radio("Client do you want to be:", ["Buried", "Cremated"], index=["Buried", "Cremated"].index(st.session_state['c_bur']))
    st.session_state['s_bur'] = h4.radio("Spouse do you want to be:", ["Buried", "Cremated"], index=["Buried", "Cremated"].index(st.session_state['s_bur']))
    
    st.session_state['hc_prepaid'] = st.text_input("If you have a prepaid funeral or dispostion, please provide the contract number, organization, and contact info:", value=st.session_state['hc_prepaid'])
    st.session_state['hc_no_prepaid'] = st.text_input("If you do not have prepaid arrangements, how do you want your remains to be disposed of:", value=st.session_state['hc_no_prepaid'])
    
    h5, h6 = st.columns(2)
    st.session_state['c_org'] = h5.radio("Client do you want to donate organs?", ["All", "None", "Certain"], index=["All", "None", "Certain"].index(st.session_state['c_org']))
    st.session_state['s_org'] = h6.radio("Spouse do you want to donate organs?", ["All", "None", "Certain"], index=["All", "None", "Certain"].index(st.session_state['s_org']))
    
    h7, h8 = st.columns(2)
    st.session_state['c_op'] = h7.radio("Client for what purpose", ["Transplant", "Research", "Stem Cell/Cloning", "Education"], index=["Transplant", "Research", "Stem Cell/Cloning", "Education"].index(st.session_state['c_op']))
    st.session_state['s_op'] = h8.radio("Spouse for what purpose", ["Transplant", "Research", "Stem Cell/Cloning", "Education"], index=["Transplant", "Research", "Stem Cell/Cloning", "Education"].index(st.session_state['s_op']))
    
    st.session_state['hc_organs_list'] = st.text_input("If you have any restrictions on organ donation, please provide details here:", value=st.session_state['hc_organs_list'])
    
    h7a, h8a = st.columns(2)
    st.session_state['c_sui'] = h7a.radio("Client if legal, are you open to assisted suicide?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['c_sui']))
    st.session_state['s_sui'] = h8a.radio("Spouse if legal, are you open to assisted suicide?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['s_sui']))
    st.session_state['c_out'] = h7a.radio("Client do you want to make sure you spend time outdoors?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['c_out']))
    st.session_state['s_out'] = h8a.radio("Spouse do you want to make sure you spend time outdoors?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['s_out']))
    st.session_state['c_oth'] = h7a.radio("Client do you want to make sure you have social interaction?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['c_oth']))
    st.session_state['s_oth'] = h8a.radio("Spouse do you want to make sure you have social interaction?", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['s_oth']))
    
    st.session_state['hc_allergies'] = st.text_input("Are you allergic to medications or anything found in a Dr.'s office or hospital:", value=st.session_state['hc_allergies'])
    
    st.write("")
    h9, h10 = st.columns(2)
    st.session_state['ls_c'] = h9.radio("Client Life Support Choice", ["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], index=["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"].index(st.session_state['ls_c']))
    st.session_state['ls_s'] = h10.radio("Spouse Life Support Choice", ["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], index=["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"].index(st.session_state['ls_s']))

    st.markdown("---")
    st.subheader("Healthcare Agents")
    st.write("Client HC Agents")
    st.session_state["c_hc_agents_data"] = st.data_editor(st.session_state["c_hc_agents_data"], num_rows="dynamic", use_container_width=True, key="chc_editor")
    c_h1, c_h2 = st.columns(2)
    st.session_state['c_hca_act'] = c_h1.selectbox("Client how do you want the above HC agents to act", ["Separate", "Jointly"], index=["Separate", "Jointly"].index(st.session_state['c_hca_act']))
    st.session_state['chca_jt'] = c_h2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=["Majority", "Individual", "Unanimous"].index(st.session_state['chca_jt']), key="chca_jt")

    st.write("Spouse HC Agents")
    st.session_state["s_hc_agents_data"] = st.data_editor(st.session_state["s_hc_agents_data"], num_rows="dynamic", use_container_width=True, key="shc_editor")
    s_h1, s_h2 = st.columns(2)
    st.session_state['s_hca_act'] = s_h1.selectbox("Spouse how do you want the above HC agents to act", ["Separate", "Jointly"], index=["Separate", "Jointly"].index(st.session_state['s_hca_act']))
    st.session_state['shca_jt'] = s_h2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=["Majority", "Individual", "Unanimous"].index(st.session_state['shca_jt']), key="shca_jt")
    
    st.markdown("---")
    st.subheader("Financial POA Agents")
    st.write("Client POA Agents")
    st.session_state["c_poa_agents_data"] = st.data_editor(st.session_state["c_poa_agents_data"], num_rows="dynamic", use_container_width=True, key="cpoa_editor")
    c_p1, c_p2, c_p3 = st.columns(3)
    st.session_state['c_poa_act'] = c_p1.selectbox("Client how do you want the above POA agents to act", ["Separate", "Jointly"], index=["Separate", "Jointly"].index(st.session_state['c_poa_act']))
    st.session_state['cpoa_jt'] = c_p2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=["Majority", "Individual", "Unanimous"].index(st.session_state['cpoa_jt']), key="cpoa_jt")
    st.session_state['c_poaw'] = c_p3.radio("Client when do you want the power to take effect", ["Now", "Upon Incapacity"], index=["Now", "Upon Incapacity"].index(st.session_state['c_poaw']))
    st.session_state['c_poag'] = st.radio("Client Gifts", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['c_poag']))

    st.write("Spouse POA Agents")
    st.session_state["s_poa_agents_data"] = st.data_editor(st.session_state["s_poa_agents_data"], num_rows="dynamic", use_container_width=True, key="spoa_editor")
    s_p1, s_p2, s_p3 = st.columns(3)
    st.session_state['s_poa_act'] = s_p1.selectbox("Spouse how do you want the above POA agents to act", ["Separate", "Jointly"], index=["Separate", "Jointly"].index(st.session_state['s_poa_act']))
    st.session_state['spoa_jt'] = s_p2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=["Majority", "Individual", "Unanimous"].index(st.session_state['spoa_jt']), key="spoa_jt")
    st.session_state['s_poaw'] = s_p3.radio("Spouse when do you want the power to take effect", ["Now", "Upon Incapacity"], index=["Now", "Upon Incapacity"].index(st.session_state['s_poaw']))
    st.session_state['s_poag'] = st.radio("Spouse Gifts", ["Yes", "No"], index=["Yes", "No"].index(st.session_state['s_poag']))
    
with t_fiduciaries:
    st.subheader("Guardians of the Person")
    st.session_state["gop_data"] = st.data_editor(st.session_state["gop_data"], num_rows="dynamic", use_container_width=True, key="gop_editor")
    gop_1, gop_2 = st.columns(2)
    st.session_state['gop_act'] = gop_1.selectbox("GOP Act", ["Separate", "Jointly"], index=["Separate", "Jointly"].index(st.session_state['gop_act']))
    st.session_state['gopjt'] = gop_2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=["Majority", "Individual", "Unanimous"].index(st.session_state['gopjt']), key="gopjt")

    st.subheader("Guardians of the Estate")
    st.session_state["goe_data"] = st.data_editor(st.session_state["goe_data"], num_rows="dynamic", use_container_width=True, key="goe_editor")
    goe_1, goe_2 = st.columns(2)
    st.session_state['goe_act'] = goe_1.selectbox("GOE Act", ["Separate", "Jointly"], index=["Separate", "Jointly"].index(st.session_state['goe_act']))
    st.session_state['goejt'] = goe_2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=["Majority", "Individual", "Unanimous"].index(st.session_state['goejt']), key="goejt")

    st.subheader("Trustees")
    st.session_state["trustees_data"] = st.data_editor(st.session_state["trustees_data"], num_rows="dynamic", use_container_width=True, key="tst_editor")
    tst_1, tst_2 = st.columns(2)
    st.session_state['tst_act'] = tst_1.selectbox("How should the above trustees act?", ["Separate", "Jointly"], index=["Separate", "Jointly"].index(st.session_state['tst_act']))
    st.session_state['tstjt'] = tst_2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=["Majority", "Individual", "Unanimous"].index(st.session_state['tstjt']), key="tstjt")
    st.session_state['tst_comp'] = st.text_input("Do you want the trustee to be compensated; and if so, how?", value=st.session_state['tst_comp'])

with t_distrib:
    st.subheader("Beneficiaries")
    st.session_state["beneficiaries_data"] = st.data_editor(st.session_state["beneficiaries_data"], num_rows="dynamic", use_container_width=True, key="ben_editor")
    st.session_state['pets_instructions'] = st.text_input("Pets Instructions", value=st.session_state['pets_instructions'])

    st.subheader("Specific Gifts")
    st.session_state["gifts_data"] = st.data_editor(st.session_state["gifts_data"], num_rows="dynamic", use_container_width=True, key="gifts_editor")

    st.subheader("Disinheritance")
    st.session_state['cd_fail'] = st.multiselect("Client if everyone listed dies, rather than the assets going to the state, would you rather the assets go to heirs, charity, or both:", ["Family", "Charity"], default=st.session_state['cd_fail'])
    st.session_state['sd_fail'] = st.multiselect("Spouse if everyone listed dies, rather than the assets going to the state, would you rather the assets go to heirs, charity, or both:", ["Family", "Charity"], default=st.session_state['sd_fail'])
    
    st.session_state["disinherited_data"] = st.data_editor(st.session_state["disinherited_data"], num_rows="dynamic", use_container_width=True, key="dis_editor")
    st.session_state['char_cause'] = st.text_input("Charity Cause", value=st.session_state['char_cause'])

with t_admin:
    st.subheader("Contacts (Master List)")
    st.session_state["contacts_data"] = st.data_editor(st.session_state["contacts_data"], num_rows="dynamic", use_container_width=True, key="contacts_editor")

    st.subheader("General Information")
    st.session_state['ref_src'] = st.text_input("Referral Source", value=st.session_state['ref_src'])
    
    st.write("Documents / Services")
    st.session_state['doc_t'] = st.checkbox("Trust", value=st.session_state['doc_t'])
    st.session_state['doc_w'] = st.checkbox("Will(s)", value=st.session_state['doc_w'])
    st.session_state['doc_a'] = st.checkbox("AHCD(s)", value=st.session_state['doc_a'])
    st.session_state['doc_p'] = st.checkbox("POA(s)", value=st.session_state['doc_p'])
    st.session_state['doc_n'] = st.checkbox("NOG", value=st.session_state['doc_n'])
    st.session_state['doc_ub'] = st.checkbox("Update Existing Binder", value=st.session_state['doc_ub'])
    st.session_state['doc_nb'] = st.checkbox("Client Pay for New Binder", value=st.session_state['doc_nb'])
    st.session_state['doc_h'] = st.checkbox("Homestead Declaration", value=st.session_state['doc_h'])
    st.session_state['doc_hd'] = st.checkbox("Home Deed", value=st.session_state['doc_hd'])
    st.session_state['doc_nhd'] = st.checkbox("Non-Home Deeds", value=st.session_state['doc_nhd'])
    st.session_state['doc_tax'] = st.checkbox("Tax Planning Services", value=st.session_state['doc_tax'])
    st.session_state['doc_biz'] = st.checkbox("Business Law Services", value=st.session_state['doc_biz'])
    st.session_state['doc_ap'] = st.checkbox("Asset Protection Planning", value=st.session_state['doc_ap'])
    st.session_state['doc_oth'] = st.checkbox("Other", value=st.session_state['doc_oth'])
    st.session_state['other_notes'] = st.text_area("Other Notes", value=st.session_state['other_notes'])

    # --- SUBMIT BUTTON ---
    st.markdown("---")
    st.subheader("Submit Questionnaire")
    
    submit = st.button("Generate Documents & Email Results", type="primary")

    if submit:
        with st.spinner("Generating document and emailing securely..."):
            ctx = {}
            
            # Client & Spouse & General
            for k in ["c_id", "c_address", "c_county", "c_mailing", "c_phone", "c_email", "c_dob", "c_pob",
                      "c_former_spouses", "c_income", "c_obligations", "s_id", "s_address", "s_county",
                      "s_mailing", "s_phone", "s_email", "s_dob", "s_pob", "s_former_spouses", "s_income",
                      "s_obligations", "t_amend_num", "t_desired_name", "t_reason", "re_pri_addr", "re_pri_val",
                      "re_pri_cnty", "re_pri_move", "re_pri_eq", "re_pri_basis", "re_pri_gain", "re_sec_addr",
                      "re_sec_val", "re_sec_cnty", "re_sec_move", "re_sec_eq", "re_sec_basis", "re_sec_gain",
                      "oth_assets", "hc_spec_info", "hc_prepaid", "hc_no_prepaid", "hc_organs_list", "hc_allergies",
                      "pets_instructions", "char_cause", "ref_src", "other_notes", "tst_comp",
                      "c_hca_act", "chca_jt", "s_hca_act", "shca_jt", "c_poa_act", "cpoa_jt", "s_poa_act", "spoa_jt",
                      "gop_act", "gopjt", "goe_act", "goejt", "tst_act", "tstjt"]:
                ctx[k] = st.session_state[k]

            # Logic Checkboxes / Radios
            ctx['cb_c_male'] = cb(st.session_state['c_gen'] == "Male"); ctx['cb_c_female'] = cb(st.session_state['c_gen'] == "Female")
            ctx['cb_c_cit_y'] = cb(st.session_state['c_cit'] == "Yes"); ctx['cb_c_cit_n'] = cb(st.session_state['c_cit'] == "No")
            ctx['cb_c_wid_y'] = cb(st.session_state['c_wid'] == "Yes"); ctx['cb_c_wid_n'] = cb(st.session_state['c_wid'] == "No")
            ctx['cb_c_63_y'] = cb(st.session_state['c_63'] == "Yes"); ctx['cb_c_63_n'] = cb(st.session_state['c_63'] == "No")
            ctx['cb_s_male'] = cb(st.session_state['s_gen'] == "Male"); ctx['cb_s_female'] = cb(st.session_state['s_gen'] == "Female")
            ctx['cb_s_cit_y'] = cb(st.session_state['s_cit'] == "Yes"); ctx['cb_s_cit_n'] = cb(st.session_state['s_cit'] == "No")
            ctx['cb_s_wid_y'] = cb(st.session_state['s_wid'] == "Yes"); ctx['cb_s_wid_n'] = cb(st.session_state['s_wid'] == "No")
            ctx['cb_s_63_y'] = cb(st.session_state['s_63'] == "Yes"); ctx['cb_s_63_n'] = cb(st.session_state['s_63'] == "No")
            ctx['cb_prior_trust_y'] = cb(st.session_state['prior_trust'] == "Yes"); ctx['cb_prior_trust_n'] = cb(st.session_state['prior_trust'] == "No")
            ctx['cb_li_rev_y'] = cb(st.session_state['li_rev'] == "Yes"); ctx['cb_li_rev_n'] = cb(st.session_state['li_rev'] in ["No", "Not Applicable"])
            ctx['cb_inv_rev_y'] = cb(st.session_state['inv_rev'] == "Yes"); ctx['cb_inv_rev_n'] = cb(st.session_state['inv_rev'] in ["No", "Not Applicable"])
            ctx['cb_hs_y'] = cb(st.session_state['hs_decl'] == "Yes"); ctx['cb_hs_n'] = cb(st.session_state['hs_decl'] == "No")
            ctx['cb_inh_y'] = cb(st.session_state['inh'] == "Yes"); ctx['cb_inh_n'] = cb(st.session_state['inh'] == "No")
            ctx['cb_ip_y'] = cb(st.session_state['ip'] == "Yes"); ctx['cb_ip_n'] = cb(st.session_state['ip'] == "No")
            ctx['cb_c_pn_y'] = cb(st.session_state['c_pn'] == "Yes"); ctx['cb_c_pn_n'] = cb(st.session_state['c_pn'] == "No")
            ctx['cb_s_pn_y'] = cb(st.session_state['s_pn'] == "Yes"); ctx['cb_s_pn_n'] = cb(st.session_state['s_pn'] == "No")
            ctx['cb_c_res_y'] = cb(st.session_state['c_res'] == "Yes"); ctx['cb_c_res_n'] = cb(st.session_state['c_res'] == "No")
            ctx['cb_s_res_y'] = cb(st.session_state['s_res'] == "Yes"); ctx['cb_s_res_n'] = cb(st.session_state['s_res'] == "No")
            ctx['cb_c_d_h'] = cb(st.session_state['c_d'] == "Home"); ctx['cb_c_d_f'] = cb(st.session_state['c_d'] == "Facility")
            ctx['cb_s_d_h'] = cb(st.session_state['s_d'] == "Home"); ctx['cb_s_d_f'] = cb(st.session_state['s_d'] == "Facility")
            ctx['cb_c_sp_y'] = cb(st.session_state['c_sp'] == "Yes"); ctx['cb_c_sp_n'] = cb(st.session_state['c_sp'] == "No")
            ctx['cb_s_sp_y'] = cb(st.session_state['s_sp'] == "Yes"); ctx['cb_s_sp_n'] = cb(st.session_state['s_sp'] == "No")
            ctx['cb_c_bur'] = cb(st.session_state['c_bur'] == "Buried"); ctx['cb_c_crem'] = cb(st.session_state['c_bur'] == "Cremated")
            ctx['cb_s_bur'] = cb(st.session_state['s_bur'] == "Buried"); ctx['cb_s_crem'] = cb(st.session_state['s_bur'] == "Cremated")
            ctx['cb_c_o_all'] = cb(st.session_state['c_org'] == "All"); ctx['cb_c_o_non'] = cb(st.session_state['c_org'] == "None"); ctx['cb_c_o_cer'] = cb(st.session_state['c_org'] == "Certain")
            ctx['cb_s_o_all'] = cb(st.session_state['s_org'] == "All"); ctx['cb_s_o_non'] = cb(st.session_state['s_org'] == "None"); ctx['cb_s_o_cer'] = cb(st.session_state['s_org'] == "Certain")
            ctx['cb_c_op_t'] = cb(st.session_state['c_op'] == "Transplant"); ctx['cb_c_op_r'] = cb(st.session_state['c_op'] == "Research"); ctx['cb_c_op_s'] = cb(st.session_state['c_op'] == "Stem Cell/Cloning"); ctx['cb_c_op_e'] = cb(st.session_state['c_op'] == "Education")
            ctx['cb_s_op_t'] = cb(st.session_state['s_op'] == "Transplant"); ctx['cb_s_op_r'] = cb(st.session_state['s_op'] == "Research"); ctx['cb_s_op_s'] = cb(st.session_state['s_op'] == "Stem Cell/Cloning"); ctx['cb_s_op_e'] = cb(st.session_state['s_op'] == "Education")
            ctx['cb_c_sui_y'] = cb(st.session_state['c_sui'] == "Yes"); ctx['cb_c_sui_n'] = cb(st.session_state['c_sui'] == "No")
            ctx['cb_s_sui_y'] = cb(st.session_state['s_sui'] == "Yes"); ctx['cb_s_sui_n'] = cb(st.session_state['s_sui'] == "No")
            ctx['cb_c_out_y'] = cb(st.session_state['c_out'] == "Yes"); ctx['cb_c_out_n'] = cb(st.session_state['c_out'] == "No")
            ctx['cb_s_out_y'] = cb(st.session_state['s_out'] == "Yes"); ctx['cb_s_out_n'] = cb(st.session_state['s_out'] == "No")
            ctx['cb_c_oth_y'] = cb(st.session_state['c_oth'] == "Yes"); ctx['cb_c_oth_n'] = cb(st.session_state['c_oth'] == "No")
            ctx['cb_s_oth_y'] = cb(st.session_state['s_oth'] == "Yes"); ctx['cb_s_oth_n'] = cb(st.session_state['s_oth'] == "No")
            ctx['cb_ls1_c'] = cb(st.session_state['ls_c'] == "No life support for any reason"); ctx['cb_ls2_c'] = cb(st.session_state['ls_c'] == "Life support for as long as legally possible"); ctx['cb_ls3_c'] = cb(st.session_state['ls_c'] == "Life support only if it results in relatively full recovery in a reasonably short time")
            ctx['cb_ls1_s'] = cb(st.session_state['ls_s'] == "No life support for any reason"); ctx['cb_ls2_s'] = cb(st.session_state['ls_s'] == "Life support for as long as legally possible"); ctx['cb_ls3_s'] = cb(st.session_state['ls_s'] == "Life support only if it results in relatively full recovery in a reasonably short time")
            
            # Fiduciary Checkboxes
            ctx['cb_c_hca_sep'] = cb(st.session_state['c_hca_act'] == "Separate"); ctx['cb_c_hca_jt'] = cb(st.session_state['c_hca_act'] == "Jointly")
            ctx['cb_c_hca_maj'] = cb(st.session_state['chca_jt'] == "Majority"); ctx['cb_c_hca_ind'] = cb(st.session_state['chca_jt'] == "Individual"); ctx['cb_c_hca_un'] = cb(st.session_state['chca_jt'] == "Unanimous")
            
            ctx['cb_s_hca_sep'] = cb(st.session_state['s_hca_act'] == "Separate"); ctx['cb_s_hca_jt'] = cb(st.session_state['s_hca_act'] == "Jointly")
            ctx['cb_s_hca_maj'] = cb(st.session_state['shca_jt'] == "Majority"); ctx['cb_s_hca_ind'] = cb(st.session_state['shca_jt'] == "Individual"); ctx['cb_s_hca_un'] = cb(st.session_state['shca_jt'] == "Unanimous")

            ctx['cb_c_poa_sep'] = cb(st.session_state['c_poa_act'] == "Separate"); ctx['cb_c_poa_jt'] = cb(st.session_state['c_poa_act'] == "Jointly")
            ctx['cb_c_poa_maj'] = cb(st.session_state['cpoa_jt'] == "Majority"); ctx['cb_c_poa_ind'] = cb(st.session_state['cpoa_jt'] == "Individual"); ctx['cb_c_poa_un'] = cb(st.session_state['cpoa_jt'] == "Unanimous")

            ctx['cb_s_poa_sep'] = cb(st.session_state['s_poa_act'] == "Separate"); ctx['cb_s_poa_jt'] = cb(st.session_state['s_poa_act'] == "Jointly")
            ctx['cb_s_poa_maj'] = cb(st.session_state['spoa_jt'] == "Majority"); ctx['cb_s_poa_ind'] = cb(st.session_state['spoa_jt'] == "Individual"); ctx['cb_s_poa_un'] = cb(st.session_state['spoa_jt'] == "Unanimous")

            ctx['cb_gop_sep'] = cb(st.session_state['gop_act'] == "Separate"); ctx['cb_gop_jt'] = cb(st.session_state['gop_act'] == "Jointly")
            ctx['cb_gop_maj'] = cb(st.session_state['gopjt'] == "Majority"); ctx['cb_gop_ind'] = cb(st.session_state['gopjt'] == "Individual"); ctx['cb_gop_un'] = cb(st.session_state['gopjt'] == "Unanimous")

            ctx['cb_goe_sep'] = cb(st.session_state['goe_act'] == "Separate"); ctx['cb_goe_jt'] = cb(st.session_state['goe_act'] == "Jointly")
            ctx['cb_goe_maj'] = cb(st.session_state['goejt'] == "Majority"); ctx['cb_goe_ind'] = cb(st.session_state['goejt'] == "Individual"); ctx['cb_goe_un'] = cb(st.session_state['goejt'] == "Unanimous")

            ctx['cb_tst_sep'] = cb(st.session_state['tst_act'] == "Separate"); ctx['cb_tst_jt'] = cb(st.session_state['tst_act'] == "Jointly")
            ctx['cb_tst_maj'] = cb(st.session_state['tstjt'] == "Majority"); ctx['cb_tst_ind'] = cb(st.session_state['tstjt'] == "Individual"); ctx['cb_tst_un'] = cb(st.session_state['tstjt'] == "Unanimous")

            ctx['cb_c_poaw_n'] = cb(st.session_state['c_poaw'] == "Now"); ctx['cb_c_poaw_i'] = cb(st.session_state['c_poaw'] == "Upon Incapacity")
            ctx['cb_c_poag_y'] = cb(st.session_state['c_poag'] == "Yes"); ctx['cb_c_poag_n'] = cb(st.session_state['c_poag'] == "No")
            ctx['cb_s_poaw_n'] = cb(st.session_state['s_poaw'] == "Now"); ctx['cb_s_poaw_i'] = cb(st.session_state['s_poaw'] == "Upon Incapacity")
            ctx['cb_s_poag_y'] = cb(st.session_state['s_poag'] == "Yes"); ctx['cb_s_poag_n'] = cb(st.session_state['s_poag'] == "No")
            
            ctx['cb_cd_fam'] = cb("Family" in st.session_state['cd_fail']); ctx['cb_cd_char'] = cb("Charity" in st.session_state['cd_fail'])
            ctx['cb_sd_fam'] = cb("Family" in st.session_state['sd_fail']); ctx['cb_sd_char'] = cb("Charity" in st.session_state['sd_fail'])
            
            ctx['cb_doc_t'] = cb(st.session_state['doc_t']); ctx['cb_doc_w'] = cb(st.session_state['doc_w']); ctx['cb_doc_a'] = cb(st.session_state['doc_a']); ctx['cb_doc_p'] = cb(st.session_state['doc_p'])
            ctx['cb_doc_n'] = cb(st.session_state['doc_n']); ctx['cb_doc_ub'] = cb(st.session_state['doc_ub']); ctx['cb_doc_nb'] = cb(st.session_state['doc_nb']); ctx['cb_doc_h'] = cb(st.session_state['doc_h'])
            ctx['cb_doc_hd'] = cb(st.session_state['doc_hd']); ctx['cb_doc_nhd'] = cb(st.session_state['doc_nhd']); ctx['cb_doc_tax'] = cb(st.session_state['doc_tax'])
            ctx['cb_doc_biz'] = cb(st.session_state['doc_biz']); ctx['cb_doc_ap'] = cb(st.session_state['doc_ap']); ctx['cb_doc_oth'] = cb(st.session_state['doc_oth'])
            
            ctx['cb_am1_y'] = cb(False); ctx['cb_am1_n'] = cb(False); ctx['cb_buy_y'] = cb(False); ctx['cb_buy_n'] = cb(False)
            ctx['sst_age'] = ""; ctx['cost_fee'] = ""; ctx['rel_plan'] = ""; ctx['cons_date'] = ""; ctx['sign_dc'] = ""; ctx['trust_type'] = ""

            # Inject Dynamic Tables
            ctx['accounts'] = st.session_state['accounts_data']
            ctx['businesses'] = st.session_state['businesses_data']
            ctx['owed'] = st.session_state['owed_data']
            ctx['rentals'] = st.session_state['rentals_data']
            ctx['c_hc_agents'] = st.session_state['c_hc_agents_data']
            ctx['s_hc_agents'] = st.session_state['s_hc_agents_data']
            ctx['c_poa_agents'] = st.session_state['c_poa_agents_data']
            ctx['s_poa_agents'] = st.session_state['s_poa_agents_data']
            ctx['gop_agents'] = st.session_state['gop_data']
            ctx['goe_agents'] = st.session_state['goe_data']
            ctx['trustees'] = st.session_state['trustees_data']
            ctx['beneficiaries'] = st.session_state['beneficiaries_data']
            ctx['gifts'] = st.session_state['gifts_data']
            ctx['disinherited'] = st.session_state['disinherited_data']
            ctx['contacts'] = st.session_state['contacts_data']

            # File Processing
            input_template = "EP Q (with Fields) - Rev 7-17-26.docx"
            safe_name = st.session_state['c_id'].replace(" ", "_") if st.session_state['c_id'] else "Client"
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            output_filename = f"{safe_name}_EP_Questionnaire_{timestamp}.docx"
            temp_path = f"{output_filename}"

            try:
                # Generate Document
                doc = DocxTemplate(input_template)
                doc.render(context=ctx)
                doc.save(temp_path)
                
                # Send Email
                msg = EmailMessage()
                msg['Subject'] = f"New EP Intake: {output_filename}"
                msg['From'] = SENDER_EMAIL
                msg['To'] = ", ".join(RECEIVER_EMAILS)
                
                body_text = "A new client intake questionnaire has been completed.\n\nPlease find two files attached:\n1. The generated Word document.\n2. A JSON data file containing all of the client's form inputs (upload this back into the sidebar of the app to restore/edit the answers if needed)."
                msg.set_content(body_text)

                # Attach Word Doc
                with open(temp_path, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.wordprocessingml.document', filename=output_filename)

                # Attach JSON
                email_data = {k: st.session_state[k] for k in default_keys.keys() if k in st.session_state}
                email_json_data = json.dumps(email_data, indent=4)
                json_name = f"EP_Client_Data_{safe_name}.json"
                msg.add_attachment(email_json_data.encode('utf-8'), maintype='application', subtype='json', filename=json_name)

                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
                    smtp.send_message(msg)

                os.remove(temp_path)
                st.success("✅ Success! The completed questionnaire and JSON backup have been securely emailed to your office.")
                
            except Exception as e:
                st.error(f"Error compiling document: {e}")
```

```text:EP Questionnaire Text:ep_q_dynamic.txt
estate planning Questionnaire

[BarthCalderon Logo]
This questionnaire will solely be used by our office and since your answers go to an attorney, you are assured confidentiality and protected by the attorney-client privilege. Please answer completely and truthfully. If you run out of room filling anything in, please print the page again and attach it. PRINT CLEARLY USING BLUE INK.

Client Information
Type of valid ID and name as it appears on it (please print): **{{ c_id }}** 
Residence Address: **{{ c_address }}**;  Residence County: **{{ c_county }}**
Mailing Address: **{{ c_mailing }}**; Primary Phone Number: **{{ c_phone }}** 
Email Address: **{{ c_email }}**; Date of Birth: **{{ c_dob }}**; Birth State or Country: **{{ c_pob }}** 
Are You: {{ cb_c_male }} Male {{ cb_c_female }} Female; US Citizen: {{ cb_c_cit_y }} Yes {{ cb_c_cit_n }} No; Widow/Divorced: {{ cb_c_wid_y }} Yes {{ cb_c_wid_n }} No; Over 63: {{ cb_c_63_y }} Yes {{ cb_c_63_n }} No 
If widow/divorce, all former spouse's full names & dates of death/divorce: **{{ c_former_spouses }}**
Approx. income $**{{ c_income }}**; Amt. court ordered financial obligations: **{{ c_obligations }}**

Spouse Information
Type of valid ID and name as it appears on it (please print): **{{ s_id }}** 
Residence Address: **{{ s_address }}**; Residence County: **{{ s_county }}**
Mailing Address: **{{ s_mailing }}**; Primary Phone Number: **{{ s_phone }}** 
Email Address: **{{ s_email }}**; Date of Birth: **{{ s_dob }}**; Birth State or Country: **{{ s_pob }}** 
Are You: {{ cb_s_male }} Male {{ cb_s_female }} Female; US Citizen: {{ cb_s_cit_y }} Yes {{ cb_s_cit_n }} No; Widow/Divorced: {{ cb_s_wid_y }} Yes {{ cb_s_wid_n }} No; Over 63: {{ cb_s_63_y }} Yes {{ cb_s_63_n }} No 
If widow/divorce, all former spouse's full names & dates of death/divorce: **{{ s_former_spouses }}** 
Approx. income $**{{ s_income }}**; Amt. court ordered financial obligations: **{{ s_obligations }}**

Trust Information
Have you ever had a trust before: {{ cb_prior_trust_n }} No {{ cb_prior_trust_y }} Yes (If Yes, you MUST provide a copy of the most recent trust) 
Amendment/Restatement #: **{{ t_amend_num }}**; Desired name of trust: The **{{ t_desired_name }}** Trust 
We know why you need a trust, but why do you want one: **{{ t_reason }}**

Assets (For Owner: H = Husband, W = Wife, S/B = Client only or Both H & W)
Cash, Bank Accounts, CD’s, Bonds, Retirement, Life Insurance, Annuities, Securities, Timeshares: 
Have you had your life insurance policy reviewed in the last year? {{ cb_li_rev_y }} Yes {{ cb_li_rev_n }} No 
Do you have someone that oversees all your investments to ensure you are meeting your goals? {{ cb_inv_rev_y }} Yes {{ cb_inv_rev_n }} No 

{% for account in accounts %}
Account Type: **{{ account['Type'] }}**; Approx Value: **{{ account['Value'] }}**; Institution Name: **{{ account['Institution'] }}**; Owner: **{{ account['Owner'] }}** 
{% endfor %}

Businesses or Hobbies that Make Money: 
{% for biz in businesses %}
Business Name: **{{ biz['Name'] }}**; Type: **{{ biz['Type'] }}** 
Date Started: **{{ biz['Date Started'] }}**; Entity # **{{ biz['Entity #'] }}**; State: **{{ biz['State'] }}**; Tax Election: **{{ biz['Tax Election'] }}**
# of Shares or % Owned: **{{ biz['% Owned'] }}**; Business Activity: **{{ biz['Activity'] }}**; Approx. Value: **{{ biz['Value'] }}**
Business Succession Plan: **{{ biz['Succession Plan?'] }}**; Disaster or Disability Plan: **{{ biz['Disaster Plan?'] }}**
{% endfor %}

Money Owed to you: 
{% for loan in owed %}
Who Owes: **{{ loan['Who Owes'] }}**; How Much: **{{ loan['Amount'] }}**; Date of Loan: **{{ loan['Date'] }}**; Secured: **{{ loan['Secured'] }}**; 
Purpose: **{{ loan['Purpose'] }}**; Does it Need to be Paid Back at Death: **{{ loan['Payback at death?'] }}** 
{% endfor %}

Real Estate: 
Do you have a homestead declaration filed to protect equity in your home? {{ cb_hs_y }} Yes {{ cb_hs_n }} No 
Primary Residence Address: **{{ re_pri_addr }}**; Current Value: **{{ re_pri_val }}** 
County: **{{ re_pri_cnty }}**; Move In Date: **{{ re_pri_move }}**; Equity %: **{{ re_pri_eq }}**; Basis: **{{ re_pri_basis }}**; Gain: **{{ re_pri_gain }}**

Secondary Residence Address: **{{ re_sec_addr }}**; Current Value: **{{ re_sec_val }}** 
County: **{{ re_sec_cnty }}**; Move In Date: **{{ re_sec_move }}**; Equity %: **{{ re_sec_eq }}**; Basis: **{{ re_sec_basis }}**; Gain: **{{ re_sec_gain }}**

{% for rental in rentals %}
Rental Property Address: **{{ rental['Address'] }}**; Current Value: **{{ rental['Value'] }}** 
County: **{{ rental['County'] }}**; Move In Date: **{{ rental['Move Date'] }}**; Equity %: **{{ rental['Equity %'] }}**; Basis: **{{ rental['Basis'] }}**; Gain: **{{ rental['Gain'] }}** 
{% endfor %}

Other Assets: 
Will you inherit anything from a trust (from parents, siblings, etc.)? {{ cb_inh_y }} Yes {{ cb_inh_n }} No 
Do you own any intellectual property (Patents, Trademarks, Copyrights, etc.)? {{ cb_ip_y }} Yes {{ cb_ip_n }} No 
Please list any additional assets here: **{{ oth_assets }}**

Contact Information (H = Husband, W = Wife, S/B = Client only or Both H & W)
Please list all of the following information for everyone you have named in this form (other than yourselves):
{% for contact in contacts %}
Full Legal Name: **{{ contact['Name'] }}**; Date of Birth: **{{ contact['DOB'] }}**; Relationship to you: **{{ contact['Relation'] }}** 
Related to: **{{ contact['Related to (S/B/H/W)'] }}**; Mailing Address: **{{ contact['Address'] }}** 
Primary Phone Number: **{{ contact['Phone'] }}**; Email Address: **{{ contact['Email'] }}**
{% endfor %}

Other Trust Types
There are many types of trusts that can be created in addition to your standard living trust to accomplish your goals. The following is a brief description of some of the more popular trusts. Please read each of them and let me know when we meet if would like to explore adding any of these trust to your estate plan.
[... Keep all of your original Trust Type descriptions here ...]

Advanced Health Care Directive (also known as Living Will or Health Care Power of Attorney)
Do you have any objection to using pain medication if you have pain: Client {{ cb_c_pn_y }} Yes {{ cb_c_pn_n }} No; Spouse {{ cb_s_pn_y }} Yes {{ cb_s_pn_n }} No 
If I am found unconscious, I wish to be resuscitated: Client {{ cb_c_res_y }} Yes {{ cb_c_res_n }} No; Spouse {{ cb_s_res_y }} Yes {{ cb_s_res_n }} No 
Where do you prefer to die? Client {{ cb_c_d_h }} Home {{ cb_c_d_f }} Facility; Spouse {{ cb_s_d_h }} Home {{ cb_s_d_f }} Facility 
In a hospital, does a specialist have to be consulted before treatment: Client {{ cb_c_sp_y }} Yes {{ cb_c_sp_n }} No; Spouse {{ cb_s_sp_y }} Yes {{ cb_s_sp_n }} No 
If yes, name & contact info: **{{ hc_spec_info }}** 
Do you want to be: Client {{ cb_c_bur }} Buried {{ cb_c_crem }} Cremated; Spouse {{ cb_s_bur }} Buried {{ cb_s_crem }} Cremated 
If prepaid, list location, contact info, & contract #: **{{ hc_prepaid }}** 
If NO prepaid, what happens to your remains: **{{ hc_no_prepaid }}** 
Do you want to donate organs: Client {{ cb_c_o_all }} All {{ cb_c_o_non }} None {{ cb_c_o_cer }} Certain; Spouse {{ cb_s_o_all }} All {{ cb_s_o_non }} None {{ cb_s_o_cer }} Certain 
If only certain ones, list which ones: **{{ hc_organs_list }}** 
For what purposes: Client {{ cb_c_op_t }} Transplant {{ cb_c_op_r }} Research {{ cb_c_op_s }} Stem Cell/Cloning {{ cb_c_op_e }} Education 
For what purposes: Spouse {{ cb_s_op_t }} Transplant {{ cb_s_op_r }} Research {{ cb_s_op_s }} Stem Cell/Cloning {{ cb_s_op_e }} Education 
Would you consider assisted suicide if you were terminally ill: Client {{ cb_c_sui_y }} Yes {{ cb_c_sui_n }} No; Spouse {{ cb_s_sui_y }} Yes {{ cb_s_sui_n }} No 
Do you like being outdoors (verses being inside all the time): Client {{ cb_c_out_y }} Yes {{ cb_c_out_n }} No; Spouse {{ cb_s_out_y }} Yes {{ cb_s_out_n }} No 
Do you like spending time with others (verses being alone): Client {{ cb_c_oth_y }} Yes {{ cb_c_oth_n }} No; Spouse {{ cb_s_oth_y }} Yes {{ cb_s_oth_n }} No 
List any allergies (medications/anything found in a hospital): **{{ hc_allergies }}** 
Pick one statement that best describes your idea on prolonging life (life support): 
I do not want any form of life support for any reason: Client {{ cb_ls1_c }} Spouse {{ cb_ls1_s }} 
Prolong my life as long as possible within the limits of accepted medical standards: Client {{ cb_ls2_c }} Spouse {{ cb_ls2_s }} 
Prolong my life only if doctors think I will recover relatively quick and complete: Client {{ cb_ls3_c }} Spouse {{ cb_ls3_s }}

Client’s Agents Who Can Make Healthcare Decisions:
{% for agent in c_hc_agents %}
Agent #{{ loop.index }} Full Name: **{{ agent['Name'] }}** 
{% endfor %}
If you listed multiple agents, how should they act: {{ cb_c_hca_sep }} Separate {{ cb_c_hca_jt }} Jointly; If jointly, then by: {{ cb_c_hca_maj }} Majority {{ cb_c_hca_ind }} Individual {{ cb_c_hca_un }} Unanimous

Spouse’s Agents Who Can Make Healthcare Decisions:
{% for agent in s_hc_agents %}
Agent #{{ loop.index }} Full Name: **{{ agent['Name'] }}** 
{% endfor %}
If you listed multiple agents, how should they act: {{ cb_s_hca_sep }} Separate {{ cb_s_hca_jt }} Jointly; If jointly, then by: {{ cb_s_hca_maj }} Majority {{ cb_s_hca_ind }} Individual {{ cb_s_hca_un }} Unanimous

Financial Power of Attorney Agents
Client’s Agents Who Can Have Power Over Your Assets When You Are Incapacitated:
{% for agent in c_poa_agents %}
Agent #{{ loop.index }} Full Name: **{{ agent['Name'] }}** 
{% endfor %}
If you listed multiple agents, how should they act: {{ cb_c_poa_sep }} Separate {{ cb_c_poa_jt }} Jointly; If jointly, then by: {{ cb_c_poa_maj }} Majority {{ cb_c_poa_ind }} Individual {{ cb_c_poa_un }} Unanimous

Spouse’s Agents Who Can Have Power Over Your Assets When You Are Incapacitated:
{% for agent in s_poa_agents %}
Agent #{{ loop.index }} Full Name: **{{ agent['Name'] }}** 
{% endfor %}
If you listed multiple agents, how should they act: {{ cb_s_poa_sep }} Separate {{ cb_s_poa_jt }} Jointly; If jointly, then by: {{ cb_s_poa_maj }} Majority {{ cb_s_poa_ind }} Individual {{ cb_s_poa_un }} Unanimous 
When should this take effect: Client {{ cb_c_poaw_n }} Now {{ cb_c_poaw_i }} Upon Incapacity; Spouse {{ cb_s_poaw_n }} Now {{ cb_s_poaw_i }} Upon Incapacity 
Can Agent change Trust and make gifts: Client {{ cb_c_poag_y }} Yes {{ cb_c_poag_n }} No; Spouse {{ cb_s_poag_y }} Yes {{ cb_s_poag_n }} No

Nomination Of Guardians (ONLY FILL IN IF YOU HAVE MINOR CHILDREN)
GOP - Guardians of the Person (who will have physical custody):
{% for agent in gop_agents %}
Agent #{{ loop.index }} Full Name: **{{ agent['Name'] }}** 
{% endfor %}
If you listed multiple guardians, how should they act: {{ cb_gop_sep }} Separate {{ cb_gop_jt }} Jointly; If jointly, then by: {{ cb_gop_maj }} Majority {{ cb_gop_ind }} Individual {{ cb_gop_un }} Unanimous

Guardians of the Estate (who will be in charge of the money – can be the same or different from GOP):
{% for agent in goe_agents %}
Agent #{{ loop.index }} Full Name: **{{ agent['Name'] }}** 
{% endfor %}
If you listed multiple guardians, how should they act: {{ cb_goe_sep }} Separate {{ cb_goe_jt }} Jointly; If jointly, then by: {{ cb_goe_maj }} Majority {{ cb_goe_ind }} Individual {{ cb_goe_un }} Unanimous

Trust Beneficiaries
Trustee & Executor When You Die:
Your trustee is the person that will have to do all of the work when you die. They will be in charge of collecting money owed, paying debts and expenses, and then dividing everything up. Who would you like that to be? 
{% for trustee in trustees %}
Trustee #{{ loop.index }} Full Name: **{{ trustee['Name'] }}** 
{% endfor %}
If you listed multiple trustees, how should they act: {{ cb_tst_sep }} Separate {{ cb_tst_jt }} Jointly; If jointly, then by: {{ cb_tst_maj }} Majority {{ cb_tst_ind }} Individual {{ cb_tst_un }} Unanimous 
Should they be compensated for doing this. If so, how (% of estate, Set $, Hourly Rate, Etc.) **{{ tst_comp }}**

In Trust or Outright:
When you die, you can pass everything to your beneficiaries either in trust or outright. If you choose outright, your beneficiaries will get the assets in their name, and everything is simple. However, if your beneficiaries are sued, go through a divorce, file bankruptcy, or have creditors after them, then the assets could be taken in any of those situations. If instead you choose in trust, then they can still be in charge of the assets (or not – up to you) and as long as the assets stay in the trust they are protected from all of these situations.
Who Will Inherit Your Assets:
{% for ben in beneficiaries %}
Full Legal Name: **{{ ben['Name'] }}**; US Citizen: **{{ ben['US Citizen?'] }}**; % or $ Amount: **{{ ben['% or $ Amount'] }}** 
If he/she Dies, then to: **{{ ben['If Dies, to'] }}** OR **{{ ben['If Other, name'] }}**; What if he/she/they die, then to: **{{ ben['What if they die?'] }}** 
Is he/she receiving government financial benefits/disability: **{{ ben['Gov Benefits?'] }}**; How to Inherit: **{{ ben['Inherit'] }}**
{% endfor %}
Pets:
What do you want to happen to your pets: **{{ pets_instructions }}**

Specific Gifts
{% for gift in gifts %}
Complete Description of the Gift: **{{ gift['Gift'] }}**; Who will receive this gift: **{{ gift['Beneficiary Name'] }}** 
This gift be ditributed at the death of: **{{ gift['Death of (H/W/Both)'] }}**; Should this gift go: **{{ gift['Distribute'] }}**
If this person is unable to receive this gift, then: **{{ gift['If unable'] }}** - **{{ gift['Else Name'] }}**
{% endfor %}

Disinheritance
If a scenario plays out where all of your listed beneficiaries die before or at the same time as you, then we need a catch all provision to make sure the assets do not go to the state. Do you prefer to keep the money in the family, or give it to charity: Client {{ cb_cd_fam }} Family {{ cb_cd_char }} Charity; Spouse {{ cb_sd_fam }} Family {{ cb_sd_char }} Charity 
If Family, is there anyone at all in your blood line that you want to make sure never gets any of your assets? 
{% for dis in disinherited %}
Full Legal Name: **{{ dis['Name'] }}**; Relation to You: **{{ dis['Relation'] }}**; Why do you not wish them to get any assets: **{{ dis['Reason'] }}** 
{% endfor %}
If Chairty, what charity or cause: **{{ char_cause }}**

General
Referral Source: **{{ ref_src }}**; Trust Protectors: Appointed by CPA or Attorney. 
Beneficiary of SST at age: **{{ sst_age }}**; Power to Amend after 1st Death: {{ cb_am1_y }} Yes {{ cb_am1_n }} No; Cost: **{{ cost_fee }}** 
Specific Religion Planning: **{{ rel_plan }}**; Consultation Date: **{{ cons_date }}** 
Signing Date/City: **{{ sign_dc }}**; Trust Type: **{{ trust_type }}** 
Do you want the beneficiaries to be able to buy out the other beneficiaries: {{ cb_buy_y }} Yes {{ cb_buy_n }} No; 
Documents to be updated / Services to be provided: 
{{ cb_doc_t }} Trust 
{{ cb_doc_w }} Will(s) 
{{ cb_doc_a }} AHCD(s) 
{{ cb_doc_p }} POA(s) 
{{ cb_doc_n }} NOG 
{{ cb_doc_ub }} Update Existing Binder 
{{ cb_doc_nb }} Client Pay for New Binder 
{{ cb_doc_h }} Homestead Declaration 
{{ cb_doc_hd }} Home Deed 
{{ cb_doc_nhd }} Non-Home Deeds 
{{ cb_doc_tax }} Tax Planning Services 
{{ cb_doc_biz }} Business Law Services 
{{ cb_doc_ap }} Asset Protection Planning 
{{ cb_doc_oth }} Other
Other Notes
**{{ other_notes }}**
