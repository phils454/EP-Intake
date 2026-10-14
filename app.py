import streamlit as st
import os
import smtplib
from email.message import EmailMessage
from docxtpl import DocxTemplate
from datetime import datetime
import json
import io

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
    "contacts_data": [{"Name": "", "DOB": "", "Relation": "", "Related to (Client/Spouse/Both)": "", "Address": "", "Phone": "", "Email": ""}]
}

for k, v in default_keys.items():
    if k not in st.session_state:
        st.session_state[k] = v

def cb(condition):
    return "☑" if condition else "☐"

def safe_index(lst, val):
    return lst.index(val) if val in lst else 0

# --- Main App Interface ---
st.set_page_config(page_title="BarthCalderon Intake Form", layout="wide")

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
    st.session_state['c_gen'] = c8.radio("Gender", ["Male", "Female"], index=safe_index(["Male", "Female"], st.session_state['c_gen']))
    st.session_state['c_cit'] = c9.radio("US Citizen?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['c_cit']))
    st.session_state['c_wid'] = c10.radio("Widowed/Divorced?", ["No", "Yes"], index=safe_index(["No", "Yes"], st.session_state['c_wid']))
    st.session_state['c_63'] = c11.radio("Over 63?", ["No", "Yes"], index=safe_index(["No", "Yes"], st.session_state['c_63']))
    
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
    st.session_state['s_gen'] = s8.radio("Spouse Gender", ["Male", "Female"], index=safe_index(["Male", "Female"], st.session_state['s_gen']))
    st.session_state['s_cit'] = s9.radio("Spouse US Citizen?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['s_cit']))
    st.session_state['s_wid'] = s10.radio("Spouse Widowed/Divorced?", ["No", "Yes"], index=safe_index(["No", "Yes"], st.session_state['s_wid']))
    st.session_state['s_63'] = s11.radio("Spouse Over 63?", ["No", "Yes"], index=safe_index(["No", "Yes"], st.session_state['s_63']))
    
    st.session_state['s_former_spouses'] = st.text_input("Spouse Former Spouses & Dates", value=st.session_state['s_former_spouses'])
    
    s12, s13 = st.columns(2)
    st.session_state['s_income'] = s12.text_input("Spouse Approx. Income", value=st.session_state['s_income'])
    st.session_state['s_obligations'] = s13.text_input("Spouse Court Obligations", value=st.session_state['s_obligations'])

with t_assets:
    st.subheader("Trust History")
    ta1, ta2 = st.columns(2)
    st.session_state['prior_trust'] = ta1.radio("Do you already have a trust?", ["No", "Yes"], index=safe_index(["No", "Yes"], st.session_state['prior_trust']))
    st.session_state['t_amend_num'] = ta2.text_input("Amendment/Restatement #", value=st.session_state['t_amend_num'])
    st.session_state['t_desired_name'] = st.text_input("Desired Name of Trust", value=st.session_state['t_desired_name'])
    st.session_state['t_reason'] = st.text_area("Why do you want a trust?", value=st.session_state['t_reason'])
    
    st.markdown("---")
    st.subheader("Financial Accounts")
    ta3, ta4 = st.columns(2)
    st.session_state['li_rev'] = ta3.radio("Have you had your life insurance reviewed recently?", ["Yes", "No", "Not Applicable"], index=safe_index(["Yes", "No", "Not Applicable"], st.session_state['li_rev']))
    st.session_state['inv_rev'] = ta4.radio("Do you want a recommendation for a financial advisor?", ["Yes", "No", "Not Applicable"], index=safe_index(["Yes", "No", "Not Applicable"], st.session_state['inv_rev']))
    
    st.session_state["accounts_data"] = st.data_editor(st.session_state["accounts_data"], num_rows="dynamic", use_container_width=True, key="accounts_editor")
            
    st.markdown("---")
    st.subheader("Businesses")
    st.session_state["businesses_data"] = st.data_editor(st.session_state["businesses_data"], num_rows="dynamic", use_container_width=True, key="biz_editor")

    st.markdown("---")
    st.subheader("Money Owed to You")
    st.session_state["owed_data"] = st.data_editor(st.session_state["owed_data"], num_rows="dynamic", use_container_width=True, key="owed_editor")

    st.markdown("---")
    st.subheader("Real Estate")
    st.session_state['hs_decl'] = st.radio("Do you have a homestead declaration filed for your primary home?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['hs_decl']))
    
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
    st.session_state['inh'] = st.radio("Will you inherit any assets from anyone else in a trust?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['inh']))
    st.session_state['ip'] = st.radio("Do you own any intellectual property?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['ip']))
    st.session_state['oth_assets'] = st.text_area("Other assets", value=st.session_state['oth_assets'])

with t_health:
    st.subheader("Health Directives")
    h1, h2 = st.columns(2)
    st.session_state['c_pn'] = h1.radio("Client do you have any objection to pain meds?", ["No", "Yes"], index=safe_index(["No", "Yes"], st.session_state['c_pn']))
    st.session_state['s_pn'] = h2.radio("Spouse do you have any objection to pain meds?", ["No", "Yes"], index=safe_index(["No", "Yes"], st.session_state['s_pn']))
    st.session_state['c_res'] = h1.radio("Client if you are found unresponsive, do you want EMT's to try to resuscitate you?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['c_res']))
    st.session_state['s_res'] = h2.radio("Spouse if you are found unresponsive, do you want EMT's to try to resuscitate you?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['s_res']))
    st.session_state['c_d'] = h1.radio("Client where do you prefer to die?", ["Home", "Facility"], index=safe_index(["Home", "Facility"], st.session_state['c_d']))
    st.session_state['s_d'] = h2.radio("Spouse where do you prefer to die?", ["Home", "Facility"], index=safe_index(["Home", "Facility"], st.session_state['s_d']))
    st.session_state['c_sp'] = h1.radio("Client can any Dr. treat you in an emergency, or do they need to consult a specialist first?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['c_sp']))
    st.session_state['s_sp'] = h2.radio("Spouse can any Dr. treat you in an emergency, or do they need to consult a specialist first?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['s_sp']))
    
    st.session_state['hc_spec_info'] = st.text_input("If you marked that a specialist needs to be consulted, please provide his/her name, phone, email, address:", value=st.session_state['hc_spec_info'])
    
    h3, h4 = st.columns(2)
    st.session_state['c_bur'] = h3.radio("Client do you want to be:", ["Buried", "Cremated"], index=safe_index(["Buried", "Cremated"], st.session_state['c_bur']))
    st.session_state['s_bur'] = h4.radio("Spouse do you want to be:", ["Buried", "Cremated"], index=safe_index(["Buried", "Cremated"], st.session_state['s_bur']))
    
    st.session_state['hc_prepaid'] = st.text_input("If you have a prepaid funeral or dispostion, please provide the contract number, organization, and contact info:", value=st.session_state['hc_prepaid'])
    st.session_state['hc_no_prepaid'] = st.text_input("If you do not have prepaid arrangements, how do you want your remains to be disposed of:", value=st.session_state['hc_no_prepaid'])
    
    h5, h6 = st.columns(2)
    st.session_state['c_org'] = h5.radio("Client do you want to donate organs?", ["All", "None", "Certain"], index=safe_index(["All", "None", "Certain"], st.session_state['c_org']))
    st.session_state['s_org'] = h6.radio("Spouse do you want to donate organs?", ["All", "None", "Certain"], index=safe_index(["All", "None", "Certain"], st.session_state['s_org']))
    
    h7, h8 = st.columns(2)
    st.session_state['c_op'] = h7.radio("Client for what purpose", ["Transplant", "Research", "Stem Cell/Cloning", "Education"], index=safe_index(["Transplant", "Research", "Stem Cell/Cloning", "Education"], st.session_state['c_op']))
    st.session_state['s_op'] = h8.radio("Spouse for what purpose", ["Transplant", "Research", "Stem Cell/Cloning", "Education"], index=safe_index(["Transplant", "Research", "Stem Cell/Cloning", "Education"], st.session_state['s_op']))
    
    st.session_state['hc_organs_list'] = st.text_input("If you have any restrictions on organ donation, please provide details here:", value=st.session_state['hc_organs_list'])
    
    h7a, h8a = st.columns(2)
    st.session_state['c_sui'] = h7a.radio("Client if legal, are you open to assisted suicide?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['c_sui']))
    st.session_state['s_sui'] = h8a.radio("Spouse if legal, are you open to assisted suicide?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['s_sui']))
    st.session_state['c_out'] = h7a.radio("Client do you want to make sure you spend time outdoors?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['c_out']))
    st.session_state['s_out'] = h8a.radio("Spouse do you want to make sure you spend time outdoors?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['s_out']))
    st.session_state['c_oth'] = h7a.radio("Client do you want to make sure you have social interaction?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['c_oth']))
    st.session_state['s_oth'] = h8a.radio("Spouse do you want to make sure you have social interaction?", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['s_oth']))
    
    st.session_state['hc_allergies'] = st.text_input("Are you allergic to medications or anything found in a Dr.'s office or hospital:", value=st.session_state['hc_allergies'])
    
    st.write("")
    h9, h10 = st.columns(2)
    st.session_state['ls_c'] = h9.radio("Client Life Support Choice", ["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], index=safe_index(["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], st.session_state['ls_c']))
    st.session_state['ls_s'] = h10.radio("Spouse Life Support Choice", ["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], index=safe_index(["No life support for any reason", "Life support for as long as legally possible", "Life support only if it results in relatively full recovery in a reasonably short time"], st.session_state['ls_s']))

    st.markdown("---")
    st.subheader("Healthcare Agents")
    st.write("Client HC Agents")
    st.session_state["c_hc_agents_data"] = st.data_editor(st.session_state["c_hc_agents_data"], num_rows="dynamic", use_container_width=True, key="chc_editor")
    c_h1, c_h2 = st.columns(2)
    st.session_state['c_hca_act'] = c_h1.selectbox("Client how do you want the above HC agents to act", ["Separate", "Jointly"], index=safe_index(["Separate", "Jointly"], st.session_state['c_hca_act']))
    st.session_state['chca_jt'] = c_h2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=safe_index(["Majority", "Individual", "Unanimous"], st.session_state['chca_jt']), key="c_hca_jt_val")

    st.write("Spouse HC Agents")
    st.session_state["s_hc_agents_data"] = st.data_editor(st.session_state["s_hc_agents_data"], num_rows="dynamic", use_container_width=True, key="shc_editor")
    s_h1, s_h2 = st.columns(2)
    st.session_state['s_hca_act'] = s_h1.selectbox("Spouse how do you want the above HC agents to act", ["Separate", "Jointly"], index=safe_index(["Separate", "Jointly"], st.session_state['s_hca_act']))
    st.session_state['shca_jt'] = s_h2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=safe_index(["Majority", "Individual", "Unanimous"], st.session_state['shca_jt']), key="s_hca_jt_val")
    
    st.markdown("---")
    st.subheader("Financial POA Agents")
    st.write("Client POA Agents")
    st.session_state["c_poa_agents_data"] = st.data_editor(st.session_state["c_poa_agents_data"], num_rows="dynamic", use_container_width=True, key="cpoa_editor")
    c_p1, c_p2, c_p3 = st.columns(3)
    st.session_state['c_poa_act'] = c_p1.selectbox("Client how do you want the above POA agents to act", ["Separate", "Jointly"], index=safe_index(["Separate", "Jointly"], st.session_state['c_poa_act']))
    st.session_state['cpoa_jt'] = c_p2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=safe_index(["Majority", "Individual", "Unanimous"], st.session_state['cpoa_jt']), key="c_poa_jt_val")
    st.session_state['c_poaw'] = c_p3.radio("Client when do you want the power to take effect", ["Now", "Upon Incapacity"], index=safe_index(["Now", "Upon Incapacity"], st.session_state['c_poaw']))
    st.session_state['c_poag'] = st.radio("Client Gifts", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['c_poag']))

    st.write("Spouse POA Agents")
    st.session_state["s_poa_agents_data"] = st.data_editor(st.session_state["s_poa_agents_data"], num_rows="dynamic", use_container_width=True, key="spoa_editor")
    s_p1, s_p2, s_p3 = st.columns(3)
    st.session_state['s_poa_act'] = s_p1.selectbox("Spouse how do you want the above POA agents to act", ["Separate", "Jointly"], index=safe_index(["Separate", "Jointly"], st.session_state['s_poa_act']))
    st.session_state['spoa_jt'] = s_p2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=safe_index(["Majority", "Individual", "Unanimous"], st.session_state['spoa_jt']), key="s_poa_jt_val")
    st.session_state['s_poaw'] = s_p3.radio("Spouse when do you want the power to take effect", ["Now", "Upon Incapacity"], index=safe_index(["Now", "Upon Incapacity"], st.session_state['s_poaw']))
    st.session_state['s_poag'] = st.radio("Spouse Gifts", ["Yes", "No"], index=safe_index(["Yes", "No"], st.session_state['s_poag']))

with t_fiduciaries:
    st.subheader("Guardians of the Person")
    st.session_state["gop_data"] = st.data_editor(st.session_state["gop_data"], num_rows="dynamic", use_container_width=True, key="gop_editor")
    gop_1, gop_2 = st.columns(2)
    st.session_state['gop_act'] = gop_1.selectbox("GOP Act", ["Separate", "Jointly"], index=safe_index(["Separate", "Jointly"], st.session_state['gop_act']))
    st.session_state['gopjt'] = gop_2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=safe_index(["Majority", "Individual", "Unanimous"], st.session_state['gopjt']), key="gop_jt_val")

    st.subheader("Guardians of the Estate")
    st.session_state["goe_data"] = st.data_editor(st.session_state["goe_data"], num_rows="dynamic", use_container_width=True, key="goe_editor")
    goe_1, goe_2 = st.columns(2)
    st.session_state['goe_act'] = goe_1.selectbox("GOE Act", ["Separate", "Jointly"], index=safe_index(["Separate", "Jointly"], st.session_state['goe_act']))
    st.session_state['goejt'] = goe_2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=safe_index(["Majority", "Individual", "Unanimous"], st.session_state['goejt']), key="goe_jt_val")

    st.subheader("Trustees")
    st.session_state["trustees_data"] = st.data_editor(st.session_state["trustees_data"], num_rows="dynamic", use_container_width=True, key="tst_editor")
    tst_1, tst_2 = st.columns(2)
    st.session_state['tst_act'] = tst_1.selectbox("How should the above trustees act?", ["Separate", "Jointly"], index=safe_index(["Separate", "Jointly"], st.session_state['tst_act']))
    st.session_state['tstjt'] = tst_2.selectbox("If Jointly", ["Majority", "Individual", "Unanimous"], index=safe_index(["Majority", "Individual", "Unanimous"], st.session_state['tstjt']), key="tst_jt_val")
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
    st.subheader("Contacts (Please list all of the following information for everyone you have named in this form (other than yourselves)):")
    st.session_state["contacts_data"] = st.data_editor(st.session_state["contacts_data"], num_rows="dynamic", use_container_width=True, key="contacts_editor")

    st.subheader("General Information")
    st.session_state['ref_src'] = st.text_input("Referral Source", value=st.session_state['ref_src'])

    # --- SUBMIT BUTTON ---
    st.markdown("---")
    st.subheader("Submit Questionnaire")
    
    submit = st.button("Generate Documents & Email Results", type="primary")

    if submit:
        with st.spinner("Generating document and emailing securely..."):
            ctx = {}
            
            # General string mappings straight from session_state to avoid NameError
            text_keys = [
                "c_id", "c_address", "c_county", "c_mailing", "c_phone", "c_email", "c_dob", "c_pob",
                "c_former_spouses", "c_income", "c_obligations", "s_id", "s_address", "s_county",
                "s_mailing", "s_phone", "s_email", "s_dob", "s_pob", "s_former_spouses", "s_income",
                "s_obligations", "t_amend_num", "t_desired_name", "t_reason", "re_pri_addr", "re_pri_val",
                "re_pri_cnty", "re_pri_move", "re_pri_eq", "re_pri_basis", "re_pri_gain", "re_sec_addr",
                "re_sec_val", "re_sec_cnty", "re_sec_move", "re_sec_eq", "re_sec_basis", "re_sec_gain",
                "oth_assets", "hc_spec_info", "hc_prepaid", "hc_no_prepaid", "hc_organs_list", "hc_allergies",
                "pets_instructions", "char_cause", "ref_src", "other_notes", "tst_comp",
                "c_hca_act", "chca_jt", "s_hca_act", "shca_jt", "c_poa_act", "cpoa_jt", "s_poa_act", "spoa_jt",
                "gop_act", "gopjt", "goe_act", "goejt", "tst_act", "tstjt"
            ]
            for k in text_keys:
                ctx[k] = st.session_state.get(k, "")

            # Logic Checkboxes / Radios
            ctx['cb_c_male'] = cb(st.session_state.get('c_gen') == "Male"); ctx['cb_c_female'] = cb(st.session_state.get('c_gen') == "Female")
            ctx['cb_c_cit_y'] = cb(st.session_state.get('c_cit') == "Yes"); ctx['cb_c_cit_n'] = cb(st.session_state.get('c_cit') == "No")
            ctx['cb_c_wid_y'] = cb(st.session_state.get('c_wid') == "Yes"); ctx['cb_c_wid_n'] = cb(st.session_state.get('c_wid') == "No")
            ctx['cb_c_63_y'] = cb(st.session_state.get('c_63') == "Yes"); ctx['cb_c_63_n'] = cb(st.session_state.get('c_63') == "No")
            ctx['cb_s_male'] = cb(st.session_state.get('s_gen') == "Male"); ctx['cb_s_female'] = cb(st.session_state.get('s_gen') == "Female")
            ctx['cb_s_cit_y'] = cb(st.session_state.get('s_cit') == "Yes"); ctx['cb_s_cit_n'] = cb(st.session_state.get('s_cit') == "No")
            ctx['cb_s_wid_y'] = cb(st.session_state.get('s_wid') == "Yes"); ctx['cb_s_wid_n'] = cb(st.session_state.get('s_wid') == "No")
            ctx['cb_s_63_y'] = cb(st.session_state.get('s_63') == "Yes"); ctx['cb_s_63_n'] = cb(st.session_state.get('s_63') == "No")
            ctx['cb_prior_trust_y'] = cb(st.session_state.get('prior_trust') == "Yes"); ctx['cb_prior_trust_n'] = cb(st.session_state.get('prior_trust') == "No")
            ctx['cb_li_rev_y'] = cb(st.session_state.get('li_rev') == "Yes"); ctx['cb_li_rev_n'] = cb(st.session_state.get('li_rev') in ["No", "Not Applicable"])
            ctx['cb_inv_rev_y'] = cb(st.session_state.get('inv_rev') == "Yes"); ctx['cb_inv_rev_n'] = cb(st.session_state.get('inv_rev') in ["No", "Not Applicable"])
            ctx['cb_hs_y'] = cb(st.session_state.get('hs_decl') == "Yes"); ctx['cb_hs_n'] = cb(st.session_state.get('hs_decl') == "No")
            ctx['cb_inh_y'] = cb(st.session_state.get('inh') == "Yes"); ctx['cb_inh_n'] = cb(st.session_state.get('inh') == "No")
            ctx['cb_ip_y'] = cb(st.session_state.get('ip') == "Yes"); ctx['cb_ip_n'] = cb(st.session_state.get('ip') == "No")
            ctx['cb_c_pn_y'] = cb(st.session_state.get('c_pn') == "Yes"); ctx['cb_c_pn_n'] = cb(st.session_state.get('c_pn') == "No")
            ctx['cb_s_pn_y'] = cb(st.session_state.get('s_pn') == "Yes"); ctx['cb_s_pn_n'] = cb(st.session_state.get('s_pn') == "No")
            ctx['cb_c_res_y'] = cb(st.session_state.get('c_res') == "Yes"); ctx['cb_c_res_n'] = cb(st.session_state.get('c_res') == "No")
            ctx['cb_s_res_y'] = cb(st.session_state.get('s_res') == "Yes"); ctx['cb_s_res_n'] = cb(st.session_state.get('s_res') == "No")
            ctx['cb_c_d_h'] = cb(st.session_state.get('c_d') == "Home"); ctx['cb_c_d_f'] = cb(st.session_state.get('c_d') == "Facility")
            ctx['cb_s_d_h'] = cb(st.session_state.get('s_d') == "Home"); ctx['cb_s_d_f'] = cb(st.session_state.get('s_d') == "Facility")
            ctx['cb_c_sp_y'] = cb(st.session_state.get('c_sp') == "Yes"); ctx['cb_c_sp_n'] = cb(st.session_state.get('c_sp') == "No")
            ctx['cb_s_sp_y'] = cb(st.session_state.get('s_sp') == "Yes"); ctx['cb_s_sp_n'] = cb(st.session_state.get('s_sp') == "No")
            ctx['cb_c_bur'] = cb(st.session_state.get('c_bur') == "Buried"); ctx['cb_c_crem'] = cb(st.session_state.get('c_bur') == "Cremated")
            ctx['cb_s_bur'] = cb(st.session_state.get('s_bur') == "Buried"); ctx['cb_s_crem'] = cb(st.session_state.get('s_bur') == "Cremated")
            ctx['cb_c_o_all'] = cb(st.session_state.get('c_org') == "All"); ctx['cb_c_o_non'] = cb(st.session_state.get('c_org') == "None"); ctx['cb_c_o_cer'] = cb(st.session_state.get('c_org') == "Certain")
            ctx['cb_s_o_all'] = cb(st.session_state.get('s_org') == "All"); ctx['cb_s_o_non'] = cb(st.session_state.get('s_org') == "None"); ctx['cb_s_o_cer'] = cb(st.session_state.get('s_org') == "Certain")
            ctx['cb_c_op_t'] = cb(st.session_state.get('c_op') == "Transplant"); ctx['cb_c_op_r'] = cb(st.session_state.get('c_op') == "Research"); ctx['cb_c_op_s'] = cb(st.session_state.get('c_op') == "Stem Cell/Cloning"); ctx['cb_c_op_e'] = cb(st.session_state.get('c_op') == "Education")
            ctx['cb_s_op_t'] = cb(st.session_state.get('s_op') == "Transplant"); ctx['cb_s_op_r'] = cb(st.session_state.get('s_op') == "Research"); ctx['cb_s_op_s'] = cb(st.session_state.get('s_op') == "Stem Cell/Cloning"); ctx['cb_s_op_e'] = cb(st.session_state.get('s_op') == "Education")
            ctx['cb_c_sui_y'] = cb(st.session_state.get('c_sui') == "Yes"); ctx['cb_c_sui_n'] = cb(st.session_state.get('c_sui') == "No")
            ctx['cb_s_sui_y'] = cb(st.session_state.get('s_sui') == "Yes"); ctx['cb_s_sui_n'] = cb(st.session_state.get('s_sui') == "No")
            ctx['cb_c_out_y'] = cb(st.session_state.get('c_out') == "Yes"); ctx['cb_c_out_n'] = cb(st.session_state.get('c_out') == "No")
            ctx['cb_s_out_y'] = cb(st.session_state.get('s_out') == "Yes"); ctx['cb_s_out_n'] = cb(st.session_state.get('s_out') == "No")
            ctx['cb_c_oth_y'] = cb(st.session_state.get('c_oth') == "Yes"); ctx['cb_c_oth_n'] = cb(st.session_state.get('c_oth') == "No")
            ctx['cb_s_oth_y'] = cb(st.session_state.get('s_oth') == "Yes"); ctx['cb_s_oth_n'] = cb(st.session_state.get('s_oth') == "No")
            ctx['cb_ls1_c'] = cb(st.session_state.get('ls_c') == "No life support for any reason"); ctx['cb_ls2_c'] = cb(st.session_state.get('ls_c') == "Life support for as long as legally possible"); ctx['cb_ls3_c'] = cb(st.session_state.get('ls_c') == "Life support only if it results in relatively full recovery in a reasonably short time")
            ctx['cb_ls1_s'] = cb(st.session_state.get('ls_s') == "No life support for any reason"); ctx['cb_ls2_s'] = cb(st.session_state.get('ls_s') == "Life support for as long as legally possible"); ctx['cb_ls3_s'] = cb(st.session_state.get('ls_s') == "Life support only if it results in relatively full recovery in a reasonably short time")
            
            # Fiduciary Checkboxes mapped properly
            ctx['cb_c_hca_sep'] = cb(st.session_state.get('c_hca_act') == "Separate"); ctx['cb_c_hca_jt'] = cb(st.session_state.get('c_hca_act') == "Jointly")
            ctx['cb_c_hca_maj'] = cb(st.session_state.get('chca_jt') == "Majority"); ctx['cb_c_hca_ind'] = cb(st.session_state.get('chca_jt') == "Individual"); ctx['cb_c_hca_un'] = cb(st.session_state.get('chca_jt') == "Unanimous")
            
            ctx['cb_s_hca_sep'] = cb(st.session_state.get('s_hca_act') == "Separate"); ctx['cb_s_hca_jt'] = cb(st.session_state.get('s_hca_act') == "Jointly")
            ctx['cb_s_hca_maj'] = cb(st.session_state.get('shca_jt') == "Majority"); ctx['cb_s_hca_ind'] = cb(st.session_state.get('shca_jt') == "Individual"); ctx['cb_s_hca_un'] = cb(st.session_state.get('shca_jt') == "Unanimous")

            ctx['cb_c_poa_sep'] = cb(st.session_state.get('c_poa_act') == "Separate"); ctx['cb_c_poa_jt'] = cb(st.session_state.get('c_poa_act') == "Jointly")
            ctx['cb_c_poa_maj'] = cb(st.session_state.get('cpoa_jt') == "Majority"); ctx['cb_c_poa_ind'] = cb(st.session_state.get('cpoa_jt') == "Individual"); ctx['cb_c_poa_un'] = cb(st.session_state.get('cpoa_jt') == "Unanimous")

            ctx['cb_s_poa_sep'] = cb(st.session_state.get('s_poa_act') == "Separate"); ctx['cb_s_poa_jt'] = cb(st.session_state.get('s_poa_act') == "Jointly")
            ctx['cb_s_poa_maj'] = cb(st.session_state.get('spoa_jt') == "Majority"); ctx['cb_s_poa_ind'] = cb(st.session_state.get('spoa_jt') == "Individual"); ctx['cb_s_poa_un'] = cb(st.session_state.get('spoa_jt') == "Unanimous")

            ctx['cb_gop_sep'] = cb(st.session_state.get('gop_act') == "Separate"); ctx['cb_gop_jt'] = cb(st.session_state.get('gop_act') == "Jointly")
            ctx['cb_gop_maj'] = cb(st.session_state.get('gopjt') == "Majority"); ctx['cb_gop_ind'] = cb(st.session_state.get('gopjt') == "Individual"); ctx['cb_gop_un'] = cb(st.session_state.get('gopjt') == "Unanimous")

            ctx['cb_goe_sep'] = cb(st.session_state.get('goe_act') == "Separate"); ctx['cb_goe_jt'] = cb(st.session_state.get('goe_act') == "Jointly")
            ctx['cb_goe_maj'] = cb(st.session_state.get('goejt') == "Majority"); ctx['cb_goe_ind'] = cb(st.session_state.get('goejt') == "Individual"); ctx['cb_goe_un'] = cb(st.session_state.get('goejt') == "Unanimous")

            ctx['cb_tst_sep'] = cb(st.session_state.get('tst_act') == "Separate"); ctx['cb_tst_jt'] = cb(st.session_state.get('tst_act') == "Jointly")
            ctx['cb_tst_maj'] = cb(st.session_state.get('tstjt') == "Majority"); ctx['cb_tst_ind'] = cb(st.session_state.get('tstjt') == "Individual"); ctx['cb_tst_un'] = cb(st.session_state.get('tstjt') == "Unanimous")

            ctx['cb_c_poaw_n'] = cb(st.session_state.get('c_poaw') == "Now"); ctx['cb_c_poaw_i'] = cb(st.session_state.get('c_poaw') == "Upon Incapacity")
            ctx['cb_c_poag_y'] = cb(st.session_state.get('c_poag') == "Yes"); ctx['cb_c_poag_n'] = cb(st.session_state.get('c_poag') == "No")
            ctx['cb_s_poaw_n'] = cb(st.session_state.get('s_poaw') == "Now"); ctx['cb_s_poaw_i'] = cb(st.session_state.get('s_poaw') == "Upon Incapacity")
            ctx['cb_s_poag_y'] = cb(st.session_state.get('s_poag') == "Yes"); ctx['cb_s_poag_n'] = cb(st.session_state.get('s_poag') == "No")
            
            ctx['cb_cd_fam'] = cb("Family" in st.session_state.get('cd_fail', [])); ctx['cb_cd_char'] = cb("Charity" in st.session_state.get('cd_fail', []))
            ctx['cb_sd_fam'] = cb("Family" in st.session_state.get('sd_fail', [])); ctx['cb_sd_char'] = cb("Charity" in st.session_state.get('sd_fail', []))
            
            for doc_key in ["doc_t", "doc_w", "doc_a", "doc_p", "doc_n", "doc_ub", "doc_nb", "doc_h", "doc_hd", "doc_nhd", "doc_tax", "doc_biz", "doc_ap", "doc_oth"]:
                ctx[f'cb_{doc_key}'] = cb(st.session_state.get(doc_key, False))
            
            ctx['cb_am1_y'] = cb(False); ctx['cb_am1_n'] = cb(False); ctx['cb_buy_y'] = cb(False); ctx['cb_buy_n'] = cb(False)
            ctx['sst_age'] = ""; ctx['cost_fee'] = ""; ctx['rel_plan'] = ""; ctx['cons_date'] = ""; ctx['sign_dc'] = ""; ctx['trust_type'] = ""

            # Inject Dynamic Tables
            ctx['accounts'] = st.session_state.get('accounts_data', [])
            ctx['businesses'] = st.session_state.get('businesses_data', [])
            ctx['owed'] = st.session_state.get('owed_data', [])
            ctx['rentals'] = st.session_state.get('rentals_data', [])
            ctx['c_hc_agents'] = st.session_state.get('c_hc_agents_data', [])
            ctx['s_hc_agents'] = st.session_state.get('s_hc_agents_data', [])
            ctx['c_poa_agents'] = st.session_state.get('c_poa_agents_data', [])
            ctx['s_poa_agents'] = st.session_state.get('s_poa_agents_data', [])
            ctx['gop_agents'] = st.session_state.get('gop_data', [])
            ctx['goe_agents'] = st.session_state.get('goe_data', [])
            ctx['trustees'] = st.session_state.get('trustees_data', [])
            ctx['beneficiaries'] = st.session_state.get('beneficiaries_data', [])
            ctx['gifts'] = st.session_state.get('gifts_data', [])
            ctx['disinherited'] = st.session_state.get('disinherited_data', [])
            ctx['contacts'] = st.session_state.get('contacts_data', [])

            # File Processing
            input_template = "EP Q (with Fields) - Rev 7-17-26.docx"
            safe_name = st.session_state.get('c_id', 'Client').replace(" ", "_") if st.session_state.get('c_id') else "Client"
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
