import streamlit as st
import os
import smtplib
from email.message import EmailMessage
from pypdf import PdfReader, PdfWriter
from datetime import datetime

# Load email credentials from Streamlit Secrets
SENDER_EMAIL = st.secrets["sender_email"]
SENDER_PASSWORD = st.secrets["sender_password"]
RECEIVER_EMAIL = "phil@barthattorneys.com"

def send_email_with_pdf(pdf_path, filename):
    msg = EmailMessage()
    msg['Subject'] = f"New EP Questionnaire: {filename}"
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg.set_content("A new client intake questionnaire has been completed. The generated PDF is attached.")

    with open(pdf_path, 'rb') as f:
        pdf_data = f.read()
        
    msg.add_attachment(pdf_data, maintype='application', subtype='pdf', filename=filename)

    # Connect to Gmail's SMTP server
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)

def fill_ep_questionnaire(input_pdf_path, output_pdf_path, data_dict):
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_dict)
    with open(output_pdf_path, "wb") as output_stream:
        writer.write(output_stream)

# --- Streamlit Web Interface ---
st.title("BarthCalderon Estate Planning Questionnaire")

with st.form("ep_intake_form"):
    client_name = st.text_input("Full Legal Name")
    # Add additional intake fields here as needed
    
    submitted = st.form_submit_button("Submit Questionnaire")
    
    if submitted:
        # Map inputs to your PDF's field keys
        form_data = {
            "REPLACE_WITH_NAME_KEY": client_name
        }
        
        input_template = "EP Q.pdf"
        safe_client_name = client_name.replace(" ", "_") if client_name else "Client"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_filename = f"{safe_client_name}_EP_Questionnaire_{timestamp}.pdf"
        
        # Save temporarily on the cloud server
        temp_path = f"/tmp/{output_filename}"
        
        try:
            # 1. Generate the completed PDF
            fill_ep_questionnaire(input_template, temp_path, form_data)
            
            # 2. Email the PDF via Gmail SMTP
            send_email_with_pdf(temp_path, output_filename)
            
            # 3. Remove the temporary file from the cloud server
            os.remove(temp_path)
            
            st.success("Success! Your questionnaire has been securely submitted to the attorney.")
            
        except Exception as e:
            st.error(f"An error occurred during submission: {e}")
