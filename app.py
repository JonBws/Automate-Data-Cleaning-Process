import streamlit as st
import math
import random
import smtplib
import pandas as pd
from email.message import EmailMessage
from email.utils import formataddr
from dotenv import load_dotenv
import os
import warnings
import numpy as np 
import json
from io import BytesIO
import re

if "page" not in st.session_state: st.session_state.page = 0
if "otp" not in st.session_state: st.session_state.otp = None
if "Email" not in st.session_state: st.session_state.Email = None
if "company" not in st.session_state: st.session_state.company = None
if "amount_of_collaborator" not in st.session_state: st.session_state.amount_of_collaborator = None
if "location" not in st.session_state: st.session_state.location = None
if "industry" not in st.session_state: st.session_state.industry = None
if "role" not in st.session_state: st.session_state.role = None

def nextPage(): st.session_state.page += 1

HOST = "smtp-mail.outlook.com"
PORT = 587
load_dotenv(".env")
sender_email = st.secrets["EMAIL"]
password_email = st.secrets["PASSWORD"]

def get_otp():
    digits="0123456789"
    OTP = ""
    for _ in range(6):
        OTP+=digits[math.floor(random.random()*10)]
    return OTP

def is_valid_email(email):
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if re.match(regex, email):
        return True
    else:
        return False

def send_email(receiver_email, subject):
    valid = is_valid_email(receiver_email)
    if valid == True:
        OTP = get_otp()
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = formataddr(("Streamlit App", f"{sender_email}"))
        msg["To"] = receiver_email
        msg["BCC"] = sender_email
        msg.add_alternative(
            f"""\
                <html>
                    <body>
                        <p >Hi <strong>{receiver_email}</strong>,<p>
                        <p>This is your verification code : <strong>{OTP}</strong> </p>
                    </body>
                </html>
            """,
            subtype="html",
        )
        with smtplib.SMTP(HOST, PORT) as server:
            server.starttls()
            server.login(sender_email, password_email)
            server.sendmail(sender_email, receiver_email, msg.as_string())
                
        st.session_state.otp = OTP
    else:
        st.error("Please input valid email")
    
def send_email_after_download():
    msg = EmailMessage()
    msg["Subject"] = "Summary Form"
    msg["From"] = formataddr(("Streamlit App", f"{sender_email}"))
    msg["To"] = st.session_state.Email
    msg["BCC"] = sender_email
    msg.add_alternative(
        f"""\
            <html>
                <body>
                    <p >Hi <strong>{st.session_state.Email}</strong>,<p>
                    <p>This is your Summary Form <strong></p>
                    <p>Email : {st.session_state.Email} </p>
                    <p>Company : {st.session_state.company} </p>
                    <p>Amount Collaborator : {st.session_state.amount_of_collaborator} </p>
                    <p>Location : {st.session_state.location} </p>
                    <p>Type of industry : {st.session_state.industry} </p>
                    <p>Role in The Company: {st.session_state.role} </p>
                    <p>Total Row Number : {st.session_state.total_row} </p>
                </body>
            </html>
        """,
        subtype="html",
    )
    with smtplib.SMTP(HOST, PORT) as server:
        server.starttls()
        server.login(sender_email, password_email)
        server.sendmail(sender_email, st.session_state.Email, msg.as_string())
                
def remove_double_quote(json):
    if json is np.nan:
        return []
    elif json.startswith('"') and json.endswith('"'):
        return str(json[1:-1])

def convert_to_dict(json_string):
    if json_string == np.nan: 
        return []
    elif not isinstance(json_string, str):
        return []
    else:
          return json.loads(json_string)
    
ph = st.empty()

## Verification Page
if st.session_state.page == 0:
    with ph.container():
        st.title("Email Verification 📧")
        st.session_state.Email = st.text_input("Enter Your Email: ")
        if st.button("Sending OTP"):
            send_email(receiver_email=st.session_state.Email, subject="Streamlit Verification Code")
            
        verification_code = st.text_input("Verification OTP")
        
        if verification_code == st.session_state.otp:
            nextPage()
        elif verification_code == "":
            pass
        elif verification_code != st.session_state:
            st.write("Verification unsucessfull, Try Again")
            
            
                
## Form Page              
if st.session_state.page == 1:
    with ph.container():
        disabled_button = True
        st.markdown("## :green[Verification Sucessfull]")
        st.markdown("##### **Please fill the form to continue** ✍🏼")
        st.session_state.company = st.text_input("Company Name")
        st.session_state.amount_of_collaborator = st.selectbox("Ammount Of Collaborator", ["1-5","6-15","16-30","31-60","60+"])
        st.session_state.location = st.text_input("Location")
        st.session_state.industry = st.selectbox("Type of Industry", ["Option 1", "Option 2", "Option 3", "Option 4", "Option 5", "Option 6"])
        st.session_state.role = st.selectbox("Role in The Company", ["Owner", "Employeer", "Editor"])
        if st.session_state.company and  st.session_state.amount_of_collaborator and st.session_state.location and st.session_state.industry and st.session_state.role:
            submitted = st.button("Submit", on_click=nextPage)
        else: 
            submitted = st.button("Submit", disabled=True)
        
if st.session_state.page == 2:
    with ph.container():
        st.title("Upload Your Excel File 📁")
        st.info("Your exel file must contain taxes dissagregated or impuestos desagregados column and taxes amount or monto impuestos column", icon="⚠️")
        uploaded_file = st.file_uploader("Choose an Excel File", type="xlsx")
        if uploaded_file:
            with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    df = pd.read_excel(uploaded_file, engine = "openpyxl")
            if "TAXES_DISAGGREAGTED" and "TAXES_AMOUNT" in df.columns:
                df1 = df[["TAXES_DISAGGREGATED", "TAXES_AMOUNT"]]
                st.session_state.total_row = len(df1)
                df2 = df1.copy()
                df2["TAXES_DISAGGREGATED"] = df1["TAXES_DISAGGREGATED"].apply(remove_double_quote)
                df3 = df2.copy()
                df3["TAXES_DISAGGREGATED"] = df2["TAXES_DISAGGREGATED"].apply(convert_to_dict)
                new_column = []
                for d in df3["TAXES_DISAGGREGATED"]:
                    for item in d:
                        new_column.append(item['detail'] + ' + ' + item['financial_entity'])
                column_name = list(set(new_column))
                df4 = pd.DataFrame(columns=column_name)
                for tax in range(len(df3["TAXES_DISAGGREGATED"])):
                    df4.loc[tax] = {d['detail'] + ' + ' + d['financial_entity'] : d['amount'] for d in df3["TAXES_DISAGGREGATED"][tax]}
                df4_filled = df4.fillna(0)
                df4_filled["json_sum"] = df4_filled.sum(axis=1)
                df5 = pd.concat([df3, df4_filled], axis=1)
                df5["control"] = df5["json_sum"] - df5["TAXES_AMOUNT"]
                df.drop(["TAXES_DISAGGREGATED", "TAXES_AMOUNT"], axis=1, inplace=True)
                final_df = pd.concat([df, df5], axis = 1)
                
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine="xlsxwriter")
                final_df.to_excel(writer, index=False, sheet_name="Sheet1")
                writer.close()
                data_bytes = output.getvalue()
                
                st.download_button(label="Download Excel",
                               data=data_bytes,
                               file_name="output.xlsx",
                               on_click=send_email_after_download
                            )
                
            elif "IMPUESTOS_DESAGREGADOS" and "MONTO_IMPUESTOS" in df.columns:
                df1 = df[["IMPUESTOS_DESAGREGADOS", "MONTO_IMPUESTOS"]]
                st.session_state.total_row = len(df1)
                df2 = df1.copy()
                df2["IMPUESTOS_DESAGREGADOS"] = df1["IMPUESTOS_DESAGREGADOS"].apply(remove_double_quote)
                df3 = df2.copy()
                df3["IMPUESTOS_DESAGREGADOS"] = df2["IMPUESTOS_DESAGREGADOS"].apply(convert_to_dict)
                new_column = []
                for d in df3["IMPUESTOS_DESAGREGADOS"]:
                    for item in d:
                        new_column.append(item['detail'] + ' + ' + item['financial_entity'])
                column_name = list(set(new_column))
                df4 = pd.DataFrame(columns=column_name)
                for tax in range(len(df3["IMPUESTOS_DESAGREGADOS"])):
                    df4.loc[tax] = {d['detail'] + ' + ' + d['financial_entity'] : d['amount'] for d in df3["IMPUESTOS_DESAGREGADOS"][tax]}
                df4_filled = df4.fillna(0)
                df4_filled["json_sum"] = df4_filled.sum(axis=1)
                df5 = pd.concat([df3, df4_filled], axis=1)
                df5["control"] = df5["json_sum"] - df5["MONTO_IMPUESTOS"]
                df.drop(["IMPUESTOS_DESAGREGADOS", "MONTO_IMPUESTOS"], axis=1, inplace=True)
                final_df = pd.concat([df, df5], axis = 1)
                
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine="xlsxwriter")
                final_df.to_excel(writer, index=False, sheet_name="Sheet1")
                writer.close()
                data_bytes = output.getvalue()
                
                st.download_button(label="Download Excel",
                               data=data_bytes,
                               file_name="output.xlsx",
                               on_click=send_email_after_download
                            )
            else:
                st.error("Your Excel file not contain taxes dissagregated or impuestos desagregados column and taxes amount or monto impuestos column")
