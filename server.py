"""
Certificato - A web for sending certificates
Created by Suryan S
Copyright (c)
Licensed under MIT License
"""
import os
import shutil
import uuid

import pandas as pd
import streamlit as st
from streamlit_extras.badges import badge
from streamlit_extras.dataframe_explorer import dataframe_explorer
from streamlit_extras.let_it_rain import rain

from func import prep_cert

# Declare variables
df = None
df_set = False
name_col = None
email_col = None
prev_path = None

# Set page title and center the page heading
st.set_page_config(page_title="Certificato")

st.sidebar.write(
    "Hi, Certificato is a web app to create certificates for your participants easily and send them via mail avoiding the hassle of creating certificates manually."
    )
st.sidebar.write("Made with ❤️ by Suryan!")
st.sidebar.link_button("Check out my GitHub repo", "https://github.com/suryan-s")

html = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Courgette&display=swap');

    h1 {
        font-size: 4em;
        text-align: center;
        background: linear-gradient(90deg, rgba(63,94,251,1) 0%, rgba(252,70,107,1) 100%);
        background-clip: text;
        -webkit-background-clip: text;
        font-family: 'Courgette', cursive;
        color: transparent;
        letter-spacing: 4px;
    }
</style>
<h1>Certificato</h1>
<hr>
"""

st.markdown(html, unsafe_allow_html=True)
badge(type="github", name="suryan-s/Certificato")

st.write("#### Upload files:")

# csv file upload
# word file upload
uploaded_word_file = st.file_uploader(
    "Upload the Certificate as docx file", type=["docx", "doc"]
    )

if uploaded_word_file is not None:
    # Add an upload option to upload an Excel file
    uploaded_excel_file = st.file_uploader(
        "Upload the Excel file", type=["xlsx", "xls"]
        )

    # If an Excel file is uploaded, show the name of the columns
    if uploaded_excel_file is not None:
        df = pd.read_excel(uploaded_excel_file)
        df_set = True
        filtered_df = dataframe_explorer(df, case=False)
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
        st.divider()
        st.write("#### Set column:")
        st.write(list(df.columns))

        # Add a dropdown menu to select the name and email columns
        name_col = st.selectbox("Select the names column:", options=df.columns)
        email_col = st.selectbox("Select the email id column:", options=df.columns)

    # Add a field to enter email ID and app password
    if "df" in locals() and df_set:
        st.divider()
        st.write("#### Set credentials:")
        email = st.text_input(
            "Enter the email ID from which you want to send them certificates to:"
            )
        password = st.text_input(
            "Enter the app password of the email:", type="password"
            )

        # Validate the email and password format
        if email and password:
            if "@" in email and "." in email and len(password) >= 6:
                st.success("Credentials are valid!")
                # Show email configs
                st.divider()
                st.write("#### Set email body:")
                email_title = st.text_input("Email Title")
                email_body = st.text_area("Email Body")
                # Show a button to send the certificate
                if (
                        st.button("Send Certificate")
                        and len(email_title) > 0
                        and len(email_body) > 0
                ):
                    Result = None
                    # st.write(f"Certificate sent to {email}!")
                    with st.spinner("Your request is under progress"):
                        parent_dir = os.getcwd()
                        var = str(uuid.uuid4())
                        main_temp_folder = os.path.join(
                            parent_dir, "temp", "temp_" + var
                            )
                        main_cert_folder = os.path.join(
                            main_temp_folder, "certificates"
                            )
                        down_folder = os.path.join(parent_dir, "downloads")
                        path4 = os.path.join(main_temp_folder, "result.xlsx")
                        zip_file = os.path.join(parent_dir, "downloads", "cert_" + var)
                        zip_name = os.path.join(
                            parent_dir, "downloads", "cert_" + var + ".zip"
                            )
                        # Create the directory
                        try:
                            os.umask(0)
                            os.makedirs(main_temp_folder, exist_ok=True)
                            os.makedirs(main_cert_folder, exist_ok=True)
                            os.makedirs(down_folder, exist_ok=True)
                        except Exception as e:
                            print("Error at making dir ", e)
                            st.error("Error at making dir")
                        try:
                            Result = prep_cert(
                                df,
                                uploaded_word_file,
                                main_temp_folder,
                                name_col,
                                email_col,
                                email,
                                password,
                                email_title,
                                email_body,
                                )
                        except ValueError:
                            st.error("Number of requests must be at least 1")
                        if Result == 200:
                            # Replace loading icon with completed message
                            rain(emoji="🎈", font_size=5)
                            st.success("Your request is completed")

                            # Zip and Download button
                            shutil.make_archive(zip_file, "zip", main_temp_folder)
                            if os.path.exists(zip_name):
                                st.success("Zip file created successfully!")
                                os.chmod(zip_name, 0o777)
                                with open(zip_name, "rb") as f:
                                    bytes_data = f.read()
                                    st.download_button(
                                        "Download",
                                        data=bytes_data,
                                        file_name="certificates.zip",
                                        mime="application/zip",
                                        )
                                for contents in os.listdir(main_temp_folder):
                                    for root, dirs, files in os.walk(contents):
                                        for d in dirs:
                                            os.chmod(os.path.join(root, d), 0o777)
                                            try:
                                                shutil.rmtree(os.path.join(root, d))
                                            except Exception:
                                                pass
                                        for f in files:
                                            os.chmod(os.path.join(root, f), 0o777)
                                            try:
                                                os.remove(os.path.join(root, f))
                                            except Exception:
                                                pass
                                try:
                                    shutil.rmtree(main_temp_folder)
                                except Exception:
                                    pass
                                st.success("Completed")
                            else:
                                st.error("Zip file creation failed.")
                                st.stop()

                        elif Result == 500:
                            st.error("Certificate creation failed. Please try again.")
            else:
                st.error("Invalid credentials.")