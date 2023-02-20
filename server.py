import os
import shutil
import uuid

import pandas as pd
import streamlit as st

from func import prep_cert

# Set page title and center the page heading
st.set_page_config(page_title="Certificato")
# st.title("Certificato")
st.write(
    """
    <h1 style='text-align: center;'>Certificato</h1>
    <hr>
    """,
    unsafe_allow_html=True,
)
st.markdown("<h2 style='text-align: center;'>Made by suryan</h2>", unsafe_allow_html=True)

df = None
name_col = None
email_col = None
prev_path  = None

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
        st.write("### Column names:")
        st.write(list(df.columns))

        # Add a dropdown menu to select the name and email columns
        name_col = st.selectbox("Select the names column:", options=df.columns)
        email_col = st.selectbox("Select the email id column:", options=df.columns)

    # Add a field to enter email ID and app password
    if "df" in locals():
        # name = st.text_input("Enter the name:")

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
                email_title = st.text_input("Email Title")
                email_body = st.text_area("Email Body")
                # Show a button to send the certificate
                if st.button("Send") and len(email_title) > 0 and len(email_body) > 0:
                    Result = None
                    # st.write(f"Certificate sent to {email}!")
                    with st.spinner("Your request is under progress"):
                        # Call create_cert function
                        # filename = "\\temp\\temp_" + str(uuid.uuid4())
                        # Parent Directories
                        parent_dir = os.getcwd()
                        print("parent_dir: ", parent_dir)
                        # Path
                        # main_temp_folder = parent_dir + filename
                        var = str(uuid.uuid4())
                        main_temp_folder = os.path.join(parent_dir, 'temp', 'temp_' + var)
                        main_cert_folder = os.path.join(main_temp_folder, 'certificates')
                        down_folder = os.path.join(parent_dir, 'downloads')
                        path4 = os.path.join(main_temp_folder, 'result.xlsx')
                        main_down_folder = os.path.join(parent_dir, 'downloads', 'temp_' + var)
                        print("main_temp_folder: ", main_temp_folder)
                        print("main_cert_folder: ", main_cert_folder)
                        print("down_folder: ", down_folder)
                        # Create the directory
                        try:
                            os.umask(0)
                            os.makedirs(main_temp_folder,mode=0o777)
                            os.makedirs(main_cert_folder,mode=0o777)
                            os.makedirs(down_folder,mode=0o777)
                        except Exception as e:
                            print("Error at making dir ",e)
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
                        if Result==200:
                            # Replace loading icon with completed message
                            st.success("Your request is completed")

                            # Download button 
                            with st.spinner("Loading zip file..."):
                                shutil.make_archive(down_folder, "zip", main_temp_folder)
                                st.success("Zipping completed")
                            zip_name = '{}.zip'.format(main_down_folder)
                            os.chmod(zip_name,  0o777)
                            if st.button("Download certificates with status results?"):
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
                                                for f in files:
                                                    os.chmod(os.path.join(root, f), 0o777)
                                os.remove(zip_name)
                                shutil.rmtree(main_temp_folder)
                                st.success("Completed")
                            
                        elif Result == 500:
                            st.error("Certificate creation failed. Please try again.")
                        
                        
                        # st.success("Completed")
            else:
                st.error("Invalid credentials.")
