import os


def convert_to_pdf(docx_file, output_dir):
    try:
        # Get the name of the docx file without the extension
        file_name = os.path.splitext(docx_file)[0]
        # Construct the command to convert the file using LibreOffice
        command = f"libreoffice --headless --convert-to pdf {docx_file} --outdir {output_dir}"
        # Run the command in the shell
        os.system(command)
        # Move the output pdf file to the specified output directory
        os.rename(f"{os.getcwd()}/{file_name}.pdf", f"{output_dir}/{file_name}.pdf")
    except Exception as e:
        print("Error at convert_to_pdf ",e)



