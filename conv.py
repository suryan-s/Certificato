import os


def convert_to_pdf(docx_file, output_dir, name, filename):
    try:
        command = f"libreoffice --headless --convert-to pdf:writer_pdf_Export {docx_file} --outdir {output_dir} --infilter=\"Microsoft Word 2007-2013 XML\" "
        # Run the command in the shell
        os.system(command)
        # Move the output pdf file to the specified output directory
        os.rename(f"{output_dir}/{filename}.pdf", f"{output_dir}/{name}.pdf")
    except Exception as e:
        print("Error at convert_to_pdf ",e)



