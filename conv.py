import os
import pypandoc

def convert_to_pdf(docx_file, output_dir, name, filename):
    try:
        # command = f"libreoffice --headless --convert-to pdf:writer_pdf_Export {docx_file} --outdir {output_dir} --infilter=\"Microsoft Word 2007-2013 XML\" "
        # command = f"pandoc {docx_file} -o {output_dir}/{name}.pdf"
        command = f"unoconv -f pdf {docx_file}"
        # Run the command in the shell
        os.system(command)
        new_pdf_file = docx_file.replace(".docx", ".pdf")
        # Move the output pdf file to the specified output directory
        os.rename(new_pdf_file, f"{output_dir}/{name}.pdf")
        # output = pypandoc.convert_file(docx_file, 'pdf', outputfile=os.path.join(output_dir, f"{name}.pdf"))
        # Verify the pdf file was created
        # if output == "":
        #     print(f"PDF file created successfully at {output_dir}/{name}.pdf")
        # else:
        #     print("PDF file creation failed.")
    except Exception as e:
        print("Error at convert_to_pdf ",e)



