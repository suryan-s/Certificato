import os
import pypandoc

def convert_to_pdf(docx_file, output_dir, name, filename):
    try:
        # command = f"libreoffice --headless --convert-to pdf:writer_pdf_Export {docx_file} --outdir {output_dir} --infilter=\"Microsoft Word 2007-2013 XML\" "
        # command = f"pandoc {docx_file} -o {output_dir}/{name}.pdf"
        # Run the command in the shell
        # os.system(command)
        # Move the output pdf file to the specified output directory
        # os.rename(f"{output_dir}/{filename}.pdf", f"{output_dir}/{name}.pdf")
        output_options = {
            'geometry': 'margin=1in',
            'papersize': 'a3'
        }
        output = pypandoc.convert_file(docx_file, 'pdf', outputfile=os.path.join(output_dir, f"{name}.pdf"), extra_args=['--variable', f'{k}:{v}' for k,v in output_options.items()])
        # Verify the pdf file was created
        if output == "":
            print(f"PDF file created successfully at {output_dir}/{name}.pdf")
        else:
            print("PDF file creation failed.")
    except Exception as e:
        print("Error at convert_to_pdf ",e)



