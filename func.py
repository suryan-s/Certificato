import os
import smtplib
import subprocess
import time
import uuid
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from multiprocessing import Pool
from pathlib import Path
from subprocess import call

from conv import convert_to_pdf

client = True
try:
    import comtypes.client

except ImportError:
    client = False
    # from pypandoc.pandoc_download import download_pandoc
    # targetfolder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "pypandoc", "files")
    # download_pandoc(targetfolder=targetfolder)
    # from pypandoc.pandoc_download import download_pandoc
    # download_pandoc()
    # import pypandoc

    # import pypandoc.pandoc_download
    # command = "apt-get install pandoc"
    # os.system(command)
    # print(pypandoc.get_pandoc_version())
    # print(pypandoc.get_pandoc_path())
    # print(pypandoc.get_pandoc_formats())
    # os.environ.setdefault('PYPANDOC_PANDOC', pypandoc.get_pandoc_path())

import pandas as pd
from docxtpl import DocxTemplate


def start(var):
    # print("2")
    temp_name, temp_stat = send_mail(
        var[0], var[1], var[2], var[3], var[4], var[5], var[6], var[7]
    )
    # print("3")
    return temp_name, temp_stat


def create_cert(receiver, fileloc, docx_file):
    try:
        
        # CFG
        var = str(uuid.uuid4())
        temp_doc_file = os.path.join(fileloc, "temp_" + var + ".docx")
        temp_doc_folder = os.path.join(fileloc, "temp_" + var)
        cert_file_loc = os.path.join(fileloc, "certificates", "{}.pdf".format(receiver))
        cert_folder_loc = os.path.join(fileloc, "certificates")
        
        # print("temp dcx loc: ", temp_doc_file)
        # print("out dir: ", cert_folder_loc)

        # Fill in text
        data_to_fill = {
            "value": str(receiver),
        }

        template = DocxTemplate(docx_file)
        template.render(data_to_fill)

        # Convert to PDF
        wdFormatPDF = 17

        in_file = os.path.abspath(Path(temp_doc_file))
        cert_file_loc = os.path.abspath(Path(cert_file_loc))
        # client = None
        if client == True:
            try:
                template.save(Path(temp_doc_file))
                word = comtypes.client.CreateObject("Word.Application")
                doc = word.Documents.Open(in_file)  # type: ignore
                doc.SaveAs(cert_file_loc, FileFormat=wdFormatPDF)
                doc.Close()
                word.Quit()  # type: ignore
                os.chmod(temp_doc_file, 0o777)
                os.remove(temp_doc_file)
            except Exception as e:
                print("Error at create_cert ", e)
        else:
            template.save(Path(temp_doc_file))
            # cert_file_loc = os.path.join(fileloc, "certificates", "{}.pdf".format(receiver))
            convert_to_pdf(temp_doc_file, cert_folder_loc,receiver, "temp_" + var)
            # pypandoc.convert_file(temp_doc_file, 'pdf', outputfile=cert_file_loc)
            # subprocess.run(['unoconv', '-f', 'pdf', temp_doc_file])
            # call(
            #     f"libreoffice --headless --convert-to pdf --outdir {cert_folder_loc} {temp_doc_file}",
            #     shell=True,
            # )
    except Exception as e:
        print("Error at create_cert ", e)


def prep_cert(
    st_dataFrame,
    docx_file,
    edited_file_loc,
    st_col_name,
    st_col_email,
    st_fromaddr,
    st_appPass,
    st_subject,
    st_body,
):
    name_ = []
    stat_ = []
    work = []
    temps = []
    result = bool
    columns = ["Name", "Status"]
    # print("1")
    # print("edited_file_loc at prep_cert", edited_file_loc)

    df = st_dataFrame
    try:
        # print("4")
        for participant, mail_id in zip(df[st_col_name], df[st_col_email]):
            # print("5")
            work.append(
                [
                    edited_file_loc,
                    docx_file,
                    participant,
                    mail_id,
                    st_fromaddr,
                    st_appPass,
                    st_subject,
                    st_body,
                ]
            )
        p = Pool(int(len(work) / 2))
        # print("6")
        # print(work)
        
        temps.append(p.map(start, work))        
        
        for temp in temps[0]:
            new_temp = list(temp)
            name_.append(new_temp[0])
            stat_.append(new_temp[1])
        result = True

    except Exception as e:
        print("Error at prep_cert ", e)
        result = False
    finally:
        result_df = pd.DataFrame(list(zip(name_, stat_)), columns=columns)
        with pd.ExcelWriter(os.path.join(edited_file_loc, "result.xlsx")) as writer:
            result_df.to_excel(writer, index=False)

    # Your code to create certificate goes here
    # Return True if certificate creation is successful, else 500
    return 200 if result == True else 500


def send_mail(loc__, docx_file, name, toaddr, fromaddr, appPass, subject, body):
    msg = MIMEMultipart()
    msg["From"] = fromaddr
    msg["To"] = toaddr
    msg["Subject"] = subject
    try:
        # df = pd.DataFrame(pd.read_excel(file))

        name_copy = name.title()
        body_copy = body.replace("#", name_copy)
        # body_copy = body_copy.replace("#", name)

        msg.attach(MIMEText(body_copy, "plain"))
        # print("loc__", loc__)

        create_cert(name, loc__, docx_file)
        
        loc = os.path.join(loc__, "certificates", "{}.pdf".format(name))

        with open(loc, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")

        attachment.add_header(
            "Content-Disposition", "attachment; filename= %s" % name
        )
        msg.attach(attachment)
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()

        s.login(fromaddr, appPass)

        text = msg.as_string()
        status = s.sendmail(fromaddr, toaddr, text)
        # print(status)
        s.quit()
        return name, "Success"

    except Exception as e:
        print("Error at send_mail ", e)
        return name, "Failed"
