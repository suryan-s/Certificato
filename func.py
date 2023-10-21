# import multiprocessing
import os
import smtplib
import warnings
import uuid
from email.message import EmailMessage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from multiprocessing import Pool, cpu_count
from pathlib import Path

from jinja2 import Environment, FileSystemLoader

env = Environment(loader=FileSystemLoader("templates"))
template = env.get_template("html.html")

# Filter out the specific warning message
warnings.filterwarnings("ignore", category=UserWarning, message="Warning: to view a Streamlit app*")

client = True
try:
    import comtypes.client

except ImportError:
    client = False

import pandas as pd
from docxtpl import DocxTemplate


def convert_to_pdf(docx_file, output_dir, name, filename):
    try:
        command = f'libreoffice --headless --convert-to pdf:writer_pdf_Export {docx_file} --outdir {output_dir} --infilter="Microsoft Word 2007-2013 XML" '
        os.system(command)
        os.rename(f"{output_dir}/{filename}.pdf", f"{output_dir}/{name}.pdf")

    except Exception as e:
        print("Error at convert_to_pdf ", e)


def start(var):
    temp_name, temp_stat = send_mail_custom(
        var[0], var[1], var[2], var[3], var[4], var[5], var[6], var[7]
        )
    return temp_name, temp_stat


def create_cert(receiver, fileloc, docx_file):
    try:
        # CFG
        var = str(uuid.uuid4())
        temp_doc_file = os.path.join(fileloc, "temp_" + var + ".docx")
        temp_doc_folder = os.path.join(fileloc, "temp_" + var)
        cert_file_loc = os.path.join(fileloc, "certificates", "{}.pdf".format(receiver))
        cert_folder_loc = os.path.join(fileloc, "certificates")

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
            convert_to_pdf(temp_doc_file, cert_folder_loc, receiver, "temp_" + var)
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

    df = st_dataFrame
    try:
        for participant, mail_id in zip(df[st_col_name], df[st_col_email]):
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
        # pool = Pool(int(len(work) / 2)) if len(work) != 1 else Pool(1) 
        pool = Pool(cpu_count())
        temps = [pool.map(start, work)]
        new_temps = [list(temp) for temp in temps[0]]
        name_ = [new_temp[0] for new_temp in new_temps]
        stat_ = [new_temp[1] for new_temp in new_temps]
        result = True
        print("Finished processing")
        pool.close()

    except Exception as e:
        print("Error at prep_cert ", e)
        result = False
    finally:
        result_df = pd.DataFrame(list(zip(name_, stat_)), columns=columns)
        with pd.ExcelWriter(os.path.join(edited_file_loc, "result.xlsx")) as writer:
            result_df.to_excel(writer, index=False)

    # Return True if certificate creation is successful, else 500
    return 200 if result else 500


def send_mail(loc__, docx_file, name, toaddr, fromaddr, appPass, subject, body):
    msg = MIMEMultipart()
    msg["From"] = fromaddr
    msg["To"] = toaddr
    msg["Subject"] = subject
    try:
        # name_copy = name.title()
        body_copy = body.replace("#", name.title())

        msg.attach(MIMEText(body_copy, "plain"))

        create_cert(name, loc__, docx_file)

        loc = os.path.join(loc__, "certificates", "{}.pdf".format(name))

        with open(loc, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")

        attachment.add_header("Content-Disposition", "attachment; filename= %s" % name)
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


# def send_mail_custom():
#     msg = EmailMessage()
#     content = None
#     EMAIL_ADDRESS = 'tinkerhubucek@gmail.com'
#     EMAIL_PASSWORD = 'kiebnxpaqizoisbo'
#     msg['Subject'] = 'This is a test message from TinkerHub UCEK'
#     msg['From'] = 'tinkerhubucek@gmail.com'
#     msg['To'] = 'arjun8107@gmail.com'
#     with open('body.html', 'r', encoding="utf-8") as f:
#         content = f.read()
#     msg.set_content(content, subtype='html')


#     with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
#         smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
#         smtp.send_message(msg)


def send_mail_custom(loc__, docx_file, name, toaddr, fromaddr, appPass, subject, body):
    msg = EmailMessage()
    msg["From"] = fromaddr
    msg["To"] = toaddr
    msg["Subject"] = subject
    try:
        template_copy = template
        data = { "value": name.title() }

        rendered_html = template_copy.render(data)

        msg.set_content(rendered_html, subtype="html")

        create_cert(name, loc__, docx_file)

        loc = os.path.join(loc__, "certificates", "{}.pdf".format(name))

        with open(loc, "rb") as pdf:
            msg.add_attachment(
                pdf.read(),
                maintype="application",
                subtype="octet-stream",
                filename=name + ".pdf",
                )
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()

        s.login(fromaddr, appPass)

        text = msg.as_string()
        status = s.sendmail(fromaddr, toaddr, text)
        if status == { }:
            print("Mail sent to ", toaddr)
        else:
            print("Mail sending failed to ", toaddr)
        s.quit()
        # print("Finished")
        return name, "Success"

    except Exception as e:
        print("Error at send_mail ", e)
        return name, "Failed"