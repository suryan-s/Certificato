import os
import smtplib
import uuid
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from multiprocessing import Manager, Pool
from pathlib import Path


import subprocess

try:
    import comtypes.client
    import pypandoc
except ImportError:
    client = None
    import pypandoc
import pandas as pd
from docxtpl import DocxTemplate

# from time import sleep
import subprocess


def start(var):
        print('2')
        temp_name, temp_stat = send_mail(var[0],var[1],var[2],var[3],var[4],var[5],var[6],var[7])
        print('3')
        return temp_name, temp_stat

def create_cert(receiver,fileloc,docx_file):
    # temp_file = fileloc + '\\temp_'+ str(uuid.uuid4()) +'.docx' 
    temp_file = os.path.join(fileloc,'temp_'+ str(uuid.uuid4()) +'.docx') 
    # out_file = fileloc+"\\certificates\\{}.pdf".format(receiver)
    out_file = os.path.join(fileloc,"certificates","{}.pdf".format(receiver))
    # out_file_ = fileloc+"\\certificates\\{}.docx".format(receiver)
    out_file_ = os.path.join(fileloc,"certificates","{}.docx".format(receiver))
    out_file__ = os.path.join(fileloc,"certificates","{}.pdf".format(receiver))
    # CFG
    print("temp_file",temp_file)  
    print("out_file",out_file__) 

    # Fill in text
    data_to_fill = {'value' : str(receiver),}

    template = DocxTemplate(docx_file)
    template.render(data_to_fill)
    

    # Convert to PDF
    wdFormatPDF = 17

    in_file = os.path.abspath(Path(temp_file))
    out_file = os.path.abspath(Path(out_file))
    client = None
    if client != None:
        try:
            template.save(Path(temp_file))
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(in_file) # type: ignore
            doc.SaveAs(out_file, FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit() # type: ignore
            os.chmod(temp_file,  0o777)
            os.remove(temp_file)
        except Exception as e:
            print("Error at create_cert ",e)
    else:
        # pypandoc.convert_file(temp_file, 'pdf', outputfile=out_file__)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", out_file_])       
    

def prep_cert(st_dataFrame,docx_file, file_loc, st_col_name, st_col_email, st_fromaddr, st_appPass, st_subject, st_body):   
    
    name_ = []
    stat_ = []
    work = []
    temps = []
    result = bool
    columns = ['Name', 'Status']
    print('1')
    print("file_loc at prep_cert",file_loc)
        
    df = st_dataFrame   
    try:
        print('4')
        for participant, mail_id in  zip(df[st_col_name],df[st_col_email]):
            print('5')
            work.append([file_loc,docx_file, participant, mail_id, st_fromaddr, st_appPass, st_subject, st_body])            
        p = Pool(int(len(work)/2))
        # p = Pool(2)
        print('6')
        print(work)
        temps.append(p.map(start, work)) 
        for temp in temps[0]:
            new_temp = list(temp)
            name_.append(new_temp[0])
            stat_.append(new_temp[1])
        
    except Exception as e:
        print("Error at prep_cert ",e)
        result = False
    finally:
        result_df = pd.DataFrame(list(zip(name_,stat_)), columns=columns)
        with pd.ExcelWriter(os.path.join(file_loc,'result.xlsx')) as writer:
            result_df.to_excel(writer,index=False)
    
    # Your code to create certificate goes here
    # Return True if certificate creation is successful, else 500
    return True if result else 500

def send_mail(loc__,docx_file, name, toaddr, fromaddr, appPass, subject, body):    
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
        print("loc__",loc__)
        
        create_cert(name,loc__,docx_file)
        
        filename = name
        # loct = os.getcwd()
        # loc = loc__ + "\\certificates\\{}.pdf".format(name)
        loc = os.path.join(loc__,"certificates","{}.pdf".format(name))

        with open(loc, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")

        attachment.add_header("Content-Disposition", "attachment; filename= %s" % filename)
        msg.attach(attachment)
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()

        s.login(fromaddr, appPass)
        
        text = msg.as_string()
        status = s.sendmail(fromaddr, toaddr, text)
        # print(status)
        s.quit()
        return name, 'Success'
    
    except Exception as e:
        print("Error at send_mail ",e)
        return name, 'Failed'
    
