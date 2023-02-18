import smtplib
from email import encoders
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

participants_ =[]
df = pd.read_csv("sample.csv")
fromaddr = "ecellucekofficial@gmail.com"
toaddr = ""
appPass = "deggkntteeisiuaf"
attachment = None


body = """
Hi #,

Hope you are doing well!
It was amazing to have you participate in the Illuminate Workshop 2022 and see you grow & learn so much .We hope your participation in the future events.

All the best for your future endeavours.

The Participation certificate is hereby attatched with this mail. Please find the same .

With regards,
Team E-Cell UCEK

"""

def send_mail(name, email_):
    try:       
        toaddr = email_
        msg = MIMEMultipart()
        msg["From"] = fromaddr
        msg["To"] = toaddr
        msg["Subject"] = "Participation Certificate from E-Cell UCEK"
        
        name_copy = name.title()
        body_copy = body.replace("#", name_copy)
        # body_copy = body_copy.replace("#", name)
        
        msg.attach(MIMEText(body_copy, "plain"))
        
        filename = name
        loc = "certificate/{}.pdf".format(name)

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
        with open('sample.txt','a+') as f:
            f.write("\n")
            f.write(name)
        return "Mail sent to " + name
    
    except Exception as e:
        with open('error.txt','a+') as f:
            f.write("\n")
            f.write(name)
        return e

input("Press enter to start")

for participant, mail_id in  zip(df['Name'],df['Email id']):
    print(send_mail(str(participant), str(mail_id)))