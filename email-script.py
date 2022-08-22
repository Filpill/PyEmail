import email_address_list
import smtplib
import sys
from datetime import date
from dateutil.relativedelta import relativedelta
from os.path import basename
from dateutil.relativedelta import relativedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def email_new(email_tuple):

    # Dates Formatted as Text
    prev_month_file = (date.today() - relativedelta(months=1)).strftime("%B-%Y")
    prev_month_body = (date.today() - relativedelta(months=1)).strftime("%B %Y")

    #Addressees and Body of Text
    message = MIMEMultipart()
    message['Subject'] =  f"Python Email Testing - {email_tuple[2]} Presentation Data - {prev_month_body}"
    message['From'] = email_address_list.sender
    message['To'] = email_tuple[0]
    message['Cc'] = email_address_list.safety_analysts
    text_body = f"**THIS IS A TEST EMAIL**\n\nHi all,\n\nPlease find attached the {email_tuple[2]} presentation containing data for {prev_month_body}.\n\nKind Regards,\n\nFilip Livancic\n\n***Email sent with Python via SMTP***"
    message.attach(MIMEText(text_body))

    #Message Attachment
    file_extension = ".txt"
    powerpoint_file = email_tuple[1]+fr"_{prev_month_file}{file_extension}"
    file = sys.path[0] + fr"\powerpoints\{powerpoint_file}"
    with open(file, "rb") as fil:
        part = MIMEApplication(
            fil.read(),
            Name=basename(file)
        )
    # After the file is closed
    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file)
    message.attach(part)

    # Connect to SMTP server and Send Mail
    with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
        server.starttls()
        server.sendmail(email_address_list.sender, email_tuple[0], message.as_string())

if __name__ == '__main__':

    for email_tuple in email_address_list.address_list:
        print(f'Sending Email to: {email_tuple[0]} with attached {email_tuple[1]}')
        email_new(email_tuple)
