import pandas as pd
from datetime import datetime, timedelta
import smtplib
import os
from email.message import EmailMessage


def send_email():

    df = pd.read_excel('xxx')

    for index, row in df.iterrows():
        pandas_timestamp =row['Valid Until']
        license_date = pandas_timestamp.to_pydatetime()
        license_date = license_date.date()
        today = datetime.today() + timedelta(days=10)
        today = today.date()
        delta = today - license_date
        if delta.days > 0:
            
            email_list = ['xxx']

            msg = EmailMessage()
            msg['Subject'] = f'A LICENÇA DO CLIENTE {row["Client"]} ESTÁ PARA EXPIRAR'
            msg['From'] = 'xxx'
            msg['To'] = email_list
            msg.set_content(f'A licença do cliente {row["Client"]} vai expirar em: {delta.days} dias')

            with smtplib.SMTP('smtp.outlook.com', 587) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()

                #smtp.login(EMAIL, PASSWORD)
                smtp.login('xxx', 'xxx')

                #subject ='PYTHON EMAIL SENDER TEST'
                #body = 'TESTING ON HOW TO SEND EMAILS WITH PYTHON!!'

                #msg = f'Subject: {subject}\n\n{body}'

                #smtp.sendmail('support@netconamericas.com', email_list, msg)

                smtp.send_message(msg)
            print('A licença do cliente', row['Client'],'vai expirar em:' , delta.days, 'dias')


if __name__ == '__main__':
    send_email()
