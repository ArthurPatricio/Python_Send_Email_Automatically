import pandas as pd
from datetime import datetime, timedelta
import smtplib
import os
from email.message import EmailMessage


def send_email():
    
    # SPREADSHEET DATA SOURCE
    df = pd.read_excel('spreadsheet.xlsx', sheet_name='tab name')

    today_plus_10 = datetime.today() + timedelta(days=10)
    today_plus_10 = today_plus_10.date()
    today = datetime.today()
    today = today.date()

    for index, row in df.iterrows():
        expiration_date_timestamp =row['Valid Until']
        expiration_date = expiration_date_timestamp.to_pydatetime()
        expiration_date = expiration_date.date()
        activation_date_timestamp =row['Activation Date']
        activation_date = activation_date_timestamp.to_pydatetime()
        activation_date = activation_date.date()
        delta = today_plus_10 - expiration_date


        if delta.days > 0 and expiration_date.day - today.day > 0 and str(row['Acknowledgment']).lower() == 'nan':

            # LIST OF EMAILS TO SEND TO
            email_list = ['email address 1, email address 2, ...']

            msg = EmailMessage()
            msg['Subject'] = f'A LICENÇA DO CLIENTE {row["Client"]} ({row["Environment"]}) ESTÁ PARA EXPIRAR (EM {expiration_date.day - today.day} DIAS)'
            # EMAIL ADDRESS THAT SENDS THE EMAIL
            msg['From'] = 'email address that sends the email'
            msg['To'] = email_list
            msg.set_content(f'Uma licença do cliente {row["Client"]} de irá expirar em {expiration_date} \n\n \
Ambiente: {row["Environment"]} \n\n \
Natureza da Licença: {row["Subject"]} \n\n \
Data de Ativação: {activation_date} \n\n \
Número Serial: {row["Serial Number"]} \n\n \
Chave: {row["Key"]}')

            with smtplib.SMTP('smtp.outlook.com', 587) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                
                # EMAIL LOGIN
                smtp.login('email address for login', 'email password for login')

                smtp.send_message(msg)
            #print('A licença do cliente', row['Client'],'vai expirar em' , expiration_date)


if __name__ == '__main__':
    send_email()
