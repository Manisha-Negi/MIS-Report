import cx_Oracle
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sys

path_xl = '/home/oracle/RakshaTPAExcel/xl/'
listOfEmailIds = ['manishanegi151296@gmail.com', 'rohitcrk1@gmail.com']
to = listOfEmailIds

def WriteToExcel():
    conn = cx_Oracle.connect('CONSULTIT/Alpha1234@XePDB1')
    df = pd.read_sql("""SELECT * FROM TPARAKSHA.KMD_DAILY_CASHLESS@rds""", conn)
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(os.path.join(path_xl, "MIS.xlsx"), engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object. Note that we turn off
    # the default header and skip one row to allow us to insert a user defined
    # header.
    df.to_excel(writer, sheet_name='Inception to bill', startrow=1, header=False, index=False)
    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Inception to bill']
    # Add a header format.
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'center',
        'fg_color': '#FFA07A',
        'border': 1,
        'font_color': 'Black'})

    data = ['TPA Claim No', 'Insurance Company', 'Region', 'DO', 'BO', 'Policy NO', 'Affinity/Non Affinity', 'Corporate', 'Policy Holder Name', 'Policy Type', 'Pol Development Officer', 'Pol Development Agent', 'Policy Start', 'Policy End Date', 'Employee Code', 'Employee Name', 'MAID', 'Claiments Name', 'Benef Age', 'Benef Area Code', 'Benef Sex', 'Rel Name', 'Sum Insured', 'Balance Sum Insured', 'Basic Sum Insured', 'TPA Claim No', 'Cigna Claim No', 'Claim Type', 'Claim Type', 'Final Status', 'TPA Status', 'PA Status', 'TPA', 'Clm Received Date', 'Date of Admission', 'Date of Dischage', 'Approved Date', 'Settled Date/Decisioned Date', 'Last Audit date-LDR', 'Incurred Amount', 'Claimed Amount', 'Approved Amount', 'Reopen Status', 'Reopen Month', 'Ailment Code', 'Illness', 'Ailment Grp', 'Procedure Type Surgical Non Surgical', 'Document Required', 'Hosp ID', 'Hospital Name', 'City Name', 'NETWORK NON NETWORK', 'ICD Chapter Codes', 'Deductible', 'Cheque No', 'Cigna Approved Amount', 'Product Code', 'Denial Reasons', 'IPD/OPD', 'Customer ID', 'Reporting Month', 'Produc', 'Affinity / Non Affinity']
    row = 0
    col = 0
    

    # Write the column headers with the defined format.
    for name in data:
        # print(col_num)
        # print(value)
        worksheet.write(row, col, name, header_format)
        col += 1
    cell_data_format = workbook.add_format({'text_wrap':False})
    cell_data_format.set_align('left')
    worksheet.set_column('A:BL', 20, cell_data_format)
    
    # Close the Pandas Excel writer and output the Excel file.

    writer.save()
    conn.close()



def SendEmail():
    try:

        # open the file in bynary
        xl_file = open(os.path.join(path_xl, 'MIS.xlsx'), 'rb')
        # Step 2 - Create message object instance

        msg = MIMEMultipart()

        # Step 3 - Create message body

        message = "Dear Sir/Madam, \n\nPlease find the attached MIS report \n\nWith regards \nRaksha Health Insurance TPA"

        # Step 4 - Declare SMTP credentials

        password = "info#@5164"

        username = "info"

        # smtphost = "mailx.rakshamail.com:587"   #initially this was used...as host

        smtphost = "mailx.rakshamail.com:587"

        # Step 5 - Declare message elements

        msg['From'] = "info@rakshatpa.com"

        msg['To'] = ','.join(to)

        msg['Cc'] = ','.join(to1)

        msg['Subject'] = 'MIS REPORT'

        # Step 6 - Add the message body to the object instance

        msg.attach(MIMEText(message, 'plain'))

        payload = MIMEBase('application', 'octate-stream', Name='MIS.xlsx')

        payload.set_payload((xl_file).read())

        # enconding the binary into base64

        encoders.encode_base64(payload)

        # add header with pdf name

        payload.add_header('Content-Decomposition', 'attachment', filename='MIS.xlsx')

        msg.attach(payload)

        # Step 7 - Create the server connection

        server = smtplib.SMTP(smtphost)

        # Step 9 - Authenticate with the server

        server.login(username, password)

        # Step 10 - Send the message

        # add receiving participants 'rohitcrk1@gmail.com'

        for id in listOfEmailIds:

            server.sendmail(msg['From'], id, msg.as_string())
            # Step 11 -
            print("Successfully sent email message to %s:" % (id))

        # Step 12 - Disconnect
        server.quit()
    except Exception as e:
        print(e)



WriteToExcel()

SendEmail()
