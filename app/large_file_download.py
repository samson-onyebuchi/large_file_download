from flask import Flask, request, Response
import pymongo
import pandas as pd
from openpyxl import Workbook
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

app = Flask(__name__)

# MongoDB connection setup
try:
    client = pymongo.MongoClient(os.getenv("MONGO_URI"))
    db = client["Plaschema"]
    collection = db["subscriptions"]
except Exception as e:
    error_message = f"Error connecting to the database: {str(e)}"
    print(error_message)

#@app.route('/export-excel', methods=['GET'])
@app.route('/export-excel', methods=['POST'])
def export_excel():
    try:
        # # Email configuration
        # from_email = 'nnadisamson63@gmail.com'
        # to_email = 'nnadionyebuchi33@gmail.com'
        # subject = 'Plaschema Data Excel File'
        # body = congratulatory_message
        data = request.get_json()  
        from_email = data.get('from_email')  
        to_email = data.get('to_email')  

        cursor = collection.find({})
        data = list(cursor)

        df = pd.DataFrame(data)

        # Create a new Excel workbook
        excel_writer = pd.ExcelWriter('Plaschema data.xlsx', engine='openpyxl')
        wb = excel_writer.book

        df.to_excel(excel_writer, sheet_name='Sheet1', index=False)

        # Save the Excel workbook
        wb.save('Plaschema data.xlsx')

        # Craft a congratulatory message
        congratulatory_message = f"This is an excel file from {db}, {collection}."

        # Email configuration
        subject = 'Plaschema Data Excel File'
        body = congratulatory_message

        # Create the MIME object
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Attach the Excel file
        filename = 'Plaschema data.xlsx'
        attachment = open(filename, 'rb')
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)

        # Connect to the SMTP server and send the email
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, (os.getenv("EMAIL_PASSWORD")))
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()

        return congratulatory_message

    except Exception as e:
        error_message = f"An error occurred while processing the request: {str(e)}"
        print(error_message)
        return error_message, 500
    
