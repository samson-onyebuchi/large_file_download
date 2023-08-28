from flask import Flask, request
import pymongo
import pandas as pd
import os
import smtplib
import zipfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
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

def send_zip_email(data, from_email, to_email):
    df = pd.DataFrame(data)
    zip_filename = 'plaschema_data.zip'
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        csv_data = df.to_csv(index=False)
        zipf.writestr('plaschema_data.csv', csv_data)
    
    congratulatory_message = f"This is data from {db.name}, {collection.name}."
    subject = 'Plaschema Data ZIP File'
    body = congratulatory_message

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    part = MIMEBase('application', 'zip')
    with open(zip_filename, 'rb') as zip_file:
        part.set_payload(zip_file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{zip_filename}"')
    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, os.getenv("EMAIL_PASSWORD"))
    text = msg.as_string()
    server.sendmail(from_email, to_email, text)
    server.quit()

    os.remove(zip_filename)

@app.route('/export-email', methods=['POST'])
def export_email():
    try:
        data = request.get_json()
        from_email = data.get('from_email')
        to_email = data.get('to_email')

        cursor = collection.find({})
        data = list(cursor)

        send_zip_email(data, from_email, to_email)

        return "ZIP file with data sent successfully."

    except Exception as e:
        error_message = f"An error occurred while processing the request: {str(e)}"
        return error_message, 500
