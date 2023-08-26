from flask import Flask, request, Response
import pymongo
import pandas as pd
from openpyxl import Workbook
import os
import io
import logging

app = Flask(__name__)

# Set up logging
logging.basicConfig(level=logging.DEBUG)

# Determine the uploads directory path
uploads_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
if not os.path.exists(uploads_dir):
    os.makedirs(uploads_dir)

# MongoDB connection setup
mongo_uri = os.getenv("MONGO_URI")
if not mongo_uri:
    logging.error("MONGO_URI environment variable is not set.")
client = pymongo.MongoClient(mongo_uri)
db = client["Plaschema"]
collection = db["subscriptions"]

@app.route('/export-excel', methods=['GET'])
def export_excel():
    try:
        # Fetch data from MongoDB collection
        cursor = collection.find({})
        data = list(cursor)

        # Convert data to a DataFrame
        df = pd.DataFrame(data)

        # Create an in-memory Excel workbook
        excel_buffer = io.BytesIO()
        excel_writer = pd.ExcelWriter(excel_buffer, engine='openpyxl')
        wb = excel_writer.book

        # Convert the DataFrame to an Excel sheet
        df.to_excel(excel_writer, sheet_name='Sheet1', index=False)

        # Save the Excel workbook to the buffer
        wb.save(excel_buffer)

        # Save the buffer content to a file in the uploads directory
        file_path = os.path.join(uploads_dir, 'Plaschema_data.xlsx')
        with open(file_path, 'wb') as f:
            f.write(excel_buffer.getvalue())

        # Reset the buffer position and create a response
        excel_buffer.seek(0)
        response = Response(excel_buffer.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response.headers['Content-Disposition'] = 'attachment; filename=Plaschema_data.xlsx'

        # Craft a congratulatory message
        congratulatory_message = "Congratulations! Excel file has been generated."

        return congratulatory_message

    except Exception as e:
        logging.error(str(e))
        return "An error occurred while generating the Excel file.", 500


