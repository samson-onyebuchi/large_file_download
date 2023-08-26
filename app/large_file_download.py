from flask import Flask, request, Response
import pymongo
import pandas as pd
from openpyxl import Workbook
import os 

app = Flask(__name__)

# MongoDB connection setup
client = pymongo.MongoClient(os.getenv("MONGO_URI"))
db = client["Plaschema"]
collection = db["subscriptions"]

@app.route('/export-excel', methods=['GET'])
def export_excel():
    # Fetch data from MongoDB collection
    cursor = collection.find({})
    data = list(cursor)

    # Convert data to a DataFrame
    df = pd.DataFrame(data)

    # Create a new Excel workbook
    excel_writer = pd.ExcelWriter('Plaschema data.xlsx', engine='openpyxl')
    wb = excel_writer.book

    # Convert the DataFrame to an Excel sheet
    df.to_excel(excel_writer, sheet_name='Sheet1', index=False)

    # Save the Excel workbook
    wb.save('Plaschema data.xlsx')

    # Craft a congratulatory message
    congratulatory_message = "Congratulations! Excel file has been generated."

    return congratulatory_message


