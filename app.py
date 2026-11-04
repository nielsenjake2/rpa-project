# Insert your python code here

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

#q1
while True:
    directory = single_input("Paste the path for where the ProductData.xlsx, MasterInvoice.pdf, and Order files are located:")
    directory=directory.strip("'")

    if directory[-1]!='/':
        directory=directory+'/'

    files=['ProductData.xlsx', 'MasterInvoice.pdf','Order1.xlsx','Order2.xlsx', 'Order3.xlsx','Order4.xlsx']
    correct=0
    try:
        user_files = os.listdir(directory)
        missing_files=[]
        for file in files:
            if file not in user_files:
                missing_files.append(file)
        if len(missing_files)==0:
            correct=1
            break
        else:
            show_message('Error',f"Directory is missing files: {missing_files}. Please try again.")
    except:
        show_message('Error',"Incorrect File path. Try again.")

#Q2
desc = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
product_descriptions = {}
for index, row in desc.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices=pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
product_prices={}
for index, row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

#q3
def get_order_data(invoice_num, filename, directory):
    template = openpyxl.load_workbook(directory + filename, data_only=True)
    selected_sheet=template['OrderForm']

    order_data = {}

    company=selected_sheet['C15'].value
    name=selected_sheet['C14'].value
    email=selected_sheet['C13'].value
    date=selected_sheet['C9'].value.strftime('%m/%d/%Y')
    street=selected_sheet['C16'].value
    city=selected_sheet['C17'].value
    state=selected_sheet['C18'].value
    zipcode=str(selected_sheet['C19'].value)

    order_data['CustomerName'] = name
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerAddress1'] = street
    order_data['CustomerAddress2'] = f'{city}, {state} {zipcode}'
    order_data['CustomerCompany'] = company

    total=0
    row=24
    item=0

    while True:
        unit_num = selected_sheet[f'B{row}'].value
        qty = selected_sheet[f'C{row}'].value
        if unit_num is None:
            break
        desc = str(product_descriptions[unit_num])
        price = product_prices[unit_num]

        amount = price * qty
        total=total+amount

        order_data[f'Description{item}'] = desc
        order_data[f'Price{item}'] = f'${price:,.2f}'
        order_data[f'Quantity{item}'] = str(qty)
        order_data[f'Amount{item}'] = f'${amount:,.2f}'

        row =row+1
        item = item+1

    order_data['Total'] = f'${total:,.2f}'
    return order_data, name, email, date



#Q4
def create_invoice(invoice_num, order_data, directory):
    old_pdf = directory + 'MasterInvoice.pdf'
    new_pdf = directory + f'Invoice {invoice_num}.pdf'
    
    create_pdf(old_pdf, new_pdf, order_data)

#q5
def send_email(invoice_num, name, email, date, directory):
    message = EmailMessage()
    message['Subject'] = f'Your PythonTech Co. Order - Invoice: {invoice_num}'
    message['From'] = 'email@gmail.com'
    message['To'] = email

    body = f"""Dear {name},

    Attached is the invoice dated {date}. 

    Thank you for your business. 

    PythonTech Co.
    10 Python Drive
    Pleasent Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    message.set_content(body)

    path_to_file = directory + f'Invoice {invoice_num}.pdf'    
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    gmail_password_path = select_file('Secret file path I do not want online')
    gmail_password = open(gmail_password_path).read().strip()
    
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    mail_server.login('email@gmail.com', gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'email@gmail.com', '11/13/2023', directory)

filenames= os.listdir(directory)
filenames.sort()
invoice_num = 8515

#q6
for filename in filenames:
    if 'Order' in filename and '.xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num=invoice_num+1