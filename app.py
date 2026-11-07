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

while True:
        directory2 = single_input("Enter the path to the folder that contains the data.")
        directory1 = directory2.strip('"')
        directory = directory1 + '\\'
        try:
                filenames = os.listdir(directory)
                if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
                        break
                else: 
                        show_message("Error", "The path you entered is incorrect. Please try again.")
        except: 
                show_message("Error", "You entered an invalid path.")

product_descriptions = {}
product_prices = {}
des_unit = []
des_desc = []
pro_unit = []
pro_prices = []

df1 = pd.read_excel(directory+"ProductData.xlsx", sheet_name = "Descriptions")
for index, row in df1.iterrows():
    des_desc.append(row["Description"])
    des_unit.append(row['Unit number'])

for i,id in enumerate(des_unit):
    description = des_desc[i]
    product_descriptions[id] = description

df2 = pd.read_excel(directory+"ProductData.xlsx", sheet_name = "Prices")
for index, row in df2.iterrows():
    pro_prices.append(row["Price per unit"])
    pro_unit.append(row['Unit number'])

for i1,id1 in enumerate(pro_unit):
    price = pro_prices[i1]
    product_prices[id1] = price


df = get_pdf('MasterInvoice.pdf')


def get_order_data(invoice_num, filename, directory):

    template = openpyxl.load_workbook(directory+filename)
    
    selected_sheet = template['OrderForm']
    
    
    name = selected_sheet['C14'].value 
    email = selected_sheet['C13'].value 
    date = selected_sheet['C9'].value 
    date_str = date.strftime('%m/%d/%Y')
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value 
    state = selected_sheet['C18'].value 
    zipcode = selected_sheet['C19'].value 
    customeraddress = f'{address}, {city}, {state} {zipcode}'
    customercompany = selected_sheet['C15'].value 

    order_dict = {}
    order_dict['InvoiceNum'] = str(invoice_num)
    order_dict['Date'] = date_str
    order_dict['CustomerAddress1'] = customeraddress
    order_dict['CustomerName'] = name
    order_dict['CustomerCompany'] = customercompany  
    order_dict['CustomerAddress2'] = email  

    T = 0 

    UN0 = selected_sheet['B24'].value
    Q0 = selected_sheet['C24'].value 
    if UN0 is not None and Q0 is not None:
       P0 = product_prices[UN0] 
       A0 = Q0 * P0
       T = T + A0
       order_dict['Price0'] = f'${P0:,.2f}'
       order_dict['Quantity0'] = str(Q0)
       order_dict['Amount0'] = f'${A0:,.2f}'
       order_dict['Description0'] = product_descriptions[UN0]

    UN1 = selected_sheet['B25'].value
    Q1 = selected_sheet['C25'].value 
    if UN1 is not None and Q1 is not None:
        P1 = product_prices[UN1]
        A1 = Q1 * P1
        T = T + A1
        order_dict['Price1'] = f'${P1:,.2f}'
        order_dict['Quantity1'] = str(Q1)
        order_dict['Amount1'] = f'${A1:,.2f}'
        order_dict['Description1'] = product_descriptions[UN1]

    UN2 = selected_sheet['B26'].value    
    Q2 = selected_sheet['C26'].value 
    if UN2 is not None and Q2 is not None:
        P2 = product_prices[UN2]
        A2 = Q2 * P2
        T = T + A2
        order_dict['Price2'] = f'${P2:,.2f}'
        order_dict['Quantity2'] = str(Q2)
        order_dict['Amount2'] = f'${A2:,.2f}'
        order_dict['Description2'] = product_descriptions[UN2]
    
    order_dict['Total'] = f'${T:,.2f}'

    return order_dict, name, email, date_str

order_data, name, email, date = get_order_data (8515, "Order1.xlsx", directory)

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)

create_invoice('8515', order_data, directory)


gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_emails(invoice_num, name, email, date, directory):

    gmail_user = "mengvimean03@gmail.com"

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'


    NAME = name

    body = f"""Dear {NAME},
        
Attached is your invoice dated {date}.
    
Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""


    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = f'{directory}Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) 
            

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)


    mail_server.send_message(message)
    mail_server.quit()

send_emails (8515, 'Thomas Holt','vimean03@byu.edu', '11/12/2023', directory)



filenames = os.listdir(directory)
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        invoice_num
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_emails(invoice_num, name, email, date, directory)
 
        invoice_num = invoice_num + 1
