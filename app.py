while True:
    directory = single_input("Please input the directory where your files are located.")
    directory = directory.strip('"').strip("'")
    if directory[-1] == '/':
        break
    else:
        directory = directory + '/'
    filenames = os.listdir(directory)
    if "MasterInvoice.pdf" in filenames and "Order1.xlsx" in filenames and "ProductData.xlsx" in filenames and "Order2.xlsx" in filenames and "Order3.xlsx" in filenames and "Order4.xlsx" in filenames:
        break
    else:
        show_message('Error', 'Please enter the directory that contains necessary files')
            
show_message('Thank you','This file contains the correct files')
print(directory) 
print(filenames)
    
filenames = os.listdir(directory)
filepath = directory + 'ProductData.xlsx'
prices = pd.read_excel(filepath, sheet_name = 'Prices')
descriptions = pd.read_excel(filepath, sheet_name = 'Descriptions')
product_descriptions = {}
product_prices = {}


for index,row in descriptions.iterrows():
    descript = row['Description']
    unit_num = row['Unit number']
    product_descriptions[unit_num] = descript


for index,row in prices.iterrows():
    price = row['Price per unit']
    unit_num = row['Unit number']
    product_prices[unit_num] = price

print(product_descriptions)
print(product_prices)
df = get_pdf('MasterInvoice.pdf')
df
order_data= {}

def get_order_data(invocie_num,filename,directory):
    path = directory + filename
    order_file = openpyxl.load_workbook(path, data_only=True)
    selected_sheet = order_file['OrderForm']
    date = (selected_sheet['C9'].value).strftime('%m/%d/%Y')
    order_data['InvoiceNum'] = str(invocie_num)
    order_data['Date'] = date
    # fix date
     
    order_data['CustomerAddress1'] = str(selected_sheet['C16'].value)
    order_data['CustomerName'] = str(selected_sheet['C14'].value)
    order_data['CustomerCompany'] = str(selected_sheet['C15'].value)
    order_data['CustomerAddress2'] = f'{str(selected_sheet['C17'].value)},{str(selected_sheet['C18'].value)}, {str(selected_sheet['C19'].value)}'

    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value

    order_data['Price0'] = str(product_prices[selected_sheet['B24'].value])
    order_data['Quantity0'] = str(selected_sheet['C24'].value)
    amount0 =product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value
    order_data['Amount0']= f'${amount0:,.2f}'
    order_data['Description0'] = str(product_descriptions[selected_sheet['B24'].value])
    total = amount0

    if pd.isnull(selected_sheet['B25'].value) == False:
        order_data['Price1'] = str(product_prices[selected_sheet['B25'].value])
        order_data['Quantity1'] = str(selected_sheet['C25'].value)
        amount1 =product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value
        order_data['Amount1']= f'${amount0:,.2f}'
        order_data['Description1'] = str(product_descriptions[selected_sheet['B25'].value])
        total = total + amount1

    if pd.isnull(selected_sheet['B26'].value) == False:
        order_data['Price2'] = str(product_prices[selected_sheet['B26'].value])
        order_data['Quantity2'] = str(selected_sheet['C26'].value)
        amount2 = product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value
        order_data['Amount2']= f'${amount0:,.2f}'
        order_data['Description2'] = str(product_descriptions[selected_sheet['B26'].value])
        total = total + amount2

    
    order_data['Total'] = f'${total:,.2f}'
    return order_data, name, email, date
order_data, name, email, date = get_order_data(8515,'Order1.xlsx',directory) 
# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice{invoice_num}.pdf', order_data)
create_invoice(8515,order_data,directory)
# INSERT ANSWER FOR QUESTION 5 HERE
gmail_user = 'clarabelskopec@gmail.com'
gmail_password_path = select_file('Select your gmail password file.')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num,name,email,date,directory):



    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

# Body of the email

    
    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84602
Phone: (801) 555-2372
Email: sales@pythontechco.com"""



    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)


    path_to_file = directory+f'Invoice{invoice_num}.pdf' 
    message = add_attachment(path_to_file, message) 

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)


    mail_server.send_message(message)
    mail_server.quit()

send_email(8515,'Thomas Holt','clarabelskopec@gmail.com','11/13/2023',directory)
# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames.sort()
invoice_num= 8515
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        try:
            order_data, name, email, date = get_order_data(invoice_num,filename,directory)
            create_invoice(invoice_num,order_data,directory)
            send_email(invoice_num,name,email,date,directory)
            invoice_num = invoice_num +1
        except Exception as e:
            print(e)
            pass