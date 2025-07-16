from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from email.message import EmailMessage
import smtplib

def create_invoice_pdf(customer,item,qty,price):
    total=qty*price
    filename=(f"Invoice_{customer}.pdf")
    c= canvas.Canvas(filename,pagesize=A4)
    c.drawString(100,800,"INVOICE")
    c.drawString(100,770,f"Name : {customer}")
    c.drawString(100,740,f"Item : {item}")
    c.drawString(100,710,f"Quantity : {qty} Units")
    c.drawString(100,680,f"Price : ${price}")
    c.drawString(100,650,f"Total : ${total}")
    c.save()
    return filename


def update_excel(customer,item,qty,price):
    wb=load_workbook("sales_data_u.xlsx")
    sheet=wb.active
    sheet.append([customer,item,qty,price,qty*price])
    wb.save("sales_data_u.xlsx")

def send_email_with_invoice(sender,reciver,apppass,pdf_file):
    msg=EmailMessage()
    msg['Subject']="Your Invoice!"
    msg['From']=sender
    msg['To']=reciver
    msg.set_content("Hi, please find attached your invoice.")

    with open(pdf_file,"rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=pdf_file)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender,apppass)
        smtp.send_message(msg)    

    print("Email sent!")    

customer=input("Enter Customer Name: ")
email=input("Enter Customer Email: ")
item=input("Enter Item Name: ")
qty=float(input("Enter Quantity in KG: "))
price=float(input("Enter Price Per Unit: "))


update_excel(customer,item,qty,price)
pdf_file=create_invoice_pdf(customer,item,qty,price)
send_email_with_invoice("sargaraniruddha@gmail.com",email,"qrvs mcqj sqfz lnpb",pdf_file)

