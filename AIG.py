import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import re

invoice_list = []
invoice_count = 1

def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")

def add_item():
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty * price
    invoice_item = [qty, desc, price, line_total]
    tree.insert('', 0, values=invoice_item)
    clear_item()
    
    invoice_list.append(invoice_item)

def new_invoice():
    global invoice_count
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    email_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    
    invoice_list.clear()
    invoice_count += 1

def validate_email(email):
    if re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email):
        return True
    else:
        return False

def validate_phone(phone):
    if re.match(r"^[0-9]{10}$", phone):
        return True
    else:
        return False

def send_email(recipient, doc_name):
    sender_email = ""  # Update with your email
    sender_password = ""  # Update with your application specific password(Refer Readme File)
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient
    msg['Subject'] = "Your Invoice"
    
    body = "Please find attached your invoice."
    msg.attach(MIMEText(body, 'plain'))

    with open(doc_name, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name=doc_name)
        part['Content-Disposition'] = f'attachment; filename="{doc_name}"'
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587) 
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient, text)
        server.quit()
        messagebox.showinfo("Email Sent", "The invoice has been sent to the customer.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send email. Error: {e}")

def generate_invoice():
    global invoice_count
    doc = DocxTemplate("")  # Update with your template path
    name = first_name_entry.get() + " " + last_name_entry.get()
    phone = phone_entry.get()
    email = email_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal * (1 + salestax)

    doc.render({
        "name": name, 
        "phone": phone,
        "invoice_list": invoice_list,
        "subtotal": subtotal,
        "salestax": str(salestax * 100) + "%",
        "total": total
    })
    
    doc_name = f"{invoice_count}_new_invoice_{name.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')}.docx"
    doc.save(doc_name)
    
    send_email(email, doc_name)
    
    new_invoice()

def validate_and_generate():
    email = email_entry.get()
    phone = phone_entry.get()
    
    if not validate_email(email):
        messagebox.showerror("Invalid Email", "Please enter a valid email address.")
        return
    
    if not validate_phone(phone):
        messagebox.showerror("Invalid Phone Number", "Please enter a 10-digit phone number.")
        return
    
    generate_invoice()

window = tkinter.Tk()
window.title("Automatic Invoice Generator")

style = ttk.Style()
style.configure("TButton", font=("Arial", 12, "bold"))
style.map("TButton",
          foreground=[('!disabled', 'black')],
          background=[('!disabled', '#4a90e2'), ('pressed', '!disabled', '#4a90e2')],
          )

style.configure("Add.TButton", background="#4a90e2", foreground="black")
style.configure("Generate.TButton", background="#4a90e2", foreground="black")
style.configure("New.TButton", background="#7f8c8d", foreground="black")

frame = tkinter.Frame(window, bg="#ecf0f1")
frame.pack(padx=20, pady=10)

first_name_label = tkinter.Label(frame, text="First Name", bg="#ecf0f1", font=("Arial", 10))
first_name_label.grid(row=0, column=0, padx=5, pady=5)
last_name_label = tkinter.Label(frame, text="Last Name", bg="#ecf0f1", font=("Arial", 10))
last_name_label.grid(row=0, column=1, padx=5, pady=5)

first_name_entry = tkinter.Entry(frame, font=("Arial", 10))
last_name_entry = tkinter.Entry(frame, font=("Arial", 10))
first_name_entry.grid(row=1, column=0, padx=5, pady=5)
last_name_entry.grid(row=1, column=1, padx=5, pady=5)

phone_label = tkinter.Label(frame, text="Phone", bg="#ecf0f1", font=("Arial", 10))
phone_label.grid(row=0, column=2, padx=5, pady=5)
phone_entry = tkinter.Entry(frame, font=("Arial", 10))
phone_entry.grid(row=1, column=2, padx=5, pady=5)

email_label = tkinter.Label(frame, text="Email", bg="#ecf0f1", font=("Arial", 10))
email_label.grid(row=2, column=0, padx=5, pady=5)
email_entry = tkinter.Entry(frame, font=("Arial", 10))
email_entry.grid(row=3, column=0, padx=5, pady=5)

qty_label = tkinter.Label(frame, text="Qty", bg="#ecf0f1", font=("Arial", 10))
qty_label.grid(row=4, column=0, padx=5, pady=5)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100, font=("Arial", 10))
qty_spinbox.grid(row=5, column=0, padx=5, pady=5)

desc_label = tkinter.Label(frame, text="Description", bg="#ecf0f1", font=("Arial", 10))
desc_label.grid(row=4, column=1, padx=5, pady=5)
desc_entry = tkinter.Entry(frame, font=("Arial", 10))
desc_entry.grid(row=5, column=1, padx=5, pady=5)

price_label = tkinter.Label(frame, text="Unit Price", bg="#ecf0f1", font=("Arial", 10))
price_label.grid(row=4, column=2, padx=5, pady=5)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5, font=("Arial", 10))
price_spinbox.grid(row=5, column=2, padx=5, pady=5)

add_item_button = ttk.Button(frame, text="Add item", style="Add.TButton", command=add_item)
add_item_button.grid(row=6, column=2, pady=10, padx=10)

columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings", height=8)
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Unit Price')
tree.heading('total', text="Total")
tree.grid(row=7, column=0, columnspan=3, padx=20, pady=10)

save_invoice_button = ttk.Button(frame, text="Generate Invoice", style="Generate.TButton", command=validate_and_generate)
save_invoice_button.grid(row=8, column=0, columnspan=3, sticky="news", padx=20, pady=10)
new_invoice_button = ttk.Button(frame, text="New Invoice", style="New.TButton", command=new_invoice)
new_invoice_button.grid(row=9, column=0, columnspan=3, sticky="news", padx=20, pady=10)

window.mainloop()
