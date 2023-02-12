import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime


def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, 1)
    description__entry.delete(0,tkinter.END)
    unit_price_spinbox.delete(0, tkinter.END)
    unit_price_spinbox.insert(0, 0.0)

invoice_list =[]
def add_item():
    qty = int(qty_spinbox.get())
    desc = str(description__entry.get())
    price = float(unit_price_spinbox.get())
    line_total = float(qty*price)
    invoice_items = [qty, desc, price, line_total]
    tree.insert("", 0, values=invoice_items)
    clear_item()
    invoice_list.append(invoice_items)

def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_number__entry.delete(0,tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()


def invoice_generate():
    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get()+last_name_entry.get()
    phone = phone_number__entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal*(1 + salestax)

    doc.render({"name":name,
    "phone": phone,
    "invoice_list": invoice_list,
    "subtotal": subtotal,
    "salestax":str(salestax*100)+"%",
    "total":total})

    doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")+".docx"
    doc.save(doc_name)



window = tkinter.Tk()
window.title("Invoice Generator Form")

frame = tkinter.Frame(window)
frame.pack(padx=100, pady=30)

first_name_lable = tkinter.Label(frame, text="First Name")
first_name_lable.grid(row=0, column=0)
last_name_lable = tkinter.Label(frame, text="Last Name")
last_name_lable.grid(row=0, column=1)
phone_number__lable = tkinter.Label(frame, text="Phone Number")
phone_number__lable.grid(row=0, column=2)
qty__lable = tkinter.Label(frame, text="Qty")
qty__lable.grid(row=2, column=0)
description__lable = tkinter.Label(frame, text="Discription")
description__lable.grid(row=2, column=1)
unit_price = tkinter.Label(frame, text="Unit Price")
unit_price.grid(row=2, column=2)
add_item_button = tkinter.Button(frame, text="Add Item", command=add_item)
add_item_button.grid(row=4, column=2, pady=5)

first_name_entry = tkinter.Entry(frame) 
first_name_entry.grid(row=1, column=0)
last_name_entry = tkinter.Entry(frame)
last_name_entry.grid(row=1, column=1)
phone_number__entry= tkinter.Entry(frame)
phone_number__entry.grid(row=1, column=2)
qty_spinbox= tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)
description__entry = tkinter.Entry(frame)
description__entry.grid(row=3, column=1)
unit_price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5)
unit_price_spinbox.grid(row=3, column=2)

columns= ("qty","desc", "price", "total")
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading("qty", text="Qty")
tree.heading("desc", text="Description")
tree.heading("price", text="Unit Price")
tree.heading("total", text="Total")
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

generate_button = tkinter.Button(frame, text="Generate Invoice", command=invoice_generate)
generate_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)




window.mainloop()