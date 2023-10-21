import tkinter
from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
from tkcalendar import DateEntry
import datetime
from tkinter import messagebox

def clear_item():
    kuantitas_spinbox.delete(0, tkinter.END)
    kuantitas_spinbox.insert(0, "1")
    deskripsiBarang_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")

invoice_list = []
def add_item():
    kuantitas = int(kuantitas_spinbox.get())
    deskripsiBarang = deskripsiBarang_entry.get()
    price = float(price_spinbox.get())
    if not kuantitas or not deskripsiBarang or not price:
        messagebox.showwarning("Peringatan", "Mohon isi terlebih dahulu datanya!")
        return
    line_total = kuantitas*price
    invoice_item = [kuantitas, deskripsiBarang, price, line_total]
    tree.insert('',0, values=invoice_item)
    clear_item()
    
    invoice_list.append(invoice_item)
    subtotal = sum([item[3] for item in invoice_list])
    subtotal_label.config(text=f'Total: {subtotal:.2f}')

def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    no_invoice_entry.delete(0, tkinter.END)
    tgl_order_entry.set_date(datetime.date.today())
    tgl_jatuhtempo_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    
    invoice_list.clear()
    subtotal_label.config(text='Total: 0.0')
    
def generate_invoice():
    if len(invoice_list) == 0:
        messagebox.showwarning("Peringatan", "Tolong isi data terlebih dahulu!")
        return
    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get()+last_name_entry.get()
    phone = phone_entry.get()
    tgl_order = tgl_order_entry.get()
    tgl_jatuhtempo = tgl_jatuhtempo_entry.get()
    no_invoice = no_invoice_entry.get()
    subtotal = sum(item[3] for item in invoice_list) 
    salestax = 0.1
    total = subtotal*(1+salestax)
    
    doc.render({"name":name, 
            "phone":phone,
            "tgl_order": tgl_order,
            "tgl_jatuhtempo": tgl_jatuhtempo,
            "no_invoice" : no_invoice,
            "invoice_list": invoice_list,
            "subtotal":subtotal,
            "salestax":str(salestax*100)+"%",
            "total":total})
    
    doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)
    
    messagebox.showinfo("Invoice Complete", "Invoice Complete")
    
    new_invoice()

def delete_item():
    selected_items = tree.selection()
    for selected_item in selected_items:
        item_values = tree.item(selected_item)['values']
        for i in invoice_list:
            if i[1] == item_values[1]:
                invoice_list.remove(i)
                break
        tree.delete(selected_item)
    subtotal = sum(item[3] for item in invoice_list)
    subtotal_label.config(text=f'Total: {subtotal:.2f}')

def on_enter(event):
    add_item()

window = tkinter.Tk()
window.title("Aplikasi Penjualan PT Sheila Jaya")

style = ttk.Style(window)
style.theme_use("clam")

frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)

subtotal_label = tkinter.Label(frame, text='Total: 0.0', font=("Arial", 12, "bold"))
subtotal_label.grid(row=9, column=2)

title_label = tkinter.Label(frame, text="Aplikasi Penjualan PT Sheila Jaya", font=("Arial", 16, "bold"))
title_label.grid(row=0, column=0, columnspan=3, pady=10)

first_name_label = tkinter.Label(frame, text="First Name")
first_name_label.grid(row=1, column=0)
last_name_label = tkinter.Label(frame, text="Last Name")
last_name_label.grid(row=1, column=1)

first_name_entry = tkinter.Entry(frame)
last_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=2, column=0)
last_name_entry.grid(row=2, column=1)

phone_label = tkinter.Label(frame, text="Phone")
phone_label.grid(row=1, column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=2, column=2)

tgl_order_label = tkinter.Label(frame, text="Tanggal Order")
tgl_order_label.grid(row=3, column=0)
tgl_order_entry = DateEntry(frame, date_pattern='dd/mm/yyyy')
tgl_order_entry.grid(row=4, column=0)

no_invoice_label = tkinter.Label(frame, text="Nomor Invoice")
no_invoice_label.grid(row=3, column=1)
no_invoice_entry = tkinter.Entry(frame)
no_invoice_entry.grid(row=4, column=1)

tgl_jatuhtempo_label = tkinter.Label(frame, text="Tanggal Jatuh Tempo")
tgl_jatuhtempo_label.grid(row=3, column=2)
tgl_jatuhtempo_entry = DateEntry(frame, date_pattern='dd/mm/yyyy')
tgl_jatuhtempo_entry.grid(row=4, column=2)

kuantitas_label = tkinter.Label(frame, text="Kuantitas")
kuantitas_label.grid(row=5, column=0)
kuantitas_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
kuantitas_spinbox.grid(row=6, column=0)

deskripsiBarang_label = tkinter.Label(frame, text="Deskripsi Barang")
deskripsiBarang_label.grid(row=5, column=1)
deskripsiBarang_entry = tkinter.Entry(frame)
deskripsiBarang_entry.grid(row=6, column=1)

price_label = tkinter.Label(frame, text="Harga Satuan")
price_label.grid(row=5, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=1000000.0, increment=500.0)
price_spinbox.grid(row=6, column=2)

add_item_button = tkinter.Button(frame, text="Add item", command=add_item)
add_item_button.grid(row=7, column=1, pady=5)
add_item_button.bind('<Return>', on_enter)

delete_item_button = tkinter.Button(frame, text = "Delete item", command = delete_item)
delete_item_button.grid(row=7, column=2, pady=5)

columns = ('kuantitas', 'deskripsiBarang', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('kuantitas', text='Kuantitas')
tree.heading('deskripsiBarang', text='Deskripsi Barang')
tree.heading('price', text='Harga Satuan')
tree.heading('total', text="Total")

tree.grid(row=8, column=0, columnspan=3, padx=20, pady=10)
style = ttk.Style()
style.configure("Treeview.Heading", background="#DEB887", foreground="black")

save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=10, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=11, column=0, columnspan=3, sticky="news", padx=20, pady=5)

window.configure(bg='#DEB887')
frame.configure(bg='#FFE4C4')
title_label.configure(bg='#FFE4C4')
first_name_label.configure(bg='#FFE4C4')
last_name_label.configure(bg='#FFE4C4')
phone_label.configure(bg='#FFE4C4')
tgl_order_label.configure(bg='#FFE4C4')
no_invoice_label.configure(bg='#FFE4C4')
tgl_jatuhtempo_label.configure(bg='#FFE4C4')
kuantitas_label.configure(bg='#FFE4C4')
deskripsiBarang_label.configure(bg='#FFE4C4')
price_label.configure(bg='#FFE4C4')
add_item_button.configure(bg='#DEB887', foreground="black")
delete_item_button.configure(bg='#DEB887', foreground="black")
save_invoice_button.configure(bg='#DEB887', foreground="black")
new_invoice_button.configure(bg='#DEB887', foreground="black")
title_label.configure(bg='#FFE4C4')
subtotal_label.configure(borderwidth=2, relief="ridge")

kuantitas_spinbox.focus()
deskripsiBarang_entry.focus()
price_spinbox.focus()
window.bind('<Return>', on_enter)

image = tkinter.PhotoImage(file="D:/2. SHEILA KULIAH TELKOM/TUBES ALPRO SEMESTER 2_UAS_SHEILA DENINTA MAHARANI_1202223324/APK INVOICE GENERATOR PENJUALAN_SHEILA DENINTA MAHARANI_1202223324/logo_sheila_jaya1.png")
image = image.subsample(4) 
logo_label = tkinter.Label(frame, image=image)
logo_label.grid(row=0, column=0)
logo_label.configure(bg='#FFE4C4')

window.mainloop()