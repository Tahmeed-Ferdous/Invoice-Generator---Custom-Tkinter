import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
from docxtpl import DocxTemplate
from datetime import datetime
from tkinter import messagebox
from tkcalendar import Calendar
import random
import os
from docx2pdf import convert
import webbrowser

invoice_list = []
def add_item():
        article = article_entry.get()
        color = color_entry.get()
        five = int(five_entry.get())
        six = int(six_entry.get())
        seven = int(seven_entry.get())
        eight = int(eight_entry.get())
        nine = int(nine_entry.get())
        ten = int(ten_entry.get())
        rate=int(rate_entry.get())
        totalpair=five+six+seven+eight+nine+ten
        total=rate*totalpair

        invoice_item2 = [article, color, five, six, seven, eight, nine, ten, rate, totalpair, total]
        tree.insert('', 0, values=invoice_item2)

        invoice_list.append(invoice_item2)
        toast()
        # article_entry.delete(0, ctk.END)
        # color_entry.delete(0, ctk.END)
        # five_entry.delete(0, ctk.END)
        # six_entry.delete(0, ctk.END)
        # seven_entry.delete(0, ctk.END)
        # eight_entry.delete(0, ctk.END)
        # nine_entry.delete(0, ctk.END)
        # ten_entry.delete(0, ctk.END)

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        file_entry.delete(0, ctk.END)
        file_entry.insert(0, file_path)

def convert_to_pdf():
    docx_path = file_entry.get()
    if not docx_path or not docx_path.endswith('.docx'):
        messagebox.showwarning("Invalid File", "Please select a valid .docx file")
        return

    try:
        output_folder = "PDFs"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        pdf_path = os.path.join(output_folder, os.path.basename(docx_path).replace('.docx', '.pdf'))
        convert(docx_path, pdf_path)
        
        messagebox.showinfo("Success", f"PDF saved successfully at {pdf_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def invoicegen():
    ran=random.randint(100, 1000)
    invoicenumber_entry.delete(0, ctk.END)
    invoicenumber_entry.insert(0,'PF/BEL/DC/24-'+str(ran))

def new_invoice():
    invoicegen()
    purchaseorder_entry.delete(0, ctk.END)
    article_entry.delete(0, ctk.END)
    color_entry.delete(0, ctk.END)
    tree.delete(*tree.get_children())
    invoice_list.clear()
    delete_all_items()
    new()
    totalpair_entry.delete(0, ctk.END)
    total_entry.delete(0, ctk.END)
    grandtotal_entry.delete(0, ctk.END)

def clear():
    purchaseorder_entry.delete(0, ctk.END)
    placeofloading_entry.delete(0, ctk.END)
    placeofdelivery_entry.delete(0, ctk.END)
    description_entry.delete(0, ctk.END)


def generate_invoice():
    doc = DocxTemplate("templates/Invoice Template(BILL).docx")
    doc2 = DocxTemplate("templates/Invoice Template(DELIVERY CHALAN).docx")
    invoice_number = invoicenumber_entry.get()
    purchase_order = purchaseorder_entry.get()
    date1 = dates_entry.get()
    date2 = dates_entry.get()
    delivery_to = deliveryto_entry.get()
    place_of_loading = placeofloading_entry.get()
    place_of_delivery = placeofdelivery_entry.get()
    description = description_entry.get()
    totalpair=totalpair_entry.get()
    total=total_entry.get()
    vat=int(int(total_entry.get())*0.15)
    grand_total=int(vat)+int(total)
    packing=str(int(totalpair)//12)+'carton'

    doc.render({"invoice_number": invoice_number,
                "purchase_order": purchase_order,
                "date1": date1,
                "date2": date2,
                "delivery_to": delivery_to,
                "place_of_loading": place_of_loading,
                "place_of_delivery":place_of_delivery,
                "invoice_list": invoice_list,
                "description": description,
                "totalpair": totalpair,
                "total":total,
                "vat":vat,
                "grand_total":grand_total})

    doc2.render({"invoice_number": invoice_number,
               "purchase_order": purchase_order,
               "date1": date1,
               "date2": date2,
               "delivery_to": delivery_to,
               "place_of_loading": place_of_loading,
               "place_of_delivery":place_of_delivery,
               "invoice_list": invoice_list,
               "description": description,
               "totalpair": totalpair,
               "packing":packing})

    current_date = datetime.now().strftime("%d-%m-%Y")
    
    doc_name ="invoices/"+'BILL-'+str(purchaseorder_entry.get())+'--'+str(current_date)+ ".docx"
    doc.save(doc_name)

    doc_name ="invoices/"+'DELCHL-'+str(purchaseorder_entry.get())+'--'+str(current_date)+ ".docx"
    doc2.save(doc_name)

    messagebox.showinfo("Invoice Complete", "Invoice Complete")
    new_invoice()

def toggle_mode():
    current_mode = ctk.get_appearance_mode()
    if current_mode == "Light":
        ctk.set_appearance_mode("dark")
        toggle_button.configure(text="Light Mode")
    else:
        ctk.set_appearance_mode("light")
        toggle_button.configure(text="Dark Mode")

def toast():
    show_toast("Item Added in the Table", duration=3000)

def new():
    show_toast('New Invoice Created.', duration=3000)

def show_toast(message, duration=3000):
    toast = ctk.CTkToplevel()
    toast.geometry("200x60+0+1600") 
    toast.overrideredirect(True)
    toast.attributes("-topmost", True) 

    label = ctk.CTkLabel(toast, text=message)
    label.pack(expand=True, padx=5, pady=5)

    toast.after(duration, toast.destroy)


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue") 

window = ctk.CTk()
window.title("Invoice Generator Form")
window.geometry('1360x850+0+0')
window.iconbitmap('images/tracking.ico')
window.bind('<Escape>', lambda event: window.quit())


frame = ctk.CTkFrame(window)
frame.pack(padx=80, pady=25)

toggle_button = ctk.CTkButton(frame, text="Switch Mode", command=toggle_mode, corner_radius=50, width=10, height=25,font=ctk.CTkFont(size=10), fg_color='#686D76', hover_color='grey')
toggle_button.grid(row=0, column=1, padx=10,pady=20)

GREEN='#2DA571'
penguinfootwear = ctk.CTkLabel(frame, text="Penguin Footwear Billing Receipt Generator", fg_color='#1D6AA4', corner_radius=20)
penguinfootwear.grid(row=0, column=0,columnspan=2, sticky='nw',padx=30, pady=18)

###########################################################################  Top Part - Invoice Purchase and Date
invoicenumber = ctk.CTkLabel(frame, text="Invoice Number")
invoicenumber.grid(row=1, column=0)
purchaseorder = ctk.CTkLabel(frame, text="Purchase Order")
purchaseorder.grid(row=1, column=1)

invoicenumber_entry = ctk.CTkEntry(frame)
purchaseorder_entry = ctk.CTkEntry(frame)
invoicegen()
invoicenumber_entry.grid(row=2, column=0)
purchaseorder_entry.grid(row=2, column=1)

def insert_current_date():
    current_date = datetime.now().strftime("%d-%m-%Y")  # Format the date as YYYY-MM-DD
    dates_entry.delete(0, "end")  # Clear the entry
    dates_entry.insert(0, current_date)  # Insert the current date

dates = ctk.CTkLabel(frame, text="Date:")
dates.grid(row=1, column=2)
dates_entry = ctk.CTkEntry(frame)
dates_entry.grid(row=2, column=2)

insert_current_date()

########################################################################## Top 2 - Delievery place of loading and place of delivery

def combobox_callback(choice):
    print("combobox dropdown clicked:", choice)

deliveryto = ctk.CTkLabel(frame, text="Delivery to")
deliveryto.grid(row=3, column=0)

deliveryto_entry = ctk.CTkComboBox(frame, values=["Apex", "Bay Emporium", 'Bata'], command=combobox_callback)
deliveryto_entry.set("Apex")
deliveryto_entry.grid(row=4, column=0)

placeofloading = ctk.CTkLabel(frame, text="Place of Loading")
placeofloading.grid(row=3, column=1)
placeofloading_entry = ctk.CTkEntry(frame)
placeofloading_entry.grid(row=4, column=1)

placeofdelivery = ctk.CTkLabel(frame, text="Place of delivery")
placeofdelivery.grid(row=3, column=2)
placeofdelivery_entry = ctk.CTkEntry(frame)
placeofdelivery_entry.grid(row=4, column=2)

######################################################################## Right part - Article color and sizes

article = ctk.CTkLabel(frame, text="Article")
article.grid(row=0, column=3, sticky='s')
article_entry = ctk.CTkEntry(frame)
article_entry.grid(row=1, column=3, pady=2)

color = ctk.CTkLabel(frame, text="Color")
color.grid(row=0, column=4, sticky='s')
color_entry = ctk.CTkEntry(frame)
color_entry.grid(row=1, column=4,pady=2)

sizes = ctk.CTkLabel(frame, text="Sizes")
sizes.grid(row=2, column=3, sticky='n', columnspan=6, pady=5)

five = ctk.CTkLabel(frame, text="5 :")
five.grid(row=3, column=3, sticky='nw', padx=5)
five_entry = ctk.CTkEntry(frame)
five_entry.insert(0, "0")
five_entry.grid(row=4, column=3, padx=2)

six = ctk.CTkLabel(frame, text="6 :")
six.grid(row=3, column=4, sticky='nw', padx=15)
six_entry = ctk.CTkEntry(frame)
six_entry.insert(0, "0")
six_entry.grid(row=4, column=4, padx=2, pady=2)

seven = ctk.CTkLabel(frame, text="7 :")
seven.grid(row=5, column=3, sticky='nw', padx=5)
seven_entry = ctk.CTkEntry(frame)
seven_entry.insert(0, "0")
seven_entry.grid(row=6, column=3, padx=2)

eight = ctk.CTkLabel(frame, text="8 :")
eight.grid(row=5, column=4, sticky='nw', padx=15)
eight_entry = ctk.CTkEntry(frame)
eight_entry.insert(0, "0")
eight_entry.grid(row=6, column=4, padx=2, pady=2)

nine = ctk.CTkLabel(frame, text="9 :")
nine.grid(row=7, column=3, sticky='nw', padx=5)
nine_entry = ctk.CTkEntry(frame)
nine_entry.insert(0, "0")
nine_entry.grid(row=8, column=3, padx=2)

ten = ctk.CTkLabel(frame, text="10 :")
ten.grid(row=7, column=4, sticky='nw', padx=15)
ten_entry = ctk.CTkEntry(frame)
ten_entry.insert(0, "0")
ten_entry.grid(row=8, column=4, padx=15, pady=2)

###################################################################################  Right-mid Rate and description

frame1 = ctk.CTkFrame(frame, width=50, height=400)
frame1.grid(row=9, column=3, columnspan=3, rowspan=4, sticky='nw', pady=10)

rate = ctk.CTkLabel(frame1, text="Rate (pair) :")
rate.grid(row=0, column=0, sticky='nw', padx=20, pady=10)
rate_entry = ctk.CTkEntry(frame1)
rate_entry.insert(0, '0')
rate_entry.grid(row=0, column=0, pady=10, padx=40, sticky='e')

description = ctk.CTkLabel(frame1, text="Description :")
description.grid(row=1, column=0, sticky='nw', padx=20, pady=10)
description_entry = ctk.CTkEntry(frame1)
description_entry.grid(row=1, column=0, pady=10, padx=40, sticky='e')

###########################################################################  Right-mid Add Total and Delete Button

add_item_button = ctk.CTkButton(frame1, text="Add item", command=add_item, width=200, height=50)
add_item_button.grid(row=2, column=0, pady=10, padx=47)
window.bind('<Return>', lambda event: add_item_button.invoke())

def delete_item():
    selected_item = tree.selection()
    if selected_item:
        tree.delete(selected_item)

def delete_all_items():
    for item in tree.get_children():
        tree.delete(item)

def putinputs():
    totalpair_sum=sum(item[9] for item in invoice_list)
    totalpair_entry.delete(0, ctk.END)
    totalpair_entry.insert(0, totalpair_sum)

    total_sum=sum(item[10] for item in invoice_list)
    total_entry.delete(0, ctk.END)
    total_entry.insert(0, total_sum)

    grandtotal_sum=(total_sum*0.15)+total_sum
    grandtotal_entry.delete(0, ctk.END)
    grandtotal_entry.insert(0, grandtotal_sum)

total_btn = ctk.CTkButton(frame1, text="Total", height=40, width=100,fg_color=GREEN, command=putinputs)
total_btn.grid(row=3, column=0, sticky="nw",padx=48, pady=2)
window.bind('<Shift-Return>', lambda event: total_btn.invoke())
delete_btn = ctk.CTkButton(frame1, text="Delete", command=delete_item,width=90, height=40, fg_color='#800000')
delete_btn.grid(row=3, column=0, sticky="ne", padx=45,pady=2)

####################################################################################### Right-bottom Total pair Total packages and Grand total

totalpair = ctk.CTkLabel(frame1, text="Total pair  :")
totalpair.grid(row=4, column=0, sticky='nw', padx=20, pady=10)
totalpair_entry = ctk.CTkEntry(frame1)
totalpair_entry.grid(row=4, column=0, pady=10, padx=40, sticky='e')

total = ctk.CTkLabel(frame1, text="Total  :")
total.grid(row=5, column=0, sticky='nw', padx=20, pady=10)
total_entry = ctk.CTkEntry(frame1)
total_entry.grid(row=5, column=0, pady=10, padx=40, sticky='e')

packing = ctk.CTkLabel(frame1, text="Packages :")
packing.grid(row=6, column=0, sticky='nw', padx=20, pady=10)
packing_entry = ctk.CTkEntry(frame1)
packing_entry.grid(row=6, column=0, pady=15, padx=40, sticky='e')

grandtotal = ctk.CTkLabel(frame1, text="Grand Total :")
grandtotal.grid(row=7, column=0, sticky='nw', padx=20)
grandtotal_entry = ctk.CTkEntry(frame1)
grandtotal_entry.grid(row=7, column=0, padx=40, sticky='e')

grandtotal = ctk.CTkLabel(frame1, text="15% Vat Inclusive", font=ctk.CTkFont(size=11))
grandtotal.grid(row=8, column=0, padx=42, sticky='ne')
new_invoice_button = ctk.CTkButton(frame1, text="New Invoice", command=new_invoice, height=35)
new_invoice_button.grid(row=9, column=0, columnspan=3, sticky="news", padx=48, pady=8)
window.bind('<Control-n>', lambda event: new_invoice_button.invoke())

############################################################################################### Tree View - The Table Board

style = ttk.Style()
tree_font = ctk.CTkFont(family="Helvetica", size=32)
style.configure("Treeview", font=tree_font)
style.configure("Treeview", rowheight=50)
header_font = ctk.CTkFont(family="Helvetica", size=35) 
style.configure("Treeview.Heading", font=header_font, background="green")

columns = ('Article', 'Colour', '5', '6','7','8', '9', '10', 'Rate', 'Total Pair', 'Total')
tree = ttk.Treeview(frame, columns=columns, show="headings", height=32)

tree.heading('Article', text='Article')
tree.heading('Colour', text='Colour')
tree.heading('5', text='5')
tree.heading('6', text="6")
tree.heading('7', text="7")
tree.heading('8', text="8")
tree.heading('9', text="9")
tree.heading('10', text="10")
tree.heading('Rate', text="Rate")
tree.heading('Total Pair', text="Total Pair")
tree.heading('Total', text="Total")

for col in columns:
    tree.column(col, anchor="center")

tree.column('Article', width=130) 
tree.column('Colour', width=120) 
tree.column('5', width=100) 
tree.column('6', width=100) 
tree.column('7', width=100)
tree.column('8', width=100)
tree.column('9', width=100)
tree.column('10', width=100)
tree.column('Rate', width=120)
tree.column('Total Pair', width=140)
tree.column('Total', width=100)

tree.grid(row=9, column=0, columnspan=3, padx=0, pady=15, sticky="nsew")

########################################################################################## Generate invoice button

frame2=ctk.CTkFrame(frame)
frame2.grid(row=5,column=0, columnspan=3,rowspan=4, sticky='news', padx=10, pady=8)
save_invoice_button = ctk.CTkButton(frame2, text="Generate Invoice", font=ctk.CTkFont(size=15), command=generate_invoice, width=680, height=45)
save_invoice_button.grid(row=0, column=2, columnspan=2, sticky="news", padx=10, pady=10)
window.bind('<Control-Return>', lambda event: save_invoice_button.invoke())

########################################################################################## Side Bar
sidebar_visible = False
def toggle_sidebar():
    global sidebar_visible

    def slide_in(pos):
        if pos < 0:
            sidebar.place(x=pos, y=0)  
            frame.after(10, lambda: slide_in(pos + 10))
        else:
            sidebar.place(x=0, y=0) 

    def slide_out(pos):
        if pos > -200:
            sidebar.place(x=pos, y=0)  
            frame.after(10, lambda: slide_out(pos - 10))
        else:
            sidebar.place(x=-200, y=0)  

    if sidebar_visible:
        slide_out(0)  
        sidebar_visible = False
    else:
        slide_in(-200)  
        sidebar_visible = True 

sidebar = ctk.CTkFrame(window, width=200, height=800, corner_radius=10)
sidebar.place(x=-200, y=0) 

label = ctk.CTkLabel(sidebar, text="PDF Converter", font=("Helvetica", 24))
label.pack(pady=20, padx=20)

file_entry = ctk.CTkEntry(sidebar, width=150)
file_entry.pack(pady=10, padx=10)

browse_button = ctk.CTkButton(sidebar, text="Browse", command=browse_file)
browse_button.pack(pady=5)

convert_button = ctk.CTkButton(sidebar, text="Convert to PDF", command=convert_to_pdf)
convert_button.pack(pady=5)

def open_link():
    url = "https://www.ilovepdf.com/word_to_pdf"
    webbrowser.open(url)
link_button = ctk.CTkButton(sidebar, text="PDF Converter\n(Website)", command=open_link)
link_button.pack(pady=20)

########################################################################################## Generate converters down parts- Converter Database Clear and Reset

frame3=ctk.CTkFrame(sidebar, width=120,height=700).pack(pady=10, padx=10)

resources = ctk.CTkButton(frame2, text="Converter", command=toggle_sidebar, width=100, height=30)
resources.grid(row=1,column=2, padx=10, sticky='nw')

database = ctk.CTkButton(frame2, text="Database", width=100, height=30, fg_color=GREEN)
database.grid(row=1,column=2, padx=120, sticky='nw')

clear = ctk.CTkButton(frame2, text="Clear", width=100, height=30, fg_color='#800000', command=clear)
clear.grid(row=1,column=2, padx=230, sticky='nw')

reset = ctk.CTkButton(frame2, text="Reset", width=100, height=30, fg_color='#686D76')
reset.grid(row=1,column=3, padx=12, sticky='ne')



window.mainloop()