#import libraries make sure that all installed on your computer using pip install:
import tkinter as tk 
import customtkinter as ctk
from tkinter import * 
from tkinter import ttk
from tkinter import messagebox, filedialog
import csv
import os
import pandas as pd
import xlsxwriter
import sqlite3
import arabic_reshaper
from bidi.algorithm import get_display


#connect to sqlite and creating database file:
#make sure to install "DB Browser (SQLite)"

#create table if not exists:
connection = sqlite3.connect("Pharmacy Database.db")

#create table if not exists:
cursor = connection.execute(
    """
    CREATE TABLE IF NOT EXISTS  pharma (
    'id' INTEGER NOT NULL,
    'item_name' TEXT   NOT NULL,
    'price' REAL NOT NULL, 
    'sale' REAL NOT NULL,
    'company' TEXT   NOT NULL,
    PRIMARY KEY("id" AUTOINCREMENT)
   
    );
    """);
cursor = connection.cursor()

# Global Varibles:
root = ctk.CTk() #main window
update_window= '' #update window
add_window = '' #add window 

mydata = [] #empty list for import and export purpose

count = 0 #color the rows of treeview 

item = StringVar() #search item 
company = StringVar() #search company

t1 = StringVar() #company variable
t2 = StringVar() #sale variable
t3 = StringVar() #price variable
t4 = StringVar() #item name variable
t5 = StringVar() #id variable

#Functions:  
#make arabic letters in right order because in tkinter it exchange the order of arabic letters:
def getarabic(text):
    name_Button = text
    name_text= arabic_reshaper.reshape(name_Button)
    name_bidi= get_display(name_text)
    return name_bidi
#call and display the rows in treeview with update:
def call(rows):
    global mydata
    global count
    mydata = rows
    trv.delete(*trv.get_children())
    
    for i in rows:
        if count% 2 == 0:
            trv.insert('', 'end', values=i, tags=("even",))
        else:
            trv.insert('', 'end', values=i, tags=("odd",))
        count+=1
#Search botton:
def search_btn():
    qitem = item.get()
    qcompany = company.get()
    if qitem is None:
        query = "SELECT company,ROUND(sale, 2),ROUND(price, 2), item_name, id   FROM pharma WHERE company LIKE '%"+qcompany+"%'"
    elif qcompany is None:
        query = "SELECT company,ROUND(sale, 2),ROUND(price, 2), item_name, id   FROM pharma WHERE item_name LIKE '%"+qitem+"%'"
    else:
        query = "SELECT company,ROUND(sale, 2),ROUND(price, 2), item_name, id FROM pharma WHERE company LIKE '%"+qcompany+"%' AND item_name LIKE '%"+qitem+"%'"
    cursor.execute(query)
    rows = cursor.fetchall()
    call(rows)
    trv.focus_set()
    children = trv.get_children()
    if children:
        trv.focus(children[0])
        trv.selection_set(children[0])

#Search by pressing Enter in the keyboard:
def search_enter(event):
    search_btn()

#clear any actions:
def clear():
    query = "SELECT company,ROUND(sale, 2),ROUND(price, 2), item_name, id FROM pharma"
    cursor.execute(query)
    rows = cursor.fetchall()
    call(rows)

#get the info of specific row by clicking on it (or) pressing Enter:
def getrow(event):    
    rowid = trv.identify_row(event.y)
    item = trv.item(trv.focus())
    t1.set(item['values'][0])
    t2.set(item['values'][1])
    t3.set(item['values'][2])
    t4.set(item['values'][3])
    t5.set(item['values'][4])
    updatewindow()

#add new item by clicking on botton:
def add_new():
    item_name = t4.get()
    price = t3.get()
    sale = t2.get()
    company = t1.get()
    if (item_name is not None)and (price is not None) and (sale is not None) and (company is not None):
        query = "INSERT INTO  pharma (id, item_name, price, sale, company ) VALUES (NULL,?, ?, ?,?)"
        cursor.execute(query, ( item_name, price, sale, company))
        connection.commit()
        clear()
        close_add()
        messagebox.showinfo("Item Added", "This item: ' "+item_name+" ' has been added successfully.")
    else:
        messagebox.showerror("Error Window", "No data entered")
        return False
    
#add new item by pressing Enter:
def add_new_enter(event):
    add_new()

#update exist item by clicking on botton:
def update_item():
    item_name = t4.get()
    if messagebox.askyesno("Confirm Please", "Are you sure you want to update this Item?"):
        cursor.execute("""UPDATE pharma SET
                       item_name = :item,
                       price = :price,
                       sale = :sale,
                       company = :company
                       WHERE id = :iid""",
                       {
                       'item' : t4.get(),
                       'price' : t3.get(),
                       'sale' : t2.get(),
                       'company' : t1.get(),
                       'iid' : t5.get()   
                       })
        connection.commit()
        clear()
        close_update()
        messagebox.showinfo("Item Updated", "This item: ' " + item_name + "' has been updated successfully.")
    else:
        return True

#update exist item  by pressing Enter:
def update_item_enter(event):
    update_item()

#delete exist item:
def delete_item():
    iid = t5.get()
    item_name = t4.get()
    if messagebox.askyesno("Confirm Delete?", "Are you sure you want to delete this Item?"):
        query = "DELETE FROM pharma Where id = "+ iid
        cursor.execute(query)
        connection.commit()
        clear()
        close_update()
        messagebox.showinfo("Item Deleted", "This item: ' "+ item_name + " ' has been deleted!")
    else:
        return True

#delete the content of the table in the database:
def delete_table():
    if messagebox.askyesno("Confirm Delete?", "Are you sure you want to delete the table?"):
        query = "DELETE FROM pharma"
        query2 = "DELETE FROM SQLITE_SEQUENCE WHERE name='pharma'"
        cursor.execute(query)
        cursor.execute(query2)
        connection.commit()
        clear()
        messagebox.showinfo("Table Deleted", "The table has been deleted!")
    else:
        return True
    
#export the table into excel file (xlsx):
def exportfile():
    if len(mydata) < 1:
        messagebox.showerror("Error Window", "No data avaliable to export")
        return False
    path = filedialog.asksaveasfilename(initialdir=os.getcwd(), title="Save Window", filetypes=(("XLSX File", "*.xlsx"),("All File","*.*")))
    if (len(path) == 0):
        return False
    if (not path.endswith(".xlsx")):
        path = path + ".xlsx"
        
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "الكود")
    worksheet.write(0, 1, "اسم الصنف")
    worksheet.write(0, 2, "السعر")
    worksheet.write(0, 3, "الخصم")
    worksheet.write(0, 4, "اسم الشركة")
    for row_num, row_data in enumerate(mydata):
        for col_num, col_data in enumerate(reversed(row_data)):
            worksheet.right_to_left()
            worksheet.write(row_num+1, col_num, col_data)
            
    workbook.close()
    messagebox.showinfo("Data Exported", "Your data has been exported to ' "+os.path.basename(path)+" ' successfully.")

#import an excel file (xlsx) to the app:
def importexcel():
  mydata.clear()
  path = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select A File", filetypes=(("Excel files","*.xlsx"),("All File","*.*")))
  formatfile = r"{}".format(path)
  path, fileextension = os.path.splitext(formatfile)
  if (len(path) == 0):
      return False
  df = pd.read_excel(formatfile)
  dfrows = df.to_numpy().tolist()

  for i in  dfrows :
      mydata.append(i)
      
  # for i in mydata:

  for i in mydata:
      try:
          item_name = i[0]
          price = str(i[1])
          sale = str(i[2])
          company = i[3]
          
          if not isinstance(i[0], str):
              messagebox.showerror("Error Window", "تأكد من إدخال اسم الصنف")
              mydata.clear()
              return False
          elif (not isinstance(i[1], float)) and (not isinstance(i[1], int)):
              messagebox.showerror("Error Window", "تأكد من إدخال السعر")
              mydata.clear()
              return False
          elif (not isinstance(i[2], float)) and (not isinstance(i[2], int)):
              messagebox.showerror("Error Window", "تأكد من إدخال الخصم")
              mydata.clear() 
              return False
          elif not isinstance(i[3], str):
              messagebox.showerror("Error Window",  "تأكد من إدخال اسم الشركة")
              mydata.clear()
              return False
    
      except:
          messagebox.showerror("Error Window", "يجب ان يحتوي ملفك على: (اسم الصنف ، السعر ، الخصم ، اسم الشركة)")
          mydata.clear()
          return False
      if True:
          query = "INSERT INTO  pharma (id, item_name, price, sale, company) VALUES (NULL,?, ?, ?,?)"
          values = (item_name, price, sale, company)
          cursor.execute(query, values)
  call(mydata)        
  connection.commit()
  clear()
  messagebox.showinfo("Data Imported", "This file : ' "+os.path.basename(path)+" ' has been imported successfully.") 
#refresh the database if any changes happen to it from the database app:
def refresh():
    ent1 = Entry(section1, textvariable= item, width=30, justify='right')
    ent2 = Entry(section1, textvariable= company, width=30, justify='right')
    ent1.delete(0, 'end')
    ent2.delete(0, 'end')
    
    connection = sqlite3.connect("Pharmacy Database.db")
    cursor = connection.cursor()
    query = "SELECT company,ROUND(sale, 2),ROUND(price, 2),item_name, id from pharma"
    cursor.execute(query)
    rows = cursor.fetchall()
    call(rows)
#close update window:
def close_update():
    ent1 = Entry(update_window, textvariable=t1, width=32)
    ent2= Entry(update_window, textvariable=t2, width=32)
    ent3 = Entry(update_window, textvariable=t3, width=32)
    ent4 = Entry(update_window, textvariable=t4, width=32)
    ent5 = Entry(update_window, textvariable=t5, width=32)
    ent1.delete(0, 'end')
    ent2.delete(0, 'end')
    ent3.delete(0, 'end')
    ent4.delete(0, 'end')
    ent5.delete(0, 'end')
    update_window.destroy()

#close add window:
def close_add():
    ent1 = Entry(add_window, textvariable=t1, width=32)
    ent2= Entry(add_window, textvariable=t2, width=32)
    ent3 = Entry(add_window, textvariable=t3, width=32)
    ent4 = Entry(add_window, textvariable=t4, width=32)
    ent5 = Entry(add_window, textvariable=t5, width=32)
    ent1.delete(0, 'end')
    ent2.delete(0, 'end')
    ent3.delete(0, 'end')
    ent4.delete(0, 'end')
    ent5.delete(0, 'end')
    add_window.destroy()

#close the application:
def exitapp():
    global root
    root.destroy()
    root.quit()
    
#END OF THE FUNCTIONS

#windows of the app without the main window:

#update window with variable called "update_window" 
def updatewindow(): 
    #display:
    global update_window
    update_window = Toplevel(root, bg="#BCC6CC")
    update_window.geometry("800x250")
    update_window.title("تفاصيل الصنف")
    update_window.resizable(False, False)
    update_window.grab_set()
    
    #Label widget:
    lbl1 = Label(update_window, text ="اسم الصنف",font=(12),  bg="#BCC6CC")
    lbl2 = Label(update_window, text ="السعر",font=(12),  bg="#BCC6CC")
    lbl3 = Label(update_window, text ="الخصم",font=(12),  bg="#BCC6CC")
    lbl4 = Label(update_window, text ="اسم الشركة",font=(12),  bg="#BCC6CC")
    
    #Entry widget:
    ent1= Entry(update_window, textvariable=t4, width=40, justify='right')
    ent2 = Entry(update_window, textvariable=t3, width=40, justify='right')
    ent3 = Entry(update_window, textvariable=t2, width=40, justify='right')
    ent4 = Entry(update_window, textvariable=t1, width=40, justify='right')
    
    #pressing Enter:
    ent1.bind('<Return>', update_item_enter)
    ent2.bind('<Return>', update_item_enter)
    ent3.bind('<Return>', update_item_enter)
    ent4.bind('<Return>', update_item_enter)
    
    #Buttons:
    clos_btn = ctk.CTkButton(master=update_window ,text= getarabic("إغلاق"), command=close_update)  
    delete_btn = ctk.CTkButton(master=update_window ,text= getarabic("حذف الصنف"), command=delete_item)
    up_btn = ctk.CTkButton(master=update_window ,text= getarabic("تحديث"), command=update_item)
    
    #Responsive window to all sizes:
    for i in range(10):
           update_window.grid_rowconfigure(i, weight= 1)
           for j in range(10):
                update_window.grid_columnconfigure(j, weight=1)
                lbl1.grid(row=1, column=5, padx=10, pady=10)
                lbl2.grid(row=2, column=5, padx=10, pady=10)
                lbl3.grid(row=3, column=5, padx=10,pady=10)
                lbl4.grid(row=4, column=5, padx=10, pady=10)
                
                ent1.grid(row=1, column=4, padx=10, pady=10)
                ent2.grid(row=2, column=4, padx=10, pady=10)
                ent3.grid(row=3, column=4, padx=10, pady=10)
                ent4.grid(row=4, column=4, padx=10, pady=10)
                
                clos_btn.grid(row=6, column=3, padx=10, pady=10)
                delete_btn.grid(row=6, column=4, padx=10, pady=10)
                up_btn.grid(row=6, column=5, padx=10, pady=10)
    update_window.iconbitmap(r'img.ico')
 
#add window with variable called "add_window":
def addwindow():
    #display:
    global add_window
    add_window = Toplevel(root, bg="#BCC6CC")
    add_window.geometry("800x250")
    add_window.title("إضافة صنف جديد")
    add_window.resizable(False, False)
    add_window.grab_set() 
    add_window.iconbitmap(r'img.ico')
    
    #Label widget:
    lbl1 = Label(add_window, text ="اسم الصنف",font=(12),  bg="#BCC6CC")
    lbl2 = Label(add_window, text ="السعر",font=(12),  bg="#BCC6CC")
    lbl3 = Label(add_window, text ="الخصم",font=(12),  bg="#BCC6CC")
    lbl4 = Label(add_window, text ="اسم الشركة",font=(12),  bg="#BCC6CC")
    
    #Entry widget:
    ent1= Entry(add_window, textvariable=t4, width=40, justify='right')
    ent2 = Entry(add_window, textvariable=t3, width=40, justify='right')
    ent3 = Entry(add_window, textvariable=t2, width=40, justify='right')
    ent4 = Entry(add_window, textvariable=t1, width=40, justify='right')
    
    #pressing Enter:
    ent1.bind('<Return>', add_new_enter)
    ent2.bind('<Return>', add_new_enter)
    ent3.bind('<Return>', add_new_enter)
    ent4.bind('<Return>', add_new_enter)
    
    #Buttons:
    add_btn = ctk.CTkButton(master=add_window ,text= getarabic("إضافة صنف جديد"), command=add_new)
    clos_btn = ctk.CTkButton(master=add_window ,text= getarabic("إغلاق"), command=close_add)
    
    #Responsive window to all sizes:
    for i in range(10):
           add_window.grid_rowconfigure(i, weight= 1)
           for j in range(10):
               add_window.grid_columnconfigure(j, weight=1)
               lbl1.grid(row=1, column=5, padx=10, pady=10)
               lbl2.grid(row=2, column=5, padx=10, pady=10)
               lbl3.grid(row=3, column=5, padx=10,pady=10)
               lbl4.grid(row=4, column=5, padx=10, pady=10)
               
               ent1.grid(row=1, column=4, padx=10, pady=10)
               ent2.grid(row=2, column=4, padx=10, pady=10)
               ent3.grid(row=3, column=4, padx=10, pady=10)
               ent4.grid(row=4, column=4, padx=10, pady=10)
               clos_btn.grid(row=6, column=4, padx=10, pady=10)
               add_btn.grid(row=6, column=5, padx=10, pady=10)

    #delete after closing:
    ent1.delete(0, 'end')
    ent2.delete(0, 'end')
    ent3.delete(0, 'end')
    ent4.delete(0, 'end')


    
#sectioning the main window to three section:
#section 1 for search:
section1  = ctk.CTkFrame(master=root)
section1.pack(pady=10, padx=20, fill="both")

#section 2 for treeview:
section2  = ctk.CTkFrame(master=root)
section2.pack(pady=10, padx=10, fill="both", expand=True)

#section 3 for other bottons 
section3  = ctk.CTkFrame(master=root)
section3.pack(pady=10, padx=20, fill="both")

#start with section 2 (Treeview):
#scroll bar:
treescrolly = ttk.Scrollbar(section2)
treescrolly.pack(side=RIGHT, fill= Y)


#assign vaiable called columns to determine number of colums of treeview:
columns= (1,2,3,4,5)
# Create Treeview (trv) :
trv=ttk.Treeview(section2, columns=columns, show="headings", height="15", yscrollcommand= treescrolly.set,selectmode ="browse" )
trv.pack(expand =True, fill= BOTH)
treescrolly.config(command= trv.yview)


#heading of the treeview:
trv.heading(5, text="الكود", anchor=CENTER)
trv.heading(4, text="اسم الصنف", anchor=CENTER)
trv.heading(3, text="السعر", anchor=CENTER)
trv.heading(2, text="الخصم", anchor=CENTER)
trv.heading(1, text="اسم الشركة", anchor=CENTER)
#columns of treeview:
trv.column(1, minwidth=0, width=150, anchor=CENTER)
trv.column(2,  minwidth=0, width=40,anchor=CENTER)
trv.column(3,  minwidth=0, width=40,anchor=CENTER)
trv.column(4,  minwidth=0, width=350,anchor=CENTER)
trv.column(5, minwidth=0, width=40, anchor=CENTER)

#pressing Enter and double click:
trv.bind('<Double 1>', getrow)
trv.bind('<Return>', getrow)

#coloring of the rows:
trv.tag_configure('odd', background= '#FFFFFF')
trv.tag_configure('even', background='#F7F7F7')

#style of treeview:
style =ttk.Style()
style.theme_use("clam")

# Configure the style of Heading in Treeview widget
style.configure("Treeview", foreground="black", rowheight=25, fieldbackground="white")
style.configure("Treeview.Heading", foreground="black", font=("bold"), background="#357EC7")
style.map('Treeview', background=[('selected','#357EC7')])

# Display rows from database into treeview
query = "SELECT company,ROUND(sale, 2),ROUND(price, 2), item_name, id FROM pharma"
cursor.execute(query)
rows = cursor.fetchall()
call(rows)

# Section 1 (search part):
#Label widget:
lbl1 = Label(section1, text ="بحث الصنف", font=(12),  bg="#BCC6CC")
lbl1.pack(side=tk.RIGHT, pady=12, padx=5)
#Entry widget:
ent1 = Entry(section1, textvariable= item, width=30, justify='right')
ent1.pack(side=tk.RIGHT,pady=12, padx=10)
ent1.bind('<Return>', search_enter)
#Label widget:
lbl2 = Label(section1, text ="بحث الشركة", font=(12),  bg="#BCC6CC")
lbl2.pack(side=tk.RIGHT, pady=12, padx=5)
#Entry widget:
ent2 = Entry(section1, textvariable= company, width=30, justify='right')
ent2.pack(side=tk.RIGHT,pady=12, padx=10)
ent2.bind('<Return>', search_enter)

#Button widget:
btn = ctk.CTkButton(master=section1 ,text= getarabic("بحث"), command=search_btn)
btn.pack(side=tk.RIGHT,pady=12, padx=10)

refreshbtn =  ctk.CTkButton(master=section1 ,text= getarabic("إعادة تحميل"), command=refresh)
refreshbtn.pack(side=tk.RIGHT,pady=12, padx=10)

# Section 3 (bottons part):
expbtn = ctk.CTkButton(master=section3 ,text= getarabic("إخراج ملف") ,command=exportfile)
expbtn.pack(side=tk.RIGHT,pady=12, padx=10)


impbtn = ctk.CTkButton(master=section3 ,text= getarabic("إدخال ملف") , command=importexcel)
impbtn.pack(side=tk.RIGHT,pady=12, padx=10)


add_btn = ctk.CTkButton(master=section3 ,text= getarabic("إضافة صنف جديد"), command=addwindow)
add_btn.pack(side=tk.RIGHT,pady=12, padx=10)


deltable_btn = ctk.CTkButton(master=section3 ,text= getarabic("حذف الجدول"), command=delete_table)
deltable_btn.pack(side=tk.RIGHT,pady=12, padx=10)



extbtn = ctk.CTkButton(master=section3 ,text= getarabic("خروج"),command=exitapp)
extbtn.pack(side=tk.RIGHT,pady=12, padx=10)


# main loop
ctk.set_appearance_mode("light")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"
root.title("My Pharmacy")

# root.iconbitmap(r'img.ico')
root.geometry("1000x500")
root.mainloop()
del root
