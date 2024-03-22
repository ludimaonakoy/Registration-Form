import tkinter as tk
from tkinter import ttk
import openpyxl
# ttk(meaning) : thine Tkinter

#create a new function for loading date from excel by using the openpyxl
def load_data():
    workbook = openpyxl.load_workbook("people.xlsx")
    sheet = workbook.active

    list_values = list(sheet.values)
    print(list_values)
    #to add the column names to our treeview
    for col_name in list_values[0]:
        #i want to put my column name to the heading of the treeview
        treeview.heading(col_name, text=col_name)
    #to load the rest of values, index 0 is the column name
    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)

def insert_data():

    name = name_entry.get()
    age = int(age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"

   # print(name,age,subscription_status,employment_status)
    # insert data into Excel sheet
    workbook = openpyxl.load_workbook("people.xlsx")
    sheet = workbook.active
    row_values = [name,age,subscription_status,employment_status]
    sheet.append(row_values)
    workbook.save("people.xlsx")

    #inser data into the treeview
    treeview.insert('',tk.END, values=row_values)

    #clear the value
    name_entry.delete(0,"end")
    name_entry.insert(0,"Name")
    age_spinbox.delete(0,"end")
    age_spinbox.insert(0,"Age")
    status_combobox.set(combo_list[0])
    checkbtn.state(["!selected"])
#create the function for changing the mode of the light

def toggle_mode():
    if mode_switch.instate(["selected"]):
        Style.theme_use("forest-light")
    else:
        Style.theme_use("forest-dark")

root = tk.Tk()
root.title("Registration Form with Excel")
root.geometry("750x350+300+150")
#create a variable call style
# it will enable us to apply the themed to tkinter application
Style = ttk.Style(root)
# call the file that we have here
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
# to choose what theme we want to use
Style.theme_use("forest-dark")

#set the list of combobox list
combo_list = ["Subscribed", "Not Subscribed", "Other"]

#start creating our widget
#the first widget frame
frame = ttk.Frame(root)
frame.pack(anchor="center")

#creation of sub frame
widget_frame = ttk.LabelFrame(frame, text="Insert Data")
widget_frame.grid(row=0,column=0, padx=20,pady=10)

#creation of name widget
name_entry = ttk.Entry(widget_frame)
#first element setting up the placeholder
name_entry.insert(0,"Name")
#delation of place holder when the user type
name_entry.bind("<FocusIn>", lambda e: name_entry.delete("0","end"))
name_entry.grid(row=0,column=0, sticky="ew",padx=5,pady=(0,5))

#second element setting up the spinbox for age wich start from 18 to 100
age_spinbox = ttk.Spinbox(widget_frame, from_= 18, to=100 )
age_spinbox.grid(row=2,column=0, sticky="ew",padx=5,pady=(0,5))
age_spinbox.insert(0,"Age")

#thidth element combobox
status_combobox = ttk.Combobox(widget_frame, values=combo_list)
status_combobox.grid(row=3,column=0, sticky="ew",padx=5,pady=(0,5))
status_combobox.current(0)#set the first element of a combo_list

#fifth create checkbox
# the variable a will be used to save a boolean(0,1) value whether he is an employed or Unenployed
a = tk.BooleanVar()
checkbtn = ttk.Checkbutton(widget_frame, text="Employed", variable=a)
checkbtn.grid(row=4,column=0,sticky="nsew",padx=5,pady=(0,5))

#create a button
btn = ttk.Button(widget_frame, text="Insert", command=insert_data)
btn.grid(row=5,column=0,sticky="nsew",padx=5,pady=(0,5))

# create a separator : to separate the insert part and the switch mode(light to dark)
separator = ttk.Separator(widget_frame)
separator.grid(row=6,column=0,padx=(20,10),pady=10,sticky="ew")

#create the checkbutton for switch mode(light to dark)
mode_switch = ttk.Checkbutton(
    widget_frame,text="Mode",style="Switch",command=toggle_mode)

mode_switch.grid(row=7,column=0,padx=5,pady=10,sticky="nsew")

#create a new frame for excel view
treeframe = ttk.Frame(frame)
treeframe.grid(row=0,column=1,pady=10)
#creation of scroll-bar connect it to the treeview as main root(treeviwe)
treescroll = ttk.Scrollbar(treeframe)
treescroll.pack(side="right", fill="y")
#before we create the treeview we create first the columns by specifying just like in the Excel file
cols = ("Name","Age","Subscription","Employment")
#the creation of the treeview
treeview = ttk.Treeview(treeframe,show="headings",yscrollcommand=treescroll.set,columns=cols,height=13)
# to resize the width of the treeview
treeview.column("Name",width=100)
treeview.column("Age",width=50)
treeview.column("Subscription",width=100)
treeview.column("Employment",width=100)
treeview.pack()
treescroll.config(command=treeview.yview)
load_data()
#space in between widgets very important
for widgets in widget_frame.winfo_children():
    widgets.grid_configure(padx=5,pady=9)

root.mainloop()