import tkinter as tk
from tkinter import filedialog, Text
import os
from main import main
import time

def reset():
    error_label['text'] = ''
    dataexport_entry.delete(0,tk.END)
    project_number_entry.delete(0,tk.END)  



def gather_info():
    data_export_path = dataexport_entry.get()
    project_number = project_number_entry.get()
    print(data_export_path,project_number)

    #execute main function and capture returned values for front-end labels
    returned_values = main(data_export_path,project_number)
    print(returned_values)

    #run main and check if the info is valid for GUI

    if returned_values[0]==True:
        error_label['text'] = "Files were moved. OE file is cleaned in TAB. Be sure to remove apostropohes."
        error_label['fg'] = 'green'
    else:
        error_label['text'] = returned_values[1]
        error_label['fg'] = 'red'

    #clear fields
    dataexport_entry.delete(0,tk.END)
    project_number_entry.delete(0,tk.END)  


root = tk.Tk()
root.resizable(False,False)

#create canvas
canvas = tk.Canvas(root,height=400,width=600,bg="white",relief="ridge")
canvas.pack()

#image
image = tk.PhotoImage(file="assets/mvrg.gif")
image_label = tk.Label(image=image)
image_label.config(fg='black',bg='white')
canvas.create_window(300, 23, window=image_label)
# image_label.pack()

#create inner frame
frame  = tk.Frame(root)
frame.place(relwidth=0.8,relheight=0.8,relx=0.1,rely=0.1)

#title label
title = tk.Label(root,text='Post-Export Automation')
title.config(font=('helvetica', 20))
canvas.create_window(300, 80, window=title)

datamap_note = tk.Label(root,text='*****Please make sure a datamap already exists in the export folder*****')
datamap_note.config(fg='black',font=('helvetica',8))
canvas.create_window(300,115,window=datamap_note)

#data export label
dataexport_label = tk.Label(root, text='What is the data export path?')
dataexport_label.config(font=('helvetica', 10))
canvas.create_window(300, 150, window=dataexport_label)

#data export input box
dataexport_entry = tk.Entry(root,width=60)
canvas.create_window(300,180,window=dataexport_entry)
#dataexport_entry.pack()

#project number label
project_number_label = tk.Label(root,text="What is the project number?")
project_number_label.config(font=('helvetica', 10))
canvas.create_window(300,230,window=project_number_label)
#project_number.pack()

#project number input box
project_number_entry = tk.Entry(root)
canvas.create_window(300,260,window=project_number_entry)

#error box
error_label = tk.Label(root,text='')
error_label.config(fg="green")
canvas.create_window(300, 290, window=error_label)

#execute button
execute_button = tk.Button(root,text="Execute",padx=10,pady=5,fg="white",bg="#ddd",command=gather_info)
execute_button.config(fg='black',font=('helvetica', 10))
canvas.create_window(225,330,window=execute_button)
#execute_button.pack()

#reset button
reset_button = tk.Button(root,text="Reset",padx=10,pady=5,fg="white",bg="#ddd",command=reset)
reset_button.config(fg='black',font=('helvetica', 10))
canvas.create_window(375,330,window=reset_button)
root.mainloop()