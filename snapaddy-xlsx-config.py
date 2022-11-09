import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter.messagebox import showinfo

root = tk.Tk()
root.title('snapADDY Excel configurator')
root.geometry('400x500')

label1 = tk.Label(root, text="Hallo Welt")
label1.pack()

## Get File Paths ##
booth_staff = ''
export_file = ''
save_as_file = ''

# Get Excel Path for Booth Staff
def select_staff():
    global booth_staff
    filetypes = (
        ('Excel files', '*.xlsx'),
        ('All files', '*.*')
    )
    filename = filedialog.askopenfilename(
        title='Standpersonal auswählen', 
        initialdir='',
        filetypes=filetypes
    )
    booth_staff = filename

    if booth_staff:
        open_staff["state"] = "disabled"
        open_staff.config(text=booth_staff)
    if export_file and booth_staff and save_as_file:
        calc_result["state"] = "active"

# Get snapADDY Export Excel File
def select_exportfile():
    global export_file
    filetypes = (
        ('Excel files', '*.xlsx'),
        ('All files', '*.*')
    )
    filename = filedialog.askopenfilename(
        title='snapADDY Exportdatei auswählen', 
        initialdir='',
        filetypes=filetypes
    )
    export_file = filename

    if export_file:
        open_export["state"] = "disabled"
        open_export.config(text=export_file)
    if export_file and booth_staff and save_as_file:
        calc_result["state"] = "active"


# Get folder path where to save the final file
def save_path():
    global save_as_file

    filename = filedialog.askdirectory(
        title='snapADDY Exportdatei auswählen', 
        initialdir=''
    )
    save_as_file = filename
    print(f'save as file: {save_as_file}')

    if save_as_file:
        save_button["state"] = "disabled"
        save_button.config(text=save_as_file)
    if export_file and booth_staff and save_as_file:
        calc_result["state"] = "active"


# Calculate Result and close the App
def calculate_result():
    df_staff = pd.read_excel(booth_staff)
    df_exportfile = pd.read_excel(export_file)
    print(df_staff.head())
    print(df_exportfile.head())

    # Close Window and exit Process
    root.destroy()

## Text Boxes ##
staff_text = tk.Text(root, height=6, width=100)
staff_text_string = '''
Bitte Excel Tabelle mit 4 Spalten einfügen
Überschriften müssen in dieser Reihenfolge,
mit genau diesem Text angelegt sein:\n
Vorname | Nachname | mail |	Sales Organisation
'''
staff_text.insert(tk.END, staff_text_string)

## Buttons ##

# Button to open Booth Staff File
open_staff = tk.Button(
    root, 
    text='Excel Datei mit Standpersonal auswählen',
    command = select_staff
)

# Button to open snapADDY Export File  File
open_export = tk.Button(
    root, 
    text='snapADDY Export Datei auswählen',
    command = select_exportfile
)

# Button to ask for an export folder
save_button = tk.Button(
    root, 
    text='Speicherort für neue Datei auswählen',
    command = save_path
)

# Button to calculate results
calc_result = tk.Button(
    root, 
    text='Ergebnis Kalkulieren',
    command = calculate_result
)



## Design of all Elements ##

staff_text.pack()
open_staff.pack(expand=True)
open_export.pack(expand=True)
save_button.pack(expand=True)
calc_result.pack(expand=True)
calc_result["state"] = "disabled"



## Main Loop ##

root.mainloop()

## Create the final Excel File ##

# Load Excel Files into Pandas
staff_df = pd.read_excel(booth_staff)
export_df = pd.read_excel(export_file)

# Create Unique ID Column
export_df['uniqueId'] = export_df['Campaign ID']+"-"+export_df['Report ID'].astype(str) + "-" + export_df['Contact ID'].astype(str) 

# Redesign Staff Data Frame
staff_df_new = staff_df.loc[:,['mail', 'ID']].rename(columns={'mail':'Created by (email)'})

# Join the two tables together
final_export = pd.merge(left=export_df, right=staff_df_new, on='Created by (email)', how='left')

final_export.to_excel(f'{save_as_file}/final_export.xlsx')
