import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import json
from tkcalendar import *
from tkcalendar import DateEntry
import openpyxl, xlrd
from openpyxl import Workbook
import pathlib
import xlwt

class App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("800x800")
        self.root.title('Store Details')

        #date widget
        sel = tk.StringVar()
        ttk.Label(root, text="Requirement Date:").grid(row=3, column=0, padx=10, pady=10)
        cal = DateEntry(root, selectmode = 'day', textvariable=sel)
        cal.grid(row=3,column=1,padx=15)

        file_path = "C:/Users/vyoma/Downloads/excelapp-main (1)/excelapp-main/data.json"
        with open(file_path,"r") as file:
            self.data = json.load(file)

        #Defining the variables
        self.Codes_var = tk.StringVar()
        self.names_var = tk.StringVar()
        self.munit_var = tk.StringVar()

        #creating material code dropdown combos
        ttk.Label(root, text="Material Code:").grid(row=0, column=0, padx=10, pady=10)
        self.Codes_combo = ttk.Combobox(root, textvariable=self.Codes_var)
        self.Codes_combo.grid(row=0, column=1, padx=10, pady=10)
        self.Codes_combo.bind("<<ComboboxSelected>>", self.update_names)
        self.Codes_combo.set("Select code")

        #creating material code dropdown combos
        ttk.Label(root, text="Material Description:").grid(row=1, column=0, padx=10, pady=10)
        self.names_combo = ttk.Combobox(root, textvariable=self.names_var)
        self.names_combo.grid(row=1, column=1, padx=10, pady=10)
        self.names_combo.bind("<<ComboboxSelected>>", self.update_names)

        #creating material unit of measurement dropdown combos
        ttk.Label(root, text="Unit of measurement:").grid(row=2, column=0, padx=10, pady=10)
        self.munit_combo = ttk.Combobox(root, textvariable=self.munit_var)
        self.munit_combo.grid(row=2, column=1, padx=10, pady=10)
        
        #updating codes in dropdown
        self.update_Codes()

        #to show msg asking to select code first
        #self.names_combo.set("select material code first")
        #self.munit_combo.set("select material code first")

        #creating submit button
        self.submit_button = tk.Button(root, text="Submit Request",
                                       relief = "groove",
                                       bg='LightGreen',
                                       activebackground='White',
                                       command=self.print_selected)
        self.submit_button.grid(row=4, column=1, pady=10)

        #create clear button
        #self.clear_button = tk.Button(root, text="Clear",
         #                             relief='groove',
          #                            command=self.clear_selection)
        #self.clear_button.grid(row=3, column=2, pady=10)

    #writing the searchable function
    #def search(event):
    #    value = event.widget.get()
    #    if value == '':
    #        combo_box['value'] = self.data

    #selecting the codes
    def print_selected(self):
        country = self.Codes_var.get()
    
    def update_Codes(self):
        Codes = list(self.data.keys())
        self.names_combo["values"]  = []
        self.names_combo["values"]   = []
        self.Codes_combo["values"] = Codes

        self.names_var.set('Select Material name')
        self.munit_var.set('Select Measurement unit')

    def update_names(self, *args):
        # Get the selected Code from the codes dropdown combo
        Codes = self.Codes_var.get()
        
        # create list containing the code values
        keys = list(self.data.keys())

        # Check if the selected code is a valid key in the data dictionary
        if Codes not in keys:
            # Clear the name and unit dropdowns and return early if the code is not valid
            self.names_var.set('Select a Code')
            self.munit_var.set('Select a Code')
            self.names_combo["values"]  = []
            self.munit_combo["values"]   = []
            return
        
        # Get the list of names for the selected code
        names = sorted(list(self.data[Codes].keys()))
        
        # Update the name dropdown combo with the list of name
        self.names_combo.config(values=names)
        self.munit_combo["values"]   = []
        
        self.update_munit()

        
    def update_munit(self, *args):
        # Get the selected name from the name dropdown combo
        names = self.names_var.get()

        #let's get a list of names, whom parent is self.code_var.get()
        Codes_name = self.Codes_var.get();
        names_in_codes = list(self.data.get(Codes_name))
        
        # Get the list of cities for the selected state and country
        munit = sorted(self.data[Codes_name][names])
        
        # Update the city dropdown combo with the list of cities
        self.munit_combo.config(values=munit)
        
        # Clear the selected city variable
        self.munit_var.set("Select Measurement Unit")


# Create the Tkinter window and run the app
root = tk.Tk()
app = App(root)
root.mainloop()

