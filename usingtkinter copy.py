import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import openpyxl

class App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("800x800")
        self.root.title('Store Details')

        # StringVar to store the selected date
        self.date_var = tk.StringVar()

        # Date widget
        ttk.Label(root, text="Requirement Date:").grid(row=3, column=0, padx=10, pady=10)
        date_entry = Entry(root, textvariable=self.date_var)
        date_entry.grid(row=3, column=1, padx=15)
        date_entry.bind("<Return>", self.show_material_details)

        # Load data from Excel file
        file_path = "C:/Users/vyoma/Downloads/excelapp-main (1)/excelapp-main/excel app.xlsx"
        self.data = self.load_excel_data(file_path)

        # Defining the variables
        self.material_var = tk.StringVar()
        self.description_var = tk.StringVar()

        # Label for Material Name
        ttk.Label(root, text="Material Name:").grid(row=0, column=0, padx=10, pady=10)

        # Entry widget for typing Material Name
        self.material_entry = Entry(root, textvariable=self.material_var)
        self.material_entry.grid(row=0, column=1, padx=10, pady=10)
        self.material_entry.bind("<Return>", self.show_material_details)

    def load_excel_data(self, file_path):
        data = {}
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            for row in sheet.iter_rows(values_only=True):
                if len(row) == 3:
                    material, code, qty = row
                    data[material] = {'code': code, 'qty': qty}

            workbook.close()
        except Exception as e:
            print(f"Error loading data from Excel: {e}")

        return data

    def show_material_details(self, event):
        material_name = self.material_var.get()
        if material_name in self.data:
            material_data = self.data[material_name]
            material_code = material_data['code']
            material_qty = material_data['qty']
            material_info = f"Material Name: {material_name}\nMaterial Code: {material_code}\nQuantity: {material_qty}"
            messagebox.showinfo("Material Details", material_info)
        else:
            messagebox.showinfo("Material Details", "Material not found")

# Create the Tkinter window and run the app
root = tk.Tk()
app = App(root)
root.mainloop()
