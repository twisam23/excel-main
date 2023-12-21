import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle

class StoreApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("400x200")
        self.root.title('Store Supplies')

        # Create a ThemedStyle to apply CSS styling
        self.style = ThemedStyle(self.root)
        self.style.set_theme("plastik")

        # Style for the heading
        self.style.configure("Heading.TLabel", font=("Helvetica", 20), foreground="blue")

        # Style for the BF1 button
        self.style.configure("BF1.TButton", font=("Helvetica", 14), foreground="green")

        # Style for the option buttons
        self.style.configure("Option.TButton", font=("Helvetica", 12))

        # Create and style the heading
        heading_label = ttk.Label(root, text="Store Supplies", style="Heading.TLabel")
        heading_label.pack(pady=20)

        # BF1 Button
        bf1_button = ttk.Button(root, text="BF1", style="BF1.TButton", command=self.show_bf1_options)
        bf1_button.pack(pady=10)

    def show_bf1_options(self):
        # Close the current dialogue box
        self.root.withdraw()

        # Create a new dialog for BF1 options
        bf1_dialog = tk.Toplevel(self.root)
        bf1_dialog.title("BF1 Options")
        bf1_dialog.geometry("500x250")

        # Adding Materials Button
        add_materials_button = ttk.Button(bf1_dialog, text="Add Materials", style="Option.TButton", command=self.add_materials)
        add_materials_button.pack(pady=10)

        # Taking Out Materials Button
        take_out_materials_button = ttk.Button(bf1_dialog, text="Take Out Materials", style="Option.TButton", command=self.take_out_materials)
        take_out_materials_button.pack(pady=10)

        # Show Present Stock Button
        show_stock_button = ttk.Button(bf1_dialog, text="Show Present Stock", style="Option.TButton", command=self.show_stock)
        show_stock_button.pack(pady=10)

    def add_materials(self):
        # Implement the code for adding materials here
        pass

    def take_out_materials(self):
        # Implement the code for taking out materials here
        pass

    def show_stock(self):
        # Implement the code for showing present stock here
        pass

# Create the Tkinter window and run the app
root = tk.Tk()
app = StoreApp(root)
root.mainloop()
