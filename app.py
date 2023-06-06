import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from openpyxl import load_workbook
from datetime import datetime


def add_to_database() -> None:
    """Add the textbox contents to the database."""
    id_value = id_entry.get().strip()
    last_name_value = last_name_entry.get().strip()
    first_name_value = first_name_entry.get().strip()
    start_date_value = start_date_entry.get().strip()
    roll_off_date_value = roll_off_date_entry.get().strip()

    if not id_value or not last_name_value or not first_name_value or not start_date_value or not roll_off_date_value:
        messagebox.showerror("Error", "All fields must be filled.")
        return

    try:
        datetime.strptime(start_date_value, "%Y-%m-%d")
        datetime.strptime(roll_off_date_value, "%Y-%m-%d")
    except ValueError:
        messagebox.showerror("Error", "Invalid date format. Please use YYYY-mm-dd.")
        return

    # Open the existing workbook
    workbook = load_workbook("hackathon_user_db.xlsx")
    sheet = workbook.active

    # Find the first empty row
    row = 1
    while sheet.cell(row=row, column=1).value:
        row += 1

    # Write the textbox contents to the Excel file
    sheet.cell(row=row, column=1).value = id_value
    sheet.cell(row=row, column=2).value = last_name_value
    sheet.cell(row=row, column=3).value = first_name_value
    sheet.cell(row=row, column=4).value = start_date_value
    sheet.cell(row=row, column=5).value = roll_off_date_value

    # Save the workbook
    workbook.save("hackathon_user_db.xlsx")
    messagebox.showinfo("Success", "Data added to the database.")

    # Clear the textboxes
    id_entry.delete(0, tk.END)
    last_name_entry.delete(0, tk.END)
    first_name_entry.delete(0, tk.END)
    start_date_entry.delete(0, tk.END)
    roll_off_date_entry.delete(0, tk.END)


def close_app() -> None:
    """Close the application."""
    root.destroy()


# Create the main window
root = tk.Tk()
root.title("Data Entry App")
root.configure(bg="#dadcee")

# Create labels and entry fields
id_label = tk.Label(root, text="ID")
id_label.grid(row=0, column=0, padx=5, pady=5)
id_entry = ttk.Entry(root)
id_entry.grid(row=0, column=1, padx=5, pady=5)

last_name_label = tk.Label(root, text="Last Name")
last_name_label.grid(row=1, column=0, padx=5, pady=5)
last_name_entry = ttk.Entry(root)
last_name_entry.grid(row=1, column=1, padx=5, pady=5)

first_name_label = tk.Label(root, text="First Name")
first_name_label.grid(row=2, column=0, padx=5, pady=5)
first_name_entry = ttk.Entry(root)
first_name_entry.grid(row=2, column=1, padx=5, pady=5)

start_date_label = tk.Label(root, text="Start Date")
start_date_label.grid(row=3, column=0, padx=5, pady=5)
start_date_entry = ttk.Entry(root)
start_date_entry.grid(row=3, column=1, padx=5, pady=5)

roll_off_date_label = tk.Label(root, text="Roll-off Date")
roll_off_date_label.grid(row=4, column=0, padx=5, pady=5)
roll_off_date_entry = ttk.Entry(root)
roll_off_date_entry.grid(row=4, column=1, padx=5, pady=5)

# Create the "Add to Database" button
add_button = ttk.Button(root, text="Add to Database", command=add_to_database)
add_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

# Create the close button
close_button = ttk.Button(root, text="X", command=close_app)
close_button.grid(row=0, column=2, padx=5, pady=5)

# Run the GUI
root.mainloop()
