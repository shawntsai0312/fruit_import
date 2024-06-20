import tkinter as tk
from tkinter import messagebox
from tkinter import font

from categorizeWithInput import categorize
from toCSV import toCSV
from toXLSX import toXLSX

def export_message():
    email_input = email_content_box.get("1.0", tk.END).strip()
    file_name_input = file_name_box.get("1.0", tk.END).strip()
    if email_input == "":
        messagebox.showwarning("No Input", "Please enter the email content to export the excel file.")
        return
    if file_name_input == "":
        messagebox.showwarning("No Input", "Please enter the output file name.")
        return
    
    cherrys = categorize(email_input)
    toXLSX(cherrys, output_file=file_name_input)
    messagebox.showinfo("Exported Message", f"Export Completed! The Excel file named {file_name_input}.xlsx is saved in the 'categorizedXLSX' folder.")

# Create the main window
root = tk.Tk()
root.title("Email Content to Excel")

# Create a custom font
message_font = font.Font(family="Helvetica", size=24)
input_font = font.Font(family="Helvetica", size=16)

# Create a label for the input box
file_name_label = tk.Label(root, text="Enter output file name:", font=message_font)
file_name_label.pack(pady=5)

# Create the input box (Text widget) with larger font
file_name_box = tk.Text(root, width=50, height=2, font=input_font)
file_name_box.pack(pady=5)

# Create a label for the input box
email_content_label = tk.Label(root, text="Enter email content:", font=message_font)
email_content_label.pack(pady=5)

# Create the input box (Text widget) with larger font
email_content_box = tk.Text(root, width=50, height=10, font=input_font)
email_content_box.pack(pady=5)

# Create the export button
export_button = tk.Button(root, text="Export", command=export_message, font=message_font)
export_button.pack(pady=20)

# Run the application
root.mainloop()
