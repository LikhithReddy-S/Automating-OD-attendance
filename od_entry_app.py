import tkinter as tk
from tkinter import messagebox
import pandas as pd

# Load student list
try:
    student_df = pd.read_excel("student_list.xlsx")
    student_df.columns = student_df.columns.str.strip()  # clean headers
    student_df['Name'] = student_df['Name'].astype(str).str.strip()
    student_df['Roll Number'] = student_df['Roll Number'].astype(str).str.strip()
    roll_to_name = dict(zip(student_df['Roll Number'], student_df['Name']))
except Exception as e:
    print("Error loading student_list.xlsx:", e)
    exit()

# List to store final output
entries = []

# Get name from roll number
def fetch_name():
    roll = roll_var.get().strip()
    name = roll_to_name.get(roll)
    if name:
        name_var.set(name)
    else:
        name_var.set("")
        messagebox.showerror("Not Found", "Roll number not found.")

# Add entry with slots
def add_entry():
    name = name_var.get().strip()
    roll = roll_var.get().strip()
    slots = slot_var.get().strip()

    if not name or not roll or not slots:
        messagebox.showerror("Missing Data", "Please fill all fields.")
        return

    try:
        slot_list = [int(s.strip()) for s in slots.split(",") if s.strip().isdigit()]
        if not slot_list:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid Slots", "Enter numeric slots separated by commas.")
        return

    for slot in slot_list:
        entries.append([name, roll, slot])

    name_var.set("")
    roll_var.set("")
    slot_var.set("")
    messagebox.showinfo("Success", f"{len(slot_list)} slot(s) added for {name}.")

# Export to Excel
def export_excel():
    if not entries:
        messagebox.showwarning("No Data", "No entries to export.")
        return
    df = pd.DataFrame(entries, columns=["Name", "Roll Number", "Slot"])
    df.to_excel("od_output.xlsx", index=False)
    messagebox.showinfo("Exported", "OD list saved to od_output.xlsx")

# GUI setup
root = tk.Tk()
root.title("OD Entry Tool")
root.geometry("400x280")

tk.Label(root, text="Roll Number:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
tk.Label(root, text="Name:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
tk.Label(root, text="Slots (comma-separated):").grid(row=2, column=0, padx=10, pady=5, sticky="e")

roll_var = tk.StringVar()
name_var = tk.StringVar()
slot_var = tk.StringVar()

tk.Entry(root, textvariable=roll_var, width=30).grid(row=0, column=1, padx=10, pady=5)
tk.Entry(root, textvariable=name_var, width=30, state="readonly").grid(row=1, column=1, padx=10, pady=5)
tk.Entry(root, textvariable=slot_var, width=30).grid(row=2, column=1, padx=10, pady=5)

tk.Button(root, text="Get Name", width=15, command=fetch_name).grid(row=0, column=2, padx=5)
tk.Button(root, text="Add Entry", width=15, command=add_entry).grid(row=3, column=0, columnspan=2, pady=15)
tk.Button(root, text="Export to Excel", width=15, command=export_excel).grid(row=4, column=0, columnspan=2)

root.mainloop()
