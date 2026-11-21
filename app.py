import os
import re
from datetime import datetime
from collections import Counter
import openpyxl
from tkinter import *
from tkinter import messagebox
import matplotlib.pyplot as plt

# --- File Configuration ---
script_dir = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(script_dir, "health_insurance_data.xlsx")

HEADERS = [
    "Registration ID", "Name", "Email ID", "Phone Number", 
    "Address", "Date of Birth", "Gender", "Age", 
    "Insurance Type", "Agree to Terms"
]

def load_or_create_workbook():
    """Load existing Excel database or initialize a new workbook."""
    global wb, sheet
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Insurance Records"
        sheet.append(HEADERS)
        excel_setup()
        wb.save(EXCEL_FILE)

def excel_setup():
    """Set standard column widths for clean readability."""
    widths = {
        'A': 20, 'B': 25, 'C': 30, 'D': 16, 'E': 35,
        'F': 16, 'G': 12, 'H': 12, 'I': 18, 'J': 15
    }
    for col, width in widths.items():
        if col in sheet.column_dimensions:
            sheet.column_dimensions[col].width = width

def get_insurance_type_col_idx():
    """Dynamically find the column index for Insurance Type."""
    first_row = [str(c or '').strip().lower() for c in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    for idx, col_name in enumerate(first_row):
        if "insurance type" in col_name or col_name == "type":
            return idx
    # Fallback to column I (index 8)
    return 8

def clear():
    """Clear all user input fields in the registration form."""
    reg_id_field.delete(0, END)
    name_field.delete(0, END)
    email_field.delete(0, END)
    phone_field.delete(0, END)
    address_field.delete(0, END)
    dob_field.delete(0, END)
    gender.set(None)
    age.set(None)
    insurance_type.set(None)
    agree_terms.set(0)

def insert():
    """Validate user inputs and insert applicant record into Excel database."""
    # Check mandatory fields
    if (not reg_id_field.get().strip() or not name_field.get().strip() or 
        not email_field.get().strip() or not phone_field.get().strip() or 
        not address_field.get().strip() or not dob_field.get().strip() or 
        not gender.get() or not age.get() or not insurance_type.get() or not agree_terms.get()):
        messagebox.showerror("Input Error", "All fields are required and terms must be accepted.")
        return

    # Validate Phone Number (10 digits)
    phone_val = phone_field.get().strip()
    if not re.match(r'^\d{10}$', phone_val):
        messagebox.showerror("Input Error", "Phone Number must be exactly 10 digits.")
        return

    # Validate Email Format
    email_val = email_field.get().strip()
    if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email_val):
        messagebox.showerror("Input Error", "Please enter a valid email address (e.g. user@gmail.com).")
        return

    # Validate Date of Birth and Age Requirement (18+)
    dob_str = dob_field.get().strip()
    dob = None
    for fmt in ('%Y-%m-%d', '%Y-%d-%m', '%d-%m-%Y'):
        try:
            dob = datetime.strptime(dob_str, fmt)
            break
        except ValueError:
            pass

    if not dob:
        messagebox.showerror("Input Error", "Date of Birth must be in YYYY-MM-DD format.")
        return

    now = datetime.now()
    age_today = now.year - dob.year - ((now.month, now.day) < (dob.month, dob.day))
    if age_today < 18:
        messagebox.showerror("Input Error", "Applicant must be at least 18 years old to register.")
        return

    # Insert Record
    record = [
        reg_id_field.get().strip(),
        name_field.get().strip(),
        email_val,
        phone_val,
        address_field.get().strip(),
        dob.strftime('%Y-%m-%d'),
        gender.get(),
        age.get(),
        insurance_type.get(),
        1 if agree_terms.get() else 0
    ]

    sheet.append(record)
    try:
        wb.save(EXCEL_FILE)
        clear()
        messagebox.showinfo("Success", "Insurance registration record saved successfully!")
    except PermissionError:
        messagebox.showerror("Save Error", "Please close the Excel file if it is currently open and try again.")

def show_chart():
    """Generate and display interactive charts based on registered insurance types."""
    col_idx = get_insurance_type_col_idx()
    insurance_types = []
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row and len(row) > col_idx and row[col_idx]:
            val = str(row[col_idx]).strip()
            if val.lower() not in ["none", "", "insurance type", "type"]:
                insurance_types.append(val)

    if not insurance_types:
        messagebox.showinfo("No Data", "No insurance registration records available to display.")
        return

    type_count = Counter(insurance_types)
    labels = list(type_count.keys())
    sizes = list(type_count.values())

    selected_chart = chart_type.get()
    plt.close('all')

    if selected_chart == "Pie Chart":
        plt.figure(figsize=(8, 6))
        colors = ['#4CAF50', '#2196F3', '#FF9800', '#E91E63', '#9C27B0']
        plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140, colors=colors[:len(labels)],
                wedgeprops={'edgecolor': 'white', 'linewidth': 1.5})
        plt.title('Insurance Plan Distribution', fontsize=14, fontweight='bold', pad=15)
        plt.axis('equal')
        plt.tight_layout()
        plt.show()

    elif selected_chart == "Bar Chart":
        plt.figure(figsize=(8, 5))
        bars = plt.bar(labels, sizes, color='#1976D2', edgecolor='black', width=0.5)
        plt.title('Registrations by Insurance Type', fontsize=14, fontweight='bold', pad=15)
        plt.xlabel('Insurance Type', fontsize=12)
        plt.ylabel('Number of Applicants', fontsize=12)
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2.0, yval + 0.1, int(yval), ha='center', va='bottom', fontweight='bold')
        plt.tight_layout()
        plt.show()

    elif selected_chart == "Line Plot":
        plt.figure(figsize=(9, 5))
        plt.plot(labels, sizes, marker='o', linestyle='-', color='#00796B', linewidth=2.5, markersize=8)
        plt.title('Applicant Trend across Insurance Types', fontsize=14, fontweight='bold', pad=15)
        plt.xlabel('Insurance Type', fontsize=12)
        plt.ylabel('Number of Applicants', fontsize=12)
        plt.grid(True, linestyle='--', alpha=0.6)
        plt.tight_layout()
        plt.show()

def exit_form():
    """Safely close the application."""
    root.destroy()

# --- Initialize GUI ---
root = Tk()
root.title("Insurance Registration System")
root.geometry("660x730")
root.resizable(False, False)

# Center window on screen
window_width, window_height = 660, 730
screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
pos_x = int(screen_width / 2 - window_width / 2)
pos_y = int(screen_height / 2 - window_height / 2)
root.geometry(f"{window_width}x{window_height}+{pos_x}+{pos_y}")

load_or_create_workbook()

# Heading
Label(root, text="🛡️ Insurance Registration Form", font=("Arial", 18, "bold"), fg="#1A237E").pack(pady=12)

# Form Fields Frame
form_frame = Frame(root)
form_frame.pack(pady=5, padx=20, fill='x')

entry_font = ("Arial", 12)
labels = [
    ("Registration ID", "reg_id"),
    ("Full Name", "name"),
    ("Email Address", "email"),
    ("Phone Number", "phone"),
    ("Address", "address"),
    ("Date of Birth (YYYY-MM-DD)", "dob")
]

for label_text, key in labels:
    row_frame = Frame(form_frame)
    row_frame.pack(fill='x', pady=4)
    Label(row_frame, text=label_text, font=entry_font, width=22, anchor='e').pack(side='left', padx=10)
    entry = Entry(row_frame, font=entry_font)
    entry.pack(side='left', padx=10, fill='x', expand=True)
    globals()[f"{key}_field"] = entry

# Gender Radio Buttons
gender_frame = Frame(form_frame)
gender_frame.pack(fill='x', pady=4)
Label(gender_frame, text="Gender", font=entry_font, width=22, anchor='e').pack(side='left', padx=10)
gender = StringVar()
Radiobutton(gender_frame, text="Male", variable=gender, value="Male", font=entry_font).pack(side='left', padx=8)
Radiobutton(gender_frame, text="Female", variable=gender, value="Female", font=entry_font).pack(side='left', padx=8)
Radiobutton(gender_frame, text="Other", variable=gender, value="Other", font=entry_font).pack(side='left', padx=8)

# Age Range Dropdown
age_frame = Frame(form_frame)
age_frame.pack(fill='x', pady=4)
Label(age_frame, text="Age Group", font=entry_font, width=22, anchor='e').pack(side='left', padx=10)
age = StringVar()
age_options = ["18-25", "26-35", "36-45", "46-55", "56-65", "65+"]
age_dropdown = OptionMenu(age_frame, age, *age_options)
age_dropdown.config(font=entry_font, width=20)
age_dropdown.pack(side='left', padx=10, fill='x', expand=True)

# Insurance Type Dropdown
type_frame = Frame(form_frame)
type_frame.pack(fill='x', pady=4)
Label(type_frame, text="Insurance Type", font=entry_font, width=22, anchor='e').pack(side='left', padx=10)
insurance_type = StringVar()
type_options = ["Health", "Life", "Travel", "Vehicle", "Property"]
type_dropdown = OptionMenu(type_frame, insurance_type, *type_options)
type_dropdown.config(font=entry_font, width=20)
type_dropdown.pack(side='left', padx=10, fill='x', expand=True)

# Agree to Terms Checkbox
terms_frame = Frame(form_frame)
terms_frame.pack(fill='x', pady=8)
Label(terms_frame, text="Agreement", font=entry_font, width=22, anchor='e').pack(side='left', padx=10)
agree_terms = IntVar()
Checkbutton(terms_frame, text="I agree to the terms and conditions", variable=agree_terms, font=entry_font).pack(side='left', padx=10)

# Chart Selection
chart_frame = Frame(root, bd=1, relief='groove', padx=10, pady=6)
chart_frame.pack(pady=10, padx=25, fill='x')
Label(chart_frame, text="📊 Analytics Visualization:", font=("Arial", 12, "bold")).pack(side='left', padx=5)
chart_type = StringVar(value="Pie Chart")
chart_menu = OptionMenu(chart_frame, chart_type, "Pie Chart", "Bar Chart", "Line Plot")
chart_menu.config(font=entry_font, width=15)
chart_menu.pack(side='left', padx=10)
Button(chart_frame, text="Generate Chart", command=show_chart, font=("Arial", 11, "bold"), bg="#E0E0E0", padx=10).pack(side='left', padx=5)

# Action Buttons
button_frame = Frame(root)
button_frame.pack(pady=15)
Button(button_frame, text="Submit", command=insert, font=("Arial", 13, "bold"), bg="#4CAF50", fg="white", width=10).pack(side='left', padx=8)
Button(button_frame, text="Clear", command=clear, font=("Arial", 13), bg="#FF9800", fg="white", width=10).pack(side='left', padx=8)
Button(button_frame, text="Exit", command=exit_form, font=("Arial", 13), bg="#E53935", fg="white", width=10).pack(side='left', padx=8)

if __name__ == "__main__":
    root.mainloop()
