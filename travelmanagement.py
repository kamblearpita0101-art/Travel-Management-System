import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
import os


FILE_NAME = "travel_records.xlsx"


if not os.path.exists(FILE_NAME):
    wb = Workbook()
    ws = wb.active
    ws.title = "Travel Records"

    ws.append([
        "Customer ID",
        "Customer Name",
        "Contact",
        "Destination",
        "Travel Date",
        "Persons",
        "Package",
        "Total Cost"
    ])

    wb.save(FILE_NAME)




def login():
    username = username_entry.get()
    password = password_entry.get()

    if username == "admin" and password == "1234":
        login_window.destroy()
        dashboard()
    else:
        messagebox.showerror("Login Failed", "Invalid Username or Password")


login_window = tk.Tk()
login_window.title("Travel Management System - Login")
login_window.geometry("400x300")
login_window.resizable(False, False)

tk.Label(
    login_window,
    text="TRAVEL MANAGEMENT SYSTEM",
    font=("Arial", 16, "bold")
).pack(pady=25)

tk.Label(login_window, text="Username").pack()
username_entry = tk.Entry(login_window, width=30)
username_entry.pack(pady=5)

tk.Label(login_window, text="Password").pack()
password_entry = tk.Entry(login_window, width=30, show="*")
password_entry.pack(pady=5)

tk.Button(
    login_window,
    text="LOGIN",
    width=15,
    command=login
).pack(pady=20)




def dashboard():

    root = tk.Tk()
    root.title("Travel Management System")
    root.geometry("1100x650")
    root.resizable(False, False)

    
    tk.Label(
        root,
        text="TRAVEL MANAGEMENT SYSTEM",
        font=("Arial", 24, "bold")
    ).pack(pady=15)

    tk.Label(
        root,
        text="Manage Customer Travel Records",
        font=("Arial", 12)
    ).pack()

   
    form_frame = tk.Frame(root)
    form_frame.pack(pady=20)

   
    customer_id = tk.StringVar()
    customer_name = tk.StringVar()
    contact = tk.StringVar()
    destination = tk.StringVar()
    travel_date = tk.StringVar()
    persons = tk.StringVar()
    package = tk.StringVar()
    total_cost = tk.StringVar()

   

    tk.Label(form_frame, text="Customer ID").grid(row=0, column=0, padx=10, pady=8)
    tk.Entry(form_frame, textvariable=customer_id, width=25).grid(row=0, column=1)

    tk.Label(form_frame, text="Customer Name").grid(row=0, column=2, padx=10)
    tk.Entry(form_frame, textvariable=customer_name, width=25).grid(row=0, column=3)

    tk.Label(form_frame, text="Contact").grid(row=1, column=0, padx=10, pady=8)
    tk.Entry(form_frame, textvariable=contact, width=25).grid(row=1, column=1)

    tk.Label(form_frame, text="Destination").grid(row=1, column=2, padx=10)
    tk.Entry(form_frame, textvariable=destination, width=25).grid(row=1, column=3)

    tk.Label(form_frame, text="Travel Date").grid(row=2, column=0, padx=10, pady=8)
    tk.Entry(form_frame, textvariable=travel_date, width=25).grid(row=2, column=1)

    tk.Label(form_frame, text="No. of Persons").grid(row=2, column=2, padx=10)
    tk.Entry(form_frame, textvariable=persons, width=25).grid(row=2, column=3)

    tk.Label(form_frame, text="Package").grid(row=3, column=0, padx=10, pady=8)
    package_combo = ttk.Combobox(
        form_frame,
        textvariable=package,
        values=[
            "Basic",
            "Standard",
            "Premium"
        ],
        width=22,
        state="readonly"
    )
    package_combo.grid(row=3, column=1)

    tk.Label(form_frame, text="Total Cost").grid(row=3, column=2, padx=10)
    tk.Entry(form_frame, textvariable=total_cost, width=25).grid(row=3, column=3)



    def clear_fields():
        customer_id.set("")
        customer_name.set("")
        contact.set("")
        destination.set("")
        travel_date.set("")
        persons.set("")
        package.set("")
        total_cost.set("")



    def add_record():

        if (
            customer_id.get() == "" or
            customer_name.get() == "" or
            contact.get() == "" or
            destination.get() == "" or
            travel_date.get() == "" or
            persons.get() == "" or
            package.get() == "" or
            total_cost.get() == ""
        ):
            messagebox.showwarning(
                "Warning",
                "Please fill all fields"
            )
            return

        wb = load_workbook(FILE_NAME)
        ws = wb.active
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[0]) == customer_id.get():
                messagebox.showerror(
                    "Error",
                    "Customer ID already exists"
                )
                wb.close()
                return

        ws.append([
            customer_id.get(),
            customer_name.get(),
            contact.get(),
            destination.get(),
            travel_date.get(),
            persons.get(),
            package.get(),
            total_cost.get()
        ])

        wb.save(FILE_NAME)
        wb.close()

        messagebox.showinfo(
            "Success",
            "Travel record added successfully"
        )

        clear_fields()
        view_records()


    def view_records():

        for item in tree.get_children():
            tree.delete(item)

        wb = load_workbook(FILE_NAME)
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):
            tree.insert("", tk.END, values=row)

        wb.close()

    def search_record():

        search_id = customer_id.get()

        if search_id == "":
            messagebox.showwarning(
                "Warning",
                "Enter Customer ID to search"
            )
            return

        for item in tree.get_children():
            tree.delete(item)

        wb = load_workbook(FILE_NAME)
        ws = wb.active

        found = False

        for row in ws.iter_rows(min_row=2, values_only=True):

            if str(row[0]) == search_id:

                tree.insert("", tk.END, values=row)

                customer_name.set(row[1])
                contact.set(row[2])
                destination.set(row[3])
                travel_date.set(row[4])
                persons.set(row[5])
                package.set(row[6])
                total_cost.set(row[7])

                found = True
                break

        wb.close()

        if not found:
            messagebox.showinfo(
                "Search",
                "Customer record not found"
            )

    def select_record(event):

        selected = tree.focus()

        if selected == "":
            return

        values = tree.item(selected, "values")

        if values:
            customer_id.set(values[0])
            customer_name.set(values[1])
            contact.set(values[2])
            destination.set(values[3])
            travel_date.set(values[4])
            persons.set(values[5])
            package.set(values[6])
            total_cost.set(values[7])

    def update_record():

        if customer_id.get() == "":
            messagebox.showwarning(
                "Warning",
                "Select a record first"
            )
            return

        wb = load_workbook(FILE_NAME)
        ws = wb.active

        found = False

        for row in ws.iter_rows(min_row=2):

            if str(row[0].value) == customer_id.get():

                row[1].value = customer_name.get()
                row[2].value = contact.get()
                row[3].value = destination.get()
                row[4].value = travel_date.get()
                row[5].value = persons.get()
                row[6].value = package.get()
                row[7].value = total_cost.get()

                found = True
                break

        wb.save(FILE_NAME)
        wb.close()

        if found:
            messagebox.showinfo(
                "Success",
                "Record updated successfully"
            )
            clear_fields()
            view_records()
        else:
            messagebox.showerror(
                "Error",
                "Record not found"
            )


    def delete_record():

        selected = tree.focus()

        if selected == "":
            messagebox.showwarning(
                "Warning",
                "Please select a record to delete"
            )
            return

        values = tree.item(selected, "values")

        confirm = messagebox.askyesno(
            "Confirm Delete",
            "Are you sure you want to delete this record?"
        )

        if not confirm:
            return

        delete_id = values[0]

        wb = load_workbook(FILE_NAME)
        ws = wb.active

        for row in range(2, ws.max_row + 1):

            if str(ws.cell(row=row, column=1).value) == delete_id:
                ws.delete_rows(row, 1)
                break

        wb.save(FILE_NAME)
        wb.close()

        messagebox.showinfo(
            "Success",
            "Record deleted successfully"
        )

        clear_fields()
        view_records()


    def logout():

        answer = messagebox.askyesno(
            "Logout",
            "Do you want to logout?"
        )

        if answer:
            root.destroy()

            global login_window

            login_window = tk.Tk()
            login_window.title("Travel Management System - Login")
            login_window.geometry("400x300")
            login_window.resizable(False, False)

            tk.Label(
                login_window,
                text="TRAVEL MANAGEMENT SYSTEM",
                font=("Arial", 16, "bold")
            ).pack(pady=25)

            tk.Label(
                login_window,
                text="Username"
            ).pack()

            global username_entry
            username_entry = tk.Entry(
                login_window,
                width=30
            )
            username_entry.pack(pady=5)

            tk.Label(
                login_window,
                text="Password"
            ).pack()

            global password_entry
            password_entry = tk.Entry(
                login_window,
                width=30,
                show="*"
            )
            password_entry.pack(pady=5)

            tk.Button(
                login_window,
                text="LOGIN",
                width=15,
                command=login
            ).pack(pady=20)

            login_window.mainloop()

    

    button_frame = tk.Frame(root)
    button_frame.pack(pady=5)

    tk.Button(
        button_frame,
        text="ADD RECORD",
        width=15,
        command=add_record
    ).grid(row=0, column=0, padx=5)

    tk.Button(
        button_frame,
        text="VIEW RECORDS",
        width=15,
        command=view_records
    ).grid(row=0, column=1, padx=5)

    tk.Button(
        button_frame,
        text="SEARCH",
        width=15,
        command=search_record
    ).grid(row=0, column=2, padx=5)

    tk.Button(
        button_frame,
        text="UPDATE",
        width=15,
        command=update_record
    ).grid(row=0, column=3, padx=5)

    tk.Button(
        button_frame,
        text="DELETE",
        width=15,
        command=delete_record
    ).grid(row=0, column=4, padx=5)

    tk.Button(
        button_frame,
        text="CLEAR",
        width=15,
        command=clear_fields
    ).grid(row=0, column=5, padx=5)

    tk.Button(
        button_frame,
        text="LOGOUT",
        width=15,
        command=logout
    ).grid(row=0, column=6, padx=5)

    

    tree_frame = tk.Frame(root)
    tree_frame.pack(pady=20)

    columns = (
        "Customer ID",
        "Customer Name",
        "Contact",
        "Destination",
        "Travel Date",
        "Persons",
        "Package",
        "Total Cost"
    )

    tree = ttk.Treeview(
        tree_frame,
        columns=columns,
        show="headings",
        height=12
    )

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=125)

    tree.pack()

    tree.bind(
        "<ButtonRelease-1>",
        select_record
    )

    view_records()

    root.mainloop()

login_window.mainloop()
