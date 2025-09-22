import tkinter as tk
from tkinter import ttk,messagebox,filedialog
import sqlite3
import csv
import openpyxl

def connect_db():
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute(
        """CREATE TABLE IF NOT EXISTS students(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        roll_no TEXT UNIQUE,
        name TEXT,
        department TEXT,
        year TEXT )
        """
    )
    conn.commit()
    conn.close()

def create_admin():
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS admin(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    password TEXT )"""
    )

    try:
        cur.execute("INSERT INTO admin(username,password) VALUES(?,?)",("admin","admin123"))
    except sqlite3.IntegrityError:
        pass
    conn.commit()
    conn.close()

def insert_student(roll_no,name,department,year):
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    try:
        cur.execute("INSERT INTO students(roll_no,name,department,year) VALUES (?,?,?,?)",
                    (roll_no, name, department, year))
        conn.commit()
        messagebox.showinfo("Success", "Student Added Successfully!")
    except sqlite3.IntegrityError:
        messagebox.showerror("Error","Roll Number Already Exists!")
        conn.close()
        fetch_students()

def fetch_students():
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute("SELECT*FROM students")
    rows = cur.fetchall()
    conn.close()

    for row in tree.get_children():
        tree.delete(row)
    for row in rows:
        tree.insert("",tk.END,values=row)


def delete_student():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning","please select a student to delete.")
        return
    student_id = tree.item(selected)["values"][0]
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute("DELETE FROM students WHERE id=?",(student_id,))
    conn.commit()
    conn.close()
    fetch_students()
    messagebox.showinfo("Deleted","Student deleted successfully.")


def update_student():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Warning","please select student to update.")
        return
    student_id = tree.item(selected)["values"][0]
    roll_no = entry_roll.get()
    name = entry_name.get()
    department = entry_department.get()
    year = entry_year.get()

    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute("""
    UPDATE students SET roll_no=?,name=?,department=?,year=? WHERE id=?""",
    (roll_no,name,department,year,student_id))
    conn.commit()
    conn.close()
    fetch_students()
    messagebox.showinfo("Updated","Student updated successfully.")

def search_student():
        keyword = entry_search.get()
        if keyword == "":
            messagebox.showwarning("Warning","Please enter Roll No or Name to search.")
            return
        conn = sqlite3.connect("students.db")
        cur = conn.cursor()
        cur.execute("""
        SELECT*FROM students
        WHERE roll_no LIKE ? OR name LIKE ?
        """,('%'+keyword+'%','%'+keyword+'%'))
        rows = cur.fetchall()
        conn.close()

        for row in tree.get_children():
            tree.delete(row)
        for row in rows:
            tree.insert("",tk.END,values=row)

def export_csv():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV Files","*.csv")])
    if not file_path:
        return
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute("SELECT roll_no,name,department,year FROM students")
    rows = cur.fetchall()
    conn.close()

    with open(file_path,"w",newline="",encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Roll No","Name","Department","Year"])
        writer.writerows(rows)
    messagebox.showinfo("Success",f"Data exported to {file_path}")


def export_excel():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel Files","*.xlsx")])
    if not file_path:
        return
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute("SELECT roll_no,name,department,year FROM students")
    rows = cur.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title ="Students"

    headers = ["Roll No","Name","Department","Year"]
    ws.append(headers)

    for row in rows:
        ws.append(row)
    wb.save(file_path)
    messagebox.showinfo("Success",f"Data exported to {file_path}")

def login():
    username = entry_user.get()
    password = entry_pass.get()
    conn = sqlite3.connect("students.db")
    cur = conn.cursor()
    cur.execute("SELECT*FROM admin WHERE username =? AND password = ?",(username,password))
    row = cur.fetchone()
    conn.close()

    if row:
        login_window.destroy()
        main_app()
    else:
        messagebox.showerror("Error","Invalid Username or Password")


def show_login():
    global login_window,entry_user,entry_pass
    login_window = tk.Tk()
    login_window.title("Login")
    login_window.geometry("300x200")
    login_window.configure(bg="#f0f4f7")

    tk.Label(login_window,text="Username",font=("Arial",11)).pack(pady=5)
    entry_user = tk.Entry(login_window)
    entry_user.pack(pady=5)

    tk.Label(login_window,text="Password",font=("Arial",11)).pack(pady=5)
    entry_pass = tk.Entry(login_window,show="*")
    entry_pass.pack(pady=5)

    ttk.Button(login_window,text="Login",command=login).pack(pady=10)
    login_window.mainloop()


def main_app():
    global root,tree,entry_roll,entry_name,entry_department,entry_year,entry_search

    root = tk.Tk()
    root.title("College Student Management System ")
    root.geometry("850x600")
    root.configure(bg="#f0f4f7")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview",
                    font = ("Arial",11),
                    rowheight =28,
                    background="#ffffff",
                    fieldbackground="#f9f9f9")
    style.configure("Treeview.Heading",
                    font = ("Arial",12,"bold"),
                    foreground = "white",
                    background = "#007acc")
    style.map("Treeview",background=[("selected","#3399ff")])

    style.configure("TButton",
                    font = ("Arial",11,"bold"),
                    padding=6)
    style.map("TButton",
              background=[("active","#005f99")])

    tk.Label(root,text="Roll No",bg="#f0f4f7",font=("Arial",11)).grid(row=0,column=0,padx=10,pady=5)
    entry_roll = tk.Entry(root)
    entry_roll.grid(row=0,column=1,padx=10,pady=5)

    tk.Label(root,text="Name",bg="#f0f4f7",font=("Arial",11)).grid(row=1,column=0,padx=10,pady=5)
    entry_name = tk.Entry(root)
    entry_name.grid(row=1,column=1,padx=10,pady=5)

    tk.Label(root, text="Department", bg="#f0f4f7", font=("Arial", 11)).grid(row=2, column=0, padx=10, pady=5)
    entry_department = tk.Entry(root)
    entry_department.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(root, text="Year", bg="#f0f4f7", font=("Arial", 11)).grid(row=3, column=0, padx=10, pady=5)
    entry_year = ttk.Combobox(root,values=["First Year","Second Year","Third Year"])
    entry_year.grid(row=3, column=1, padx=10, pady=5)
    entry_year.current(0)

    tk.Label(root, text="Search (Roll No/Name )", bg="#f0f4f7", font=("Arial", 11)).grid(row=0, column=2, padx=10, pady=5)
    entry_search = tk.Entry(root)
    entry_search.grid(row=0, column=3, padx=10, pady=5)

    ttk.Button(root, text="Search",command=search_student).grid(row=1,column=2,padx=10,pady=5)
    ttk.Button(root,text="Show All",command=fetch_students).grid(row=1,column=3,padx=10,pady=5)

    ttk.Button(root, text="Add Student", command=lambda:insert_student(entry_roll.get(),
                                                                       entry_name.get(),
                                                                       entry_department.get(),
                                                                       entry_year.get())).grid(row=4, column= 0, padx=10, pady=10)
    ttk.Button(root, text="Update Student", command=update_student).grid(row=4, column=1, padx=10, pady=10)
    ttk.Button(root, text="Delete Student", command=delete_student).grid(row=4, column=2, padx=10, pady=10)
    ttk.Button(root, text="Exit", command=root.quit).grid(row=4, column=3, padx=10, pady=10)
    ttk.Button(root, text="Export CSV", command=export_csv).grid(row=6, column=0, padx=10, pady=10)
    ttk.Button(root, text="Export Excel", command=export_excel).grid(row=6, column=1, padx=10, pady=10)

    columns = ("ID","Roll No","Name","Department","Year")
    tree = ttk.Treeview(root,columns=columns,show="headings")
    for col in columns:
        tree.heading(col,text=col)
        tree.column(col,width=150)
    tree.grid(row=5,column=0,columnspan=4,padx=10,pady=20)
    fetch_students()
    root.mainloop()

if __name__ == "__main__":
    connect_db()
    create_admin()
    show_login()
