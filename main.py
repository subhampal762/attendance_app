import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime, date
import calendar
import pandas as pd
import os

# --- Configuration ---
COMPANY_NAME = "Shree Padmavati Engineers India Private Limited"
PROJECT_NAME = "Infosys Limited Noida (Sector-85)"
DB_NAME = "staff_attendance.db"

# Pre-defined Holidays (Dec 2025 to Dec 2026)
HOLIDAYS = {
    "2025-12-25": "Christmas",
    "2026-01-14": "Makar Sankranti",
    "2026-01-26": "Republic Day",
    "2026-03-06": "Holi",
    "2026-03-25": "Ram Navami",
    "2026-04-14": "Ambedkar Jayanti",
    "2026-08-15": "Independence Day",
    "2026-08-27": "Raksha Bandhan",
    "2026-10-02": "Gandhi Jayanti",
    "2026-10-20": "Dussehra",
    "2026-11-08": "Diwali",
    "2026-12-25": "Christmas"
}

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Attendance System - {COMPANY_NAME}")
        self.root.geometry("1000x700")

        # Database Setup
        self.conn = sqlite3.connect(DB_NAME)
        self.cursor = self.conn.cursor()
        self.create_tables()

        # Header
        header_frame = tk.Frame(self.root, bg="#2c3e50", pady=10)
        header_frame.pack(fill=tk.X)
        
        tk.Label(header_frame, text=COMPANY_NAME, font=("Arial", 16, "bold"), bg="#2c3e50", fg="white").pack()
        tk.Label(header_frame, text=f"Project: {PROJECT_NAME}", font=("Arial", 12), bg="#2c3e50", fg="#ecf0f1").pack()

        # Tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.tab_employees = ttk.Frame(self.notebook)
        self.tab_attendance = ttk.Frame(self.notebook)
        self.tab_report = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_employees, text="Manage Staff")
        self.notebook.add(self.tab_attendance, text="Mark Attendance")
        self.notebook.add(self.tab_report, text="Attendance Report & Excel")

        self.build_employee_tab()
        self.build_attendance_tab()
        self.build_report_tab()

    def create_tables(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS employees (
                emp_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                mobile TEXT
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS attendance (
                date TEXT,
                emp_id TEXT,
                status TEXT,
                PRIMARY KEY (date, emp_id),
                FOREIGN KEY(emp_id) REFERENCES employees(emp_id)
            )
        ''')
        self.conn.commit()

    # ---------------- TAB 1: EMPLOYEE MANAGEMENT ----------------
    def build_employee_tab(self):
        frame = ttk.LabelFrame(self.tab_employees, text="Add New Employee")
        frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(frame, text="Employee ID:").grid(row=0, column=0, padx=5, pady=5)
        self.ent_emp_id = ttk.Entry(frame)
        self.ent_emp_id.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Full Name:").grid(row=0, column=2, padx=5, pady=5)
        self.ent_emp_name = ttk.Entry(frame)
        self.ent_emp_name.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame, text="Mobile No:").grid(row=0, column=4, padx=5, pady=5)
        self.ent_emp_mobile = ttk.Entry(frame)
        self.ent_emp_mobile.grid(row=0, column=5, padx=5, pady=5)

        ttk.Button(frame, text="Add Staff", command=self.add_employee).grid(row=0, column=6, padx=10, pady=5)

        tree_frame = ttk.Frame(self.tab_employees)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("ID", "Name", "Mobile")
        self.emp_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.emp_tree.heading("ID", text="Employee ID")
        self.emp_tree.heading("Name", text="Staff Name")
        self.emp_tree.heading("Mobile", text="Mobile No")
        self.emp_tree.pack(fill="both", expand=True)
        
        self.load_employees()

    def add_employee(self):
        eid = self.ent_emp_id.get()
        name = self.ent_emp_name.get()
        mobile = self.ent_emp_mobile.get()

        if not eid or not name:
            messagebox.showerror("Error", "ID and Name are required.")
            return

        try:
            self.cursor.execute("INSERT INTO employees VALUES (?, ?, ?)", (eid, name, mobile))
            self.conn.commit()
            messagebox.showinfo("Success", "Staff added successfully!")
            self.ent_emp_id.delete(0, tk.END)
            self.ent_emp_name.delete(0, tk.END)
            self.ent_emp_mobile.delete(0, tk.END)
            self.load_employees()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Employee ID already exists!")

    def load_employees(self):
        for row in self.emp_tree.get_children():
            self.emp_tree.delete(row)
        self.cursor.execute("SELECT * FROM employees")
        for row in self.cursor.fetchall():
            self.emp_tree.insert("", tk.END, values=row)

    # ---------------- TAB 2: MARK ATTENDANCE ----------------
    def build_attendance_tab(self):
        control_frame = ttk.Frame(self.tab_attendance)
        control_frame.pack(pady=10)

        ttk.Label(control_frame, text="Select Date:").pack(side=tk.LEFT, padx=5)
        
        years = [str(y) for y in range(2025, 2028)]
        months = [str(m).zfill(2) for m in range(1, 13)]
        days = [str(d).zfill(2) for d in range(1, 32)]

        self.cb_year = ttk.Combobox(control_frame, values=years, width=5)
        self.cb_year.set(str(datetime.now().year))
        self.cb_year.pack(side=tk.LEFT, padx=2)

        self.cb_month = ttk.Combobox(control_frame, values=months, width=4)
        self.cb_month.set(str(datetime.now().month).zfill(2))
        self.cb_month.pack(side=tk.LEFT, padx=2)

        self.cb_day = ttk.Combobox(control_frame, values=days, width=4)
        self.cb_day.set(str(datetime.now().day).zfill(2))
        self.cb_day.pack(side=tk.LEFT, padx=2)

        ttk.Button(control_frame, text="Load Attendance Sheet", command=self.load_attendance_sheet).pack(side=tk.LEFT, padx=10)
        ttk.Button(control_frame, text="Save Attendance", command=self.save_attendance).pack(side=tk.LEFT, padx=10)

        self.canvas = tk.Canvas(self.tab_attendance)
        self.scrollbar = ttk.Scrollbar(self.tab_attendance, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=10)
        self.scrollbar.pack(side="right", fill="y")

        self.attendance_vars = {} 

    def load_attendance_sheet(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.attendance_vars = {}

        selected_date_str = f"{self.cb_year.get()}-{self.cb_month.get()}-{self.cb_day.get()}"
        
        try:
            dt_obj = datetime.strptime(selected_date_str, "%Y-%m-%d")
            is_sunday = dt_obj.weekday() == 6
            holiday_name = HOLIDAYS.get(selected_date_str)
        except ValueError:
            messagebox.showerror("Error", "Invalid Date Selected")
            return

        tk.Label(self.scrollable_frame, text="ID", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=10, pady=5)
        tk.Label(self.scrollable_frame, text="Name", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=10, pady=5)
        tk.Label(self.scrollable_frame, text="Status", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=10, pady=5)
        
        status_info = ""
        default_status = "P"
        
        if is_sunday:
            status_info = " (Sunday)"
            default_status = "WO"
        elif holiday_name:
            status_info = f" ({holiday_name})"
            default_status = "H"

        tk.Label(self.scrollable_frame, text=f"Date: {selected_date_str}{status_info}", fg="blue").grid(row=0, column=3, padx=10)

        self.cursor.execute("SELECT emp_id, name FROM employees")
        employees = self.cursor.fetchall()

        if not employees:
            tk.Label(self.scrollable_frame, text="No employees found. Add staff first.").grid(row=1, column=1)
            return

        self.cursor.execute("SELECT emp_id, status FROM attendance WHERE date=?", (selected_date_str,))
        existing_data = dict(self.cursor.fetchall())

        r = 1
        for emp_id, name in employees:
            tk.Label(self.scrollable_frame, text=emp_id).grid(row=r, column=0, padx=5, pady=2)
            tk.Label(self.scrollable_frame, text=name).grid(row=r, column=1, padx=5, pady=2, sticky="w")

            var = tk.StringVar(value=existing_data.get(emp_id, default_status))
            self.attendance_vars[emp_id] = var

            frame_radios = tk.Frame(self.scrollable_frame)
            frame_radios.grid(row=r, column=2, padx=5, pady=2)

            # Show P and L options ALWAYS (even on Sundays)
            tk.Radiobutton(frame_radios, text="P", variable=var, value="P", bg="#dff9fb").pack(side=tk.LEFT, padx=2)
            tk.Radiobutton(frame_radios, text="L", variable=var, value="L", bg="#ff7979").pack(side=tk.LEFT, padx=2)
            
            # Show WO or H option ONLY if it applies to that day
            if is_sunday:
                tk.Radiobutton(frame_radios, text="WO", variable=var, value="WO", bg="#ecf0f1").pack(side=tk.LEFT, padx=2)
            elif holiday_name:
                tk.Radiobutton(frame_radios, text="H", variable=var, value="H", bg="#ecf0f1").pack(side=tk.LEFT, padx=2)
            
            r += 1

    def save_attendance(self):
        date_str = f"{self.cb_year.get()}-{self.cb_month.get()}-{self.cb_day.get()}"
        if not self.attendance_vars:
            return

        try:
            for emp_id, var in self.attendance_vars.items():
                status = var.get()
                self.cursor.execute('''
                    INSERT OR REPLACE INTO attendance (date, emp_id, status)
                    VALUES (?, ?, ?)
                ''', (date_str, emp_id, status))
            
            self.conn.commit()
            messagebox.showinfo("Success", f"Attendance saved for {date_str}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------------- TAB 3: REPORT & EXCEL ----------------
    def build_report_tab(self):
        frame = ttk.Frame(self.tab_report)
        frame.pack(pady=20)

        ttk.Label(frame, text="Report Month:").pack(side=tk.LEFT, padx=5)
        
        self.rep_year = ttk.Combobox(frame, values=[str(y) for y in range(2025, 2028)], width=6)
        self.rep_year.set(str(datetime.now().year))
        self.rep_year.pack(side=tk.LEFT)

        self.rep_month = ttk.Combobox(frame, values=[str(m).zfill(2) for m in range(1, 13)], width=4)
        self.rep_month.set(str(datetime.now().month).zfill(2))
        self.rep_month.pack(side=tk.LEFT)

        ttk.Button(frame, text="Generate Excel Report", command=self.generate_excel_report).pack(side=tk.LEFT, padx=15)
        self.lbl_status = ttk.Label(self.tab_report, text="Ready to generate report.", foreground="gray")
        self.lbl_status.pack(pady=10)

    def generate_excel_report(self):
        year = self.rep_year.get()
        month = self.rep_month.get()
        month_str = f"{year}-{month}"

        self.cursor.execute("SELECT emp_id, name FROM employees")
        employees = self.cursor.fetchall()
        
        if not employees:
            messagebox.showerror("Error", "No employees found.")
            return

        _, num_days = calendar.monthrange(int(year), int(month))
        data = []

        for emp_id, name in employees:
            row_data = {'ID': emp_id, 'Name': name}
            p_count = 0
            l_count = 0
            
            self.cursor.execute('''
                SELECT substr(date, 9, 2) as day, status 
                FROM attendance 
                WHERE emp_id=? AND date LIKE ?
            ''', (emp_id, f"{month_str}%"))
            
            att_records = dict(self.cursor.fetchall())

            for day in range(1, num_days + 1):
                day_str = str(day).zfill(2)
                date_full = f"{year}-{month}-{day_str}"
                
                status = att_records.get(day_str, "")
                
                if not status:
                    dt_obj = date(int(year), int(month), day)
                    if dt_obj.weekday() == 6:
                        status = "WO"
                    elif date_full in HOLIDAYS:
                        status = "H"
                    else:
                        status = "-"

                row_data[day_str] = status

                if status == 'P': p_count += 1
                elif status == 'L': l_count += 1
            
            row_data['Total Present'] = p_count
            row_data['Total Absent'] = l_count
            data.append(row_data)

        df = pd.DataFrame(data)
        file_name = f"Attendance_Report_{month_str}.xlsx"
        
        try:
            writer = pd.ExcelWriter(file_name, engine='openpyxl')
            df.to_excel(writer, sheet_name='Attendance Report', index=False)
            
            worksheet = writer.sheets['Attendance Report']
            for column_cells in worksheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                worksheet.column_dimensions[column_cells[0].column_letter].width = length + 2

            writer.close()
            
            self.lbl_status.config(text=f"Report Saved: {os.path.abspath(file_name)}", foreground="green")
            messagebox.showinfo("Success", f"Excel generated successfully!\nFile: {file_name}")
            try: os.startfile(file_name)
            except: pass

        except Exception as e:
            messagebox.showerror("Error Exporting", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
