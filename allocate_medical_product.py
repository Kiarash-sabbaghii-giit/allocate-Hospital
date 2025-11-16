import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import pyodbc
import logging
from datetime import datetime
import sys
import os

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None

#  Configuration
MSSQL_CONN_STR_DEFAULT = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost;"
    "DATABASE=pro_hospital;"
    "Trusted_Connection=yes;"
)
LOG_FILENAME = "medlink2_stent.log"
DEFAULT_ITEM_TYPES = ["Stent", "Generic"]

#  Logging Setup
logger = logging.getLogger("medlink2_inventory")
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(LOG_FILENAME, encoding="utf-8")
fh.setLevel(logging.DEBUG)
fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
fh.setFormatter(fmt)
logger.addHandler(fh)
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
ch.setFormatter(fmt)
logger.addHandler(ch)

# Database Layer
class Database:
    def __init__(self, conn_str=None):
        self.conn_str = conn_str or MSSQL_CONN_STR_DEFAULT
        self.conn = None
        self.connect()
        self.ensure_schema()

    def connect(self):
        try:
            self.conn = pyodbc.connect(self.conn_str, autocommit=False)
            logger.info("Connected to database")
        except Exception as e:
            logger.exception("DB connect failed")
            messagebox.showerror("DB Error", f"Unable to connect to database:\n{e}")
            sys.exit(1)

    def ensure_schema(self):
        cur = self.conn.cursor()
        # hospitals
        cur.execute(
            """
            IF OBJECT_ID('dbo.hospitals','U') IS NULL
            CREATE TABLE hospitals (
                id INT IDENTITY(1,1) PRIMARY KEY,
                name NVARCHAR(200) UNIQUE NOT NULL
            )
            """
        )
        # items
        cur.execute(
            """
            IF OBJECT_ID('dbo.items','U') IS NULL
            CREATE TABLE items (
                id INT IDENTITY(1,1) PRIMARY KEY,
                item_type NVARCHAR(50) NOT NULL,
                name NVARCHAR(200) NOT NULL,
                diameter FLOAT NULL,
                length INT NULL,
                price FLOAT NULL,
                qty INT NOT NULL DEFAULT 0
            )
            """
        )
        # allocations
        cur.execute(
            """
            IF OBJECT_ID('dbo.allocations','U') IS NULL
            CREATE TABLE allocations (
                id INT IDENTITY(1,1) PRIMARY KEY,
                hospital_id INT NOT NULL FOREIGN KEY REFERENCES hospitals(id),
                item_id INT NOT NULL FOREIGN KEY REFERENCES items(id),
                qty INT NOT NULL,
                price_at_alloc FLOAT NULL,
                ts DATETIME NOT NULL DEFAULT(GETDATE())
            )
            """
        )
        self.conn.commit()

        # ensure columns exist (for upgrades)
        # add price to items if missing
        cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='items'")
        cols = [r[0].lower() for r in cur.fetchall()]
        if 'price' not in cols:
            try:
                cur.execute("ALTER TABLE items ADD price FLOAT NULL")
                self.conn.commit()
                logger.debug('Added price column to items')
            except Exception:
                pass
        # add price_at_alloc to allocations if missing
        cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='allocations'")
        cols2 = [r[0].lower() for r in cur.fetchall()]
        if 'price_at_alloc' not in cols2 and 'price_at_alloc' not in cols2:
            try:
                cur.execute("ALTER TABLE allocations ADD price_at_alloc FLOAT NULL")
                self.conn.commit()
                logger.debug('Added price_at_alloc to allocations')
            except Exception:
                pass

        # Insert default hospitals if empty
        cur.execute("SELECT COUNT(*) FROM hospitals")
        if cur.fetchone()[0] == 0:
            default_hospitals = [
                "Day hospital",
                "Bahman hospital",
                "Saman hospital",
                "Kasra hospital",
                "Pars hospital",
                "Baghiatolah hospital",
            ]
            for h in default_hospitals:
                try:
                    cur.execute("INSERT INTO hospitals (name) VALUES (?)", (h,))
                except Exception:
                    pass
            self.conn.commit()
            logger.debug("Default hospitals inserted")

        cur.close()

    # Hospital operations
    def fetch_hospitals(self):
        cur = self.conn.cursor()
        cur.execute("SELECT id, name FROM hospitals ORDER BY name")
        rows = cur.fetchall()
        cur.close()
        return rows

    def add_hospital(self, name):
        cur = self.conn.cursor()
        cur.execute("INSERT INTO hospitals (name) VALUES (?)", (name,))
        self.conn.commit()
        cur.close()
        logger.info(f"Hospital added: {name}")

    def delete_hospital(self, hid):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM hospitals WHERE id=?", (hid,))
        self.conn.commit()
        cur.close()
        logger.info(f"Hospital deleted: id={hid}")

    # Item operations
    def fetch_items(self, item_type_filter=None, name_filter=None):
        cur = self.conn.cursor()
        q = "SELECT id, item_type, name, diameter, length, price, qty FROM items"
        params = []
        clauses = []
        if item_type_filter:
            clauses.append("item_type = ?")
            params.append(item_type_filter)
        if name_filter:
            clauses.append("name LIKE ?")
            params.append(f"%{name_filter}%")
        if clauses:
            q += " WHERE " + " AND ".join(clauses)
        q += " ORDER BY item_type, name, diameter, length"
        cur.execute(q, params)
        rows = cur.fetchall()
        cur.close()
        return rows

    def add_item(self, item_type, name, diameter, length, price, qty):
        cur = self.conn.cursor()
        cur.execute(
            "INSERT INTO items (item_type, name, diameter, length, price, qty) VALUES (?, ?, ?, ?, ?, ?)",
            (item_type, name, diameter if diameter is not None else None, length if length is not None else None, price if price is not None else None, qty or 0),
        )
        self.conn.commit()
        cur.close()
        logger.info(f"Item added: {item_type} | {name} | d={diameter} l={length} price={price} qty={qty}")

    def update_item(self, item_id, name, diameter, length, price):
        cur = self.conn.cursor()
        cur.execute("UPDATE items SET name=?, diameter=?, length=?, price=? WHERE id=?", (name, diameter if diameter is not None else None, length if length is not None else None, price if price is not None else None, item_id))
        self.conn.commit()
        cur.close()
        logger.info(f"Item updated: id={item_id}")

    def set_stock(self, item_id, qty):
        cur = self.conn.cursor()
        cur.execute("UPDATE items SET qty=? WHERE id=?", (qty, item_id))
        self.conn.commit()
        cur.close()
        logger.info(f"Stock set: item_id={item_id} -> {qty}")

    def delete_item(self, item_id):
        cur = self.conn.cursor()
        cur.execute("DELETE FROM items WHERE id=?", (item_id,))
        self.conn.commit()
        cur.close()
        logger.info(f"Item deleted: id={item_id}")

    # Allocation
    def allocate(self, item_id, hospital_id, qty, price_at_alloc=None):
        cur = self.conn.cursor()
        try:
            cur.execute("SELECT qty, price FROM items WHERE id=?", (item_id,))
            row = cur.fetchone()
            if row is None:
                raise ValueError("Item not found")
            avail = int(row[0])
            current_price = float(row[1]) if row[1] is not None else None
            if price_at_alloc is None:
                price_at_alloc = current_price
            if qty <= 0:
                raise ValueError("Quantity must be > 0")
            if qty > avail:
                raise ValueError(f"Not enough stock (available={avail})")
            cur.execute("UPDATE items SET qty = qty - ? WHERE id=?", (qty, item_id))
            cur.execute("INSERT INTO allocations (hospital_id, item_id, qty, price_at_alloc, ts) VALUES (?, ?, ?, ?, ?)", (hospital_id, item_id, qty, price_at_alloc, datetime.now()))
            self.conn.commit()
            logger.info(f"Allocated qty={qty} of item_id={item_id} to hospital_id={hospital_id} price={price_at_alloc}")
        except Exception as e:
            self.conn.rollback()
            logger.exception("Allocation failed")
            raise
        finally:
            cur.close()

    def fetch_recent_allocations(self, limit=100):
        cur = self.conn.cursor()
        cur.execute(
            "SELECT a.id, h.name, i.item_type, i.name, i.diameter, i.length, a.qty, a.price_at_alloc, a.ts FROM allocations a JOIN hospitals h ON a.hospital_id=h.id JOIN items i ON a.item_id=i.id ORDER BY a.ts DESC"
        )
        rows = cur.fetchmany(limit)
        cur.close()
        return rows

    def fetch_allocations_for_report(self, hospital_id=None, date_from=None, date_to=None):
        cur = self.conn.cursor()
        q = "SELECT a.id, h.name, i.item_type, i.name, i.diameter, i.length, a.qty, a.price_at_alloc, a.ts FROM allocations a JOIN hospitals h ON a.hospital_id=h.id JOIN items i ON a.item_id=i.id"
        params = []
        clauses = []
        if hospital_id:
            clauses.append('h.id = ?')
            params.append(hospital_id)
        if date_from:
            clauses.append('a.ts >= ?')
            params.append(date_from)
        if date_to:
            clauses.append('a.ts <= ?')
            params.append(date_to)
        if clauses:
            q += ' WHERE ' + ' AND '.join(clauses)
        q += ' ORDER BY a.ts ASC'
        cur.execute(q, params)
        rows = cur.fetchall()
        cur.close()
        return rows

    def close(self):
        try:
            if self.conn:
                self.conn.close()
        except Exception:
            pass

# UI Layer
class AddItemDialog(simpledialog.Dialog):
    def __init__(self, parent, title, item_types, item=None):
        self.item_types = item_types
        self.item = item
        super().__init__(parent, title=title)

    def body(self, master):
        ttk.Label(master, text="Item Type:").grid(row=0, column=0, sticky='e')
        self.type_cb = ttk.Combobox(master, values=self.item_types, width=20)
        self.type_cb.grid(row=0, column=1, sticky='w')

        ttk.Label(master, text="Name:").grid(row=1, column=0, sticky='e')
        self.name_e = ttk.Entry(master, width=30)
        self.name_e.grid(row=1, column=1, sticky='w')

        ttk.Label(master, text="Diameter (mm) [optional]:").grid(row=2, column=0, sticky='e')
        self.dia_e = ttk.Entry(master, width=10)
        self.dia_e.grid(row=2, column=1, sticky='w')

        ttk.Label(master, text="Length (mm) [optional]:").grid(row=3, column=0, sticky='e')
        self.len_e = ttk.Entry(master, width=10)
        self.len_e.grid(row=3, column=1, sticky='w')

        ttk.Label(master, text="Price (USD) [optional]:").grid(row=4, column=0, sticky='e')
        self.price_e = ttk.Entry(master, width=15)
        self.price_e.grid(row=4, column=1, sticky='w')

        ttk.Label(master, text="Initial Qty:").grid(row=5, column=0, sticky='e')
        self.qty_e = ttk.Entry(master, width=10)
        self.qty_e.grid(row=5, column=1, sticky='w')

        # If editing existing
        if self.item:
            # item: (id, item_type, name, diameter, length, price, qty)
            _, t, name, d, l, p, q = self.item
            self.type_cb.set(t)
            self.name_e.insert(0, name)
            if d is not None:
                self.dia_e.insert(0, str(d))
            if l is not None:
                self.len_e.insert(0, str(l))
            if p is not None:
                self.price_e.insert(0, str(p))
            self.qty_e.insert(0, str(q))

        return self.name_e

    def validate(self):
        name = self.name_e.get().strip()
        if not name:
            messagebox.showwarning("Validation", "Name is required")
            return False
        return True

    def apply(self):
        self.result = {
            'type': self.type_cb.get().strip() or 'Generic',
            'name': self.name_e.get().strip(),
            'diameter': float(self.dia_e.get()) if self.dia_e.get().strip() else None,
            'length': int(self.len_e.get()) if self.len_e.get().strip() else None,
            'price': float(self.price_e.get()) if self.price_e.get().strip() else None,
            'qty': int(self.qty_e.get()) if self.qty_e.get().strip() else 0,
        }

class AddHospitalDialog(simpledialog.Dialog):
    def body(self, master):
        ttk.Label(master, text="Hospital name:").grid(row=0, column=0, sticky='e')
        self.e = ttk.Entry(master, width=40)
        self.e.grid(row=0, column=1)
        return self.e

    def validate(self):
        if not self.e.get().strip():
            messagebox.showwarning("Validation", "Name required")
            return False
        return True

    def apply(self):
        self.result = self.e.get().strip()

class ReportDialog(simpledialog.Dialog):
    def __init__(self, parent, db, hospitals):
        self.db = db
        self.hospitals = hospitals
        super().__init__(parent, title="Generate Report")

    def body(self, master):
        ttk.Label(master, text="Hospital (optional):").grid(row=0, column=0, sticky='e')
        hosp_names = ["All"] + [h[1] for h in self.hospitals]
        self.hosp_cb = ttk.Combobox(master, values=hosp_names, width=30)
        self.hosp_cb.set('All')
        self.hosp_cb.grid(row=0, column=1, sticky='w')

        ttk.Label(master, text="Date From (YYYY-MM-DD) optional:").grid(row=1, column=0, sticky='e')
        self.from_e = ttk.Entry(master, width=20)
        self.from_e.grid(row=1, column=1, sticky='w')

        ttk.Label(master, text="Date To (YYYY-MM-DD) optional:").grid(row=2, column=0, sticky='e')
        self.to_e = ttk.Entry(master, width=20)
        self.to_e.grid(row=2, column=1, sticky='w')

        return self.hosp_cb

    def apply(self):
        hosp_name = self.hosp_cb.get()
        hid = None
        if hosp_name != 'All':
            for h in self.hospitals:
                if h[1] == hosp_name:
                    hid = h[0]
                    break
        date_from = self.from_e.get().strip() or None
        date_to = self.to_e.get().strip() or None
        self.result = (hid, date_from, date_to)

class App(tk.Tk):
    def __init__(self, db: Database):
        super().__init__()
        self.db = db
        self.title("MedLink Inventory Manager")
        self.geometry("1300x780")
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        self._build_menu()
        self._create_widgets()
        self.refresh_all()

    def _build_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        inv_menu = tk.Menu(menubar, tearoff=0)
        inv_menu.add_command(label="Add Item", command=self.add_item)
        inv_menu.add_command(label="Edit Selected Item", command=self.edit_selected_item)
        inv_menu.add_command(label="Delete Selected Item", command=self.delete_selected_item)
        inv_menu.add_separator()
        inv_menu.add_command(label="Export Inventory CSV", command=self.export_csv)
        inv_menu.add_command(label="Export Inventory Excel", command=self.export_inventory_excel)
        menubar.add_cascade(label="Inventory", menu=inv_menu)

        hosp_menu = tk.Menu(menubar, tearoff=0)
        hosp_menu.add_command(label="Add Hospital", command=self.add_hospital)
        hosp_menu.add_command(label="Delete Selected Hospital", command=self.delete_selected_hospital)
        menubar.add_cascade(label="Hospitals", menu=hosp_menu)

        report_menu = tk.Menu(menubar, tearoff=0)
        report_menu.add_command(label="Generate Report", command=self.generate_report)
        report_menu.add_command(label="Export Report to Excel", command=self.export_report_excel)
        menubar.add_cascade(label="Reports", menu=report_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Refresh", command=self.refresh_all)
        menubar.add_cascade(label="View", menu=view_menu)

    def _create_widgets(self):
        top = ttk.Frame(self, padding=8)
        top.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(top, text="Mode:").pack(side=tk.LEFT)
        self.mode_cb = ttk.Combobox(top, values=["Simple","Advanced"], width=12)
        self.mode_cb.set("Advanced")
        self.mode_cb.pack(side=tk.LEFT, padx=(4,12))
        self.mode_cb.bind('<<ComboboxSelected>>', lambda e: self.on_mode_change())

        ttk.Label(top, text="Search Type:").pack(side=tk.LEFT)
        self.search_type_cb = ttk.Combobox(top, values=["All"]+DEFAULT_ITEM_TYPES, width=14)
        self.search_type_cb.set('All')
        self.search_type_cb.pack(side=tk.LEFT, padx=(4,6))

        ttk.Label(top, text="Search Name:").pack(side=tk.LEFT)
        self.search_name_e = ttk.Entry(top, width=30)
        self.search_name_e.pack(side=tk.LEFT, padx=(4,6))

        ttk.Button(top, text="Search", command=self.on_search).pack(side=tk.LEFT, padx=(6,8))
        ttk.Button(top, text="Clear", command=self.on_clear_search).pack(side=tk.LEFT)

        ttk.Label(top, text="Theme:").pack(side=tk.LEFT, padx=(12,0))
        self.theme_cb = ttk.Combobox(top, values=self.style.theme_names(), width=15)
        self.theme_cb.set(self.style.theme_use())
        self.theme_cb.pack(side=tk.LEFT, padx=(4,12))
        self.theme_cb.bind('<<ComboboxSelected>>', self.on_theme_change)

        ttk.Button(top, text="Refresh", command=self.refresh_all).pack(side=tk.LEFT)

        main = ttk.Frame(self, padding=8)
        main.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(main)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(left, text="Inventory", font=(None,11,'bold')).pack(anchor='w')
        cols = ("type","name","diameter","length","price","qty")
        self.tree = ttk.Treeview(left, columns=cols, show='headings', selectmode='browse')
        self.tree.heading('type', text='Type')
        self.tree.heading('name', text='Name')
        self.tree.heading('diameter', text='Diameter (mm)')
        self.tree.heading('length', text='Length (mm)')
        self.tree.heading('price', text='Price (USD)')
        self.tree.heading('qty', text='Qty')
        self.tree.column('type', width=100)
        self.tree.column('name', width=260)
        self.tree.column('diameter', width=100, anchor='center')
        self.tree.column('length', width=100, anchor='center')
        self.tree.column('price', width=100, anchor='center')
        self.tree.column('qty', width=80, anchor='center')
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<Double-1>', lambda e: self.edit_selected_item())

        # add scrollbar for inventory
        inv_scroll_y = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        inv_scroll_y.pack(side="right", fill="y")
        inv_scroll_x = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        inv_scroll_x.pack(side="bottom", fill="x")
        self.tree.configure(yscrollcommand=inv_scroll_y.set, xscrollcommand=inv_scroll_x.set)

        btns = ttk.Frame(left)
        btns.pack(fill=tk.X, pady=(8,0))
        ttk.Button(btns, text="Set Stock", command=self.set_stock).pack(side=tk.LEFT)
        ttk.Button(btns, text="Allocate to Hospital", command=self.allocate_to_hospital).pack(side=tk.LEFT, padx=(8,0))
        ttk.Button(btns, text="Add Item", command=self.add_item).pack(side=tk.LEFT, padx=(8,0))

        right = ttk.Frame(main, width=420)
        right.pack(side=tk.RIGHT, fill=tk.BOTH)

        ttk.Label(right, text="Hospitals", font=(None,11,'bold')).pack(anchor='w')
        self.hosp_tree = ttk.Treeview(right, columns=("name",), show='headings', height=8)
        self.hosp_tree.heading('name', text='Hospital')
        self.hosp_tree.pack(fill=tk.X)

        hosp_btns = ttk.Frame(right)
        hosp_btns.pack(fill=tk.X, pady=(6,10))
        ttk.Button(hosp_btns, text="Add Hospital", command=self.add_hospital).pack(side=tk.LEFT)
        ttk.Button(hosp_btns, text="Delete Hospital", command=self.delete_selected_hospital).pack(side=tk.LEFT, padx=(8,0))

        ttk.Label(right, text="Recent Allocations", font=(None,11,'bold')).pack(anchor='w')
        acol_cols = ("hospital","itype","iname","diameter","length","price","qty","ts")
        self.alloc_tree = ttk.Treeview(right, columns=acol_cols, show='headings', height=14)
        self.alloc_tree.heading('hospital', text='Hospital')
        self.alloc_tree.heading('itype', text='Type')
        self.alloc_tree.heading('iname', text='Item')
        self.alloc_tree.heading('diameter', text='D(mm)')
        self.alloc_tree.heading('length', text='L(mm)')
        self.alloc_tree.heading('price', text='Price(USD)')
        self.alloc_tree.heading('qty', text='Qty')
        self.alloc_tree.heading('ts', text='Timestamp')
        # Add scrollbars for allocations
        alloc_scroll_y = ttk.Scrollbar(right, orient="vertical", command=self.alloc_tree.yview)
        alloc_scroll_y.pack(side="right", fill="y")
        alloc_scroll_x = ttk.Scrollbar(right, orient="horizontal", command=self.alloc_tree.xview)
        alloc_scroll_x.pack(side="bottom", fill="x")
        self.alloc_tree.configure(yscrollcommand=alloc_scroll_y.set, xscrollcommand=alloc_scroll_x.set)

        self.alloc_tree.pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar()
        status = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w')
        status.pack(side=tk.BOTTOM, fill=tk.X)

    def on_mode_change(self):
        mode = self.mode_cb.get()
        if mode == 'Simple':
            # hide price column visually by setting width small
            self.tree.column('price', width=0)
            self.alloc_tree.column('price', width=0)
        else:
            self.tree.column('price', width=100)
            self.alloc_tree.column('price', width=90)

    def on_search(self):
        t = self.search_type_cb.get()
        if t == 'All':
            t = None
        name = self.search_name_e.get().strip() or None
        self.refresh_all(item_type_filter=t, name_filter=name)

    def on_clear_search(self):
        self.search_type_cb.set('All')
        self.search_name_e.delete(0, tk.END)
        self.refresh_all()

    def on_theme_change(self, _ev=None):
        try:
            theme = self.theme_cb.get()
            self.style.theme_use(theme)
            self.status_var.set(f"Theme: {theme}")
        except Exception as e:
            messagebox.showwarning("Theme", f"Unable to set theme: {e}")

    def refresh_all(self, item_type_filter=None, name_filter=None):
        try:
            items = self.db.fetch_items(item_type_filter=item_type_filter, name_filter=name_filter)
            hospitals = self.db.fetch_hospitals()
            allocations = self.db.fetch_recent_allocations(200)

            for i in self.tree.get_children():
                self.tree.delete(i)
            for it in items:
                iid, itype, name, d, l, p, q = it
                d_str = f"{d:.2f}" if d is not None else ""
                l_str = str(l) if l is not None else ""
                p_str = f"{p:.2f}" if p is not None else ""
                self.tree.insert('', 'end', iid=str(iid), values=(itype, name, d_str, l_str, p_str, int(q)))

            for i in self.hosp_tree.get_children():
                self.hosp_tree.delete(i)
            for h in hospitals:
                hid, name = h
                self.hosp_tree.insert('', 'end', iid=str(hid), values=(name,))

            for i in self.alloc_tree.get_children():
                self.alloc_tree.delete(i)
            for a in allocations:
                aid, hname, itype, iname, d, l, q, p, ts = a
                d_str = f"{d:.2f}" if d is not None else ""
                l_str = str(l) if l is not None else ""
                p_str = f"{p:.2f}" if p is not None else ""
                ts_str = ts.strftime('%Y-%m-%d %H:%M:%S') if hasattr(ts, 'strftime') else str(ts)
                self.alloc_tree.insert('', 'end', values=(hname, itype, iname, d_str, l_str, p_str, q, ts_str))

            self.status_var.set("Data refreshed")
            logger.debug("Refreshed UI")
        except Exception as e:
            logger.exception("Refresh failed")
            messagebox.showerror("Refresh", f"Unable to refresh data:\n{e}")

    # Hospital handlers
    def add_hospital(self):
        dlg = AddHospitalDialog(self, title="Add Hospital")
        if dlg.result:
            try:
                self.db.add_hospital(dlg.result)
                self.refresh_all()
                messagebox.showinfo("Success", "Hospital added")
            except Exception as e:
                logger.exception("Add hospital failed")
                messagebox.showerror("Error", f"Unable to add hospital:\n{e}")

    def delete_selected_hospital(self):
        sel = self.hosp_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select a hospital to delete")
            return
        hid = int(sel[0])
        name = self.hosp_tree.item(sel[0], 'values')[0]
        if messagebox.askyesno("Confirm", f"Delete hospital '{name}'? This will not delete allocations history."):
            try:
                self.db.delete_hospital(hid)
                self.refresh_all()
            except Exception as e:
                logger.exception("Delete hospital failed")
                messagebox.showerror("Error", f"Unable to delete hospital:\n{e}")

    # Item handlers
    def add_item(self):
        dlg = AddItemDialog(self, title="Add Item", item_types=DEFAULT_ITEM_TYPES)
        if dlg.result:
            try:
                r = dlg.result
                self.db.add_item(r['type'], r['name'], r['diameter'], r['length'], r['price'], r['qty'])
                self.refresh_all()
                messagebox.showinfo("Success", "Item added")
            except Exception as e:
                logger.exception("Add item failed")
                messagebox.showerror("Error", f"Unable to add item:\n{e}")

    def get_selected_item(self):
        sel = self.tree.selection()
        if not sel:
            return None
        iid = int(sel[0])
        vals = self.tree.item(sel[0], 'values')
        return iid, vals

    def edit_selected_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select an item to edit")
            return
        iid = int(sel[0])
        # fetch full item from DB to populate dialog
        items = self.db.fetch_items()
        item = next((it for it in items if it[0]==iid), None)
        if not item:
            messagebox.showerror("Error", "Item not found")
            return
        dlg = AddItemDialog(self, title="Edit Item", item_types=DEFAULT_ITEM_TYPES, item=item)
        if dlg.result:
            try:
                r = dlg.result
                self.db.update_item(iid, r['name'], r['diameter'], r['length'], r['price'])
                if r['qty'] is not None:
                    self.db.set_stock(iid, r['qty'])
                self.refresh_all()
                messagebox.showinfo("Success", "Item updated")
            except Exception as e:
                logger.exception("Edit item failed")
                messagebox.showerror("Error", f"Unable to update item:\n{e}")

    def delete_selected_item(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select an item to delete")
            return
        iid = int(sel[0])
        name = self.tree.item(sel[0], 'values')[1]
        if messagebox.askyesno("Confirm", f"Delete item '{name}'? This will not delete allocations history."):
            try:
                self.db.delete_item(iid)
                self.refresh_all()
            except Exception as e:
                logger.exception("Delete item failed")
                messagebox.showerror("Error", f"Unable to delete item:\n{e}")

    def set_stock(self):
        sel = self.get_selected_item()
        if not sel:
            messagebox.showinfo("Select", "Select an item to set stock")
            return
        iid, vals = sel
        cur_qty = int(vals[5])
        qty = simpledialog.askinteger("Set Stock", f"Current stock = {cur_qty}\nEnter new stock quantity:", parent=self, minvalue=0)
        if qty is None:
            return
        try:
            self.db.set_stock(iid, qty)
            self.refresh_all()
            messagebox.showinfo("Stock Updated", f"Stock updated to {qty}")
        except Exception as e:
            logger.exception("Set stock failed")
            messagebox.showerror("Error", f"Unable to set stock:\n{e}")

    def allocate_to_hospital(self):
        sit = self.get_selected_item()
        if not sit:
            messagebox.showinfo("Select", "Select an item to allocate")
            return
        sh = self.hosp_tree.selection()
        if not sh:
            messagebox.showinfo("Select", "Select a hospital to allocate to")
            return
        iid, ival = sit
        hid = int(sh[0])
        hname = self.hosp_tree.item(sh[0], 'values')[0]
        available = int(ival[5])
        if available <= 0:
            messagebox.showwarning("No Stock", "Selected item has zero stock.")
            return
        # ask whether to use custom price at allocation
        cur_price = ival[4]
        price_prompt = f"Current price: {cur_price if cur_price else 'N/A'}\nEnter price to record for this allocation (leave empty to use current price):"
        price_str = simpledialog.askstring("Price at allocation", price_prompt, parent=self)
        price_val = float(price_str) if price_str and price_str.strip() else None
        qty = simpledialog.askinteger("Allocate", f"Available: {available}\nEnter quantity to allocate to {hname}:", parent=self, minvalue=1, maxvalue=available)
        if qty is None:
            return
        try:
            self.db.allocate(iid, hid, qty, price_at_alloc=price_val)
            self.refresh_all()
            messagebox.showinfo("Allocated", f"Allocated {qty} units to {hname}")
        except Exception as e:
            logger.exception("Allocation failed")
            messagebox.showerror("Error", f"Allocation failed:\n{e}")

    def export_csv(self):
        try:
            import csv
            rows = self.db.fetch_items()
            fn = f"inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            with open(fn, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(['type','name','diameter_mm','length_mm','price_usd','qty'])
                for r in rows:
                    _, itype, name, d, l, p, q = r
                    w.writerow([itype, name, f"{d:.2f}" if d is not None else '', l or '', f"{p:.2f}" if p is not None else '', q])
            messagebox.showinfo("Exported", f"Inventory exported to {fn}")
            logger.info(f"Exported inventory to {fn}")
        except Exception as e:
            logger.exception("Export failed")
            messagebox.showerror("Export", f"Unable to export CSV:\n{e}")

    def export_inventory_excel(self):
        if Workbook is None:
            messagebox.showerror('Missing library', 'openpyxl is required to export Excel files. Install via pip install openpyxl')
            return
        try:
            rows = self.db.fetch_items()
            wb = Workbook()
            ws = wb.active
            ws.title = 'Inventory'
            ws.append(['Type','Name','Diameter(mm)','Length(mm)','Price(USD)','Qty'])
            for r in rows:
                _, itype, name, d, l, p, q = r
                ws.append([itype, name, f"{d:.2f}" if d is not None else '', l or '', f"{p:.2f}" if p is not None else '', q])
            fn = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files','*.xlsx')], title='Save Inventory as Excel')
            if fn:
                wb.save(fn)
                messagebox.showinfo('Saved', f'Inventory saved to {fn}')
                logger.info(f'Inventory exported to Excel: {fn}')
        except Exception as e:
            logger.exception('Export Excel failed')
            messagebox.showerror('Error', f'Unable to export Excel:\n{e}')

    # Reports
    def generate_report(self):
        hospitals = self.db.fetch_hospitals()
        dlg = ReportDialog(self, self.db, hospitals)
        if dlg.result is None:
            return
        hid, date_from, date_to = dlg.result
        # parse dates
        df = None
        dtf = None
        try:
            if date_from:
                dtf = datetime.fromisoformat(date_from)
            if date_to:
                dtt = datetime.fromisoformat(date_to)
            else:
                dtt = None
        except Exception:
            messagebox.showwarning('Date', 'Date format must be YYYY-MM-DD')
            return
        rows = self.db.fetch_allocations_for_report(hospital_id=hid, date_from=dtf, date_to=dtt)
        # build Persian report text
        lines = []
        header = 'گزارش تخصیص ها\n'
        header += f"تاریخ گزارش: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        lines.append(header)
        total_count = 0
        total_value = 0.0
        for r in rows:
            aid, hname, itype, iname, d, l, qty, price, ts = r
            d_str = f"{d:.2f}" if d is not None else ''
            l_str = str(l) if l is not None else ''
            price_str = f"{price:.2f}" if price is not None else 'N/A'
            ts_str = ts.strftime('%Y-%m-%d %H:%M') if hasattr(ts, 'strftime') else str(ts)
            line = f"{qty} عدد از '{iname}' ({itype}) سایز: {d_str}×{l_str} — قیمت: {price_str} دلار — تاریخ: {ts_str} — بیمارستان: {hname}"
            lines.append(line)
            total_count += qty
            if price is not None:
                total_value += qty * float(price)
        summary = f"\nجمع کل اقلام تخصیص داده شده: {total_count} عدد\nارزش تقریبی به دلار: {total_value:.2f}\n"
        lines.append(summary)
        report_text = "\n".join(lines)
        # show in a new window
        win = tk.Toplevel(self)
        win.title('گزارش تخصیص‌ها')
        txt = tk.Text(win, wrap='word')
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert('1.0', report_text)
        txt.config(state='disabled')
        # store last report rows for export
        self._last_report_rows = rows

    def export_report_excel(self):
        if Workbook is None:
            messagebox.showerror('Missing library', 'openpyxl is required to export Excel files. Install via pip install openpyxl')
            return
        rows = getattr(self, '_last_report_rows', None)
        if not rows:
            messagebox.showinfo('No report', 'Generate a report first (Reports -> Generate Report)')
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = 'Allocations Report'
            ws.append(['Hospital','Type','Item','Diameter(mm)','Length(mm)','Price(USD)','Qty','Timestamp'])
            for r in rows:
                _, hname, itype, iname, d, l, qty, price, ts = r
                ws.append([hname, itype, iname, f"{d:.2f}" if d is not None else '', l or '', f"{price:.2f}" if price is not None else '', qty, ts.strftime('%Y-%m-%d %H:%M:%S') if hasattr(ts,'strftime') else str(ts)])
            fn = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files','*.xlsx')], title='Save Report as Excel')
            if fn:
                wb.save(fn)
                messagebox.showinfo('Saved', f'Report saved to {fn}')
                logger.info(f'Report exported to Excel: {fn}')
        except Exception as e:
            logger.exception('Export report failed')
            messagebox.showerror('Error', f'Unable to export report:\n{e}')

def main():
    conn_str = None
    if len(sys.argv) > 1:
        conn_str = sys.argv[1]
    db = Database(conn_str)
    app = App(db)
    try:
        app.mainloop()
    finally:
        db.close()

if __name__ == '__main__':
    main()
