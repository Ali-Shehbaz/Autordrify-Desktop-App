import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
import pdfplumber
import re
import shutil
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from queue import Queue
import time
from PIL import Image
import pystray
import winreg as reg

# --- 1. CONFIGURATION ---
def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

WATCH_FOLDER = r"F:/"
CUSTOMER_LIST_FILE = resource_path("customers.txt")
ICON_FILE = resource_path("icon.png")

DESTINATIONS = {
    "SO": r"G:/.FA/Sale Orders (SO)",
    "DC": r"G:/.FA/Delivery Challans (DC)",
    "Invoice": r"G:/INVOICE",
    "Ledger": r"G:/Previous Ledgers/01-ALL NEW DOWNLOADED LEDGERS"
}

# --- 2. REGEX PATTERNS ---
so_pattern = re.compile(r"Sales Order No\.\s*(\d+)")
so_customer_pattern = re.compile(r"Sales Order\n(.*?)\s+Sales Order No\.")
so_date_pattern = re.compile(r"Date\s+(\d{2}/\d{2}/\d{4})")
dc_pattern = re.compile(r"DC No\.\s*(\d+)")
dc_customer_pattern = re.compile(r"Delivery Challan\n(.*?)\s+DC No\.")
dc_date_pattern = re.compile(r"Date\s+(\d{2}/\d{2}/\d{4})")
time_pattern = re.compile(r"_(\d{14})") 
inv_pattern = re.compile(r"Inv No\.\s*([A-Z0-9]+)") 
inv_customer_pattern = re.compile(r"Sales Tax Invoice\n(.*?)\s+Inv No\.")
inv_date_pattern = re.compile(r"Date\s+(\d{2}/\d{2}/\d{4})")
inv_dc_pattern = re.compile(r"DC No\.[ \t]*([^\n]*)")
inv_po_pattern = re.compile(r"PO No\.[ \t]*([^\n]*)")
stmt_customer_pattern = re.compile(r"Combined Account Statement \(Invoice Detail\)\n(.*?)\s+Account No")
stmt_date_range_pattern = re.compile(r"Date From:\s*(\d{2}/\d{2}/\d{4})\s*to:\s*(\d{2}/\d{2}/\d{4})")

file_queue = Queue()

# --- 3. CUSTOMER MANAGER ---
class CustomerManager(tk.Toplevel):
    def __init__(self, parent, customers, save_callback):
        super().__init__(parent)
        self.title("Manage Customers")
        self.geometry("400x550")
        self.customers = customers
        self.save_callback = save_callback
        self.setup_ui()
    def setup_ui(self):
        tk.Label(self, text="Customer List", font=("Arial", 11, "bold")).pack(pady=5)
        self.listbox = tk.Listbox(self, selectmode=tk.EXTENDED, font=("Arial", 10))
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        for c in self.customers: self.listbox.insert(tk.END, c)
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Add New", width=12, command=self.add_customer).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Delete Selected", width=12, bg="#e74c3c", fg="white", command=self.remove_customers).pack(side=tk.LEFT, padx=5)
    def add_customer(self):
        new_name = simpledialog.askstring("Add Customer", "Enter name:")
        if new_name and new_name.strip():
            name = new_name.strip()
            if name not in self.customers:
                self.customers.append(name)
                self.customers.sort(); self.refresh_list(); self.save_callback(self.customers)
    def remove_customers(self):
        selections = self.listbox.curselection()
        if selections and messagebox.askyesno("Confirm", f"Delete {len(selections)} names?"):
            for index in reversed(selections):
                name = self.listbox.get(index)
                self.customers.remove(name)
            self.refresh_list(); self.save_callback(self.customers)
    def refresh_list(self):
        self.listbox.delete(0, tk.END)
        for c in sorted(self.customers): self.listbox.insert(tk.END, c)

# --- 4. MAIN APP ---
class AutordrifyApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Autordrify Legacy")
        self.root.geometry("1150x650")
        self.root.protocol('WM_DELETE_WINDOW', self.hide_window)
        
        self.customers = self.load_customers()
        self.setup_ui()
        self.register_system_features() # Auto-register on launch
        
        threading.Thread(target=self.start_monitoring, daemon=True).start()
        threading.Thread(target=self.setup_tray, daemon=True).start()
        self.process_queue()

    def register_system_features(self):
        if not getattr(sys, 'frozen', False): return
        app_path = sys.executable
        try:
            with reg.OpenKey(reg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, reg.KEY_SET_VALUE) as key:
                reg.SetValueEx(key, "Autordrify", 0, reg.REG_SZ, f'"{app_path}" --minimized')
        except: pass
        try:
            key_path = r"Software\Classes\SystemFileAssociations\.pdf\shell\Autordrify\command"
            with reg.CreateKey(reg.HKEY_CURRENT_USER, key_path) as key:
                reg.SetValue(key, "", reg.REG_SZ, f'"{app_path}" "%1"')
            label_path = r"Software\Classes\SystemFileAssociations\.pdf\shell\Autordrify"
            with reg.CreateKey(reg.HKEY_CURRENT_USER, label_path) as key:
                reg.SetValueEx(key, "MUIVerb", 0, reg.REG_SZ, "Rename with Autordrify")
                reg.SetValueEx(key, "Icon", 0, reg.REG_SZ, app_path)
        except: pass

    def hide_window(self):
        self.root.withdraw()

    def silent_taskbar_alert(self):
        self.root.deiconify()
        self.root.lower()

    def show_window(self):
        self.root.after(0, self.root.deiconify)
        self.root.after(0, self.root.lift)

    def quit_app(self, icon, item):
        icon.stop()
        self.root.after(0, self.root.destroy)
        os._exit(0)

    def setup_tray(self):
        image = Image.open(ICON_FILE)
        menu = pystray.Menu(
            pystray.MenuItem("Show Autordrify", self.show_window),
            pystray.MenuItem("Manual Scan F:/", self.manual_scan),
            pystray.MenuItem("Exit", self.quit_app)
        )
        self.icon = pystray.Icon("Autordrify", image, "Autordrify Active", menu)
        self.icon.run()

    def load_customers(self):
        if not os.path.exists(CUSTOMER_LIST_FILE): return []
        with open(CUSTOMER_LIST_FILE, "r") as f:
            return sorted([line.strip() for line in f if line.strip()])

    def save_customers(self, customer_list):
        with open(CUSTOMER_LIST_FILE, "w") as f:
            for c in sorted(customer_list): f.write(c + "\n")
        self.customers = customer_list

    def setup_ui(self):
        header = tk.Label(self.root, text="Autordrify Background Monitor", font=("Arial", 14, "bold"), pady=10)
        header.pack()
        self.tree = ttk.Treeview(self.root, columns=("Status", "Name", "Type", "Date", "Path"), show='headings', selectmode="extended")
        self.tree.heading("Status", text="Status"); self.tree.column("Status", width=80)
        self.tree.heading("Name", text="New Filename"); self.tree.column("Name", width=420)
        self.tree.heading("Type", text="Type"); self.tree.column("Type", width=80)
        self.tree.heading("Date", text="Date"); self.tree.column("Date", width=100)
        self.tree.heading("Path", text="Location"); self.tree.column("Path", width=250)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Manual Scan F:/", bg="#3498db", fg="white", width=15, command=self.manual_scan).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Rename Selected", bg="#2ecc71", fg="white", width=15, command=self.handle_rename).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Move Selected", bg="#f39c12", fg="white", width=15, command=self.handle_move).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Manage Customers", bg="#7f8c8d", fg="white", width=18, command=self.open_customer_manager).pack(side=tk.LEFT, padx=5)

    def open_customer_manager(self):
        CustomerManager(self.root, self.customers, self.save_customers)

    def parse_pdf(self, file_path):
        filename = os.path.basename(file_path)
        try:
            with pdfplumber.open(file_path) as pdf:
                full_text = ""
                for page in pdf.pages: full_text += page.extract_text() or "" + "\n"
            if "GDNSO_" in filename:
                dc_match, cust_match, date_match, ts_match = dc_pattern.search(full_text), dc_customer_pattern.search(full_text), dc_date_pattern.search(full_text), time_pattern.search(filename)
                cust = cust_match.group(1).strip() if cust_match else (self.find_customer_fallback(full_text))
                date_raw = date_match.group(1).replace("/", "-") if date_match else "00-00-0000"
                t_str = f"{ts_match.group(1)[8:10]}-{ts_match.group(1)[10:12]}-{ts_match.group(1)[12:14]}" if ts_match else "00-00-00"
                return f"DC-{dc_match.group(1) if dc_match else 'UNK'}, {cust}, {date_raw}, {t_str}.pdf", "DC", date_raw
            elif "SO_" in filename:
                so_match, cust_match, date_match = so_pattern.search(full_text), so_customer_pattern.search(full_text), so_date_pattern.search(full_text)
                cust = cust_match.group(1).strip() if cust_match else (self.find_customer_fallback(full_text))
                date_raw = date_match.group(1).replace("/", "-") if date_match else "00-00-0000"
                return f"{so_match.group(1) if so_match else 'UNK'}, {cust}, {date_raw}.pdf", "SO", date_raw
            elif "SI_" in filename:
                inv_match, cust_match, date_match, dc_match, po_match = inv_pattern.search(full_text), inv_customer_pattern.search(full_text), inv_date_pattern.search(full_text), inv_dc_pattern.search(full_text), inv_po_pattern.search(full_text)
                cust = cust_match.group(1).strip() if cust_match else (self.find_customer_fallback(full_text))
                date_raw = date_match.group(1).replace("/", "-") if date_match else "00-00-0000"
                return f"{inv_match.group(1).strip() if inv_match else 'UNK'}, {cust}, D.C-{dc_match.group(1).strip() if dc_match else ''}, PO-{po_match.group(1).strip() if po_match else ''}, {date_raw}.pdf", "Invoice", date_raw
            elif "statement" in filename.lower():
                cust_match, date_match = stmt_customer_pattern.search(full_text), stmt_date_range_pattern.search(full_text)
                cust = cust_match.group(1).strip() if cust_match else (self.find_customer_fallback(full_text))
                d_range = f"From {date_match.group(1).replace('/', '-')} to {date_match.group(2).replace('/', '-')}" if date_match else "all dates"
                return f"Ledger, {cust}, {d_range}.pdf", "Ledger", (date_match.group(2).replace('/', '-') if date_match else "01-01-2000")
            return None, "Other", None
        except: return None, "Error", None

    def find_customer_fallback(self, text):
        for cust in self.customers:
            if cust.lower() in text.lower(): return cust
        return "UNKNOWN_CUSTOMER"

    def start_monitoring(self):
        class Handler(FileSystemEventHandler):
            def process(self, event):
                if event.is_directory: return
                path = event.src_path if not hasattr(event, 'dest_path') else event.dest_path
                if path.lower().endswith(".pdf"):
                    time.sleep(1); file_queue.put(path)
            def on_created(self, event): self.process(event)
            def on_moved(self, event): self.process(event)
        obs = Observer(); obs.schedule(Handler(), WATCH_FOLDER, recursive=False); obs.start()
        try:
            while True: time.sleep(1)
        except: obs.stop(); obs.join()
        
    def process_queue(self):
        try:
            while not file_queue.empty():
                path = file_queue.get_nowait()
                new_name, ftype, fdate = self.parse_pdf(path)
                if new_name and ftype not in ["Other", "Error"]:
                    self.tree.insert("", "end", values=("Pending", new_name, ftype, fdate, path))
                    self.silent_taskbar_alert()
        except: pass
        finally: self.root.after(1000, self.process_queue)

    def manual_scan(self):
        for file in os.listdir(WATCH_FOLDER):
            if file.lower().endswith(".pdf"): file_queue.put(os.path.join(WATCH_FOLDER, file))

    def handle_rename(self):
        selected = self.tree.selection()
        if not selected: return
        for item_id in selected:
            v = self.tree.item(item_id)['values']
            if v[0] != "Pending": continue
            new_path = os.path.join(os.path.dirname(v[4]), v[1])
            try:
                os.rename(v[4], new_path)
                self.tree.item(item_id, values=("Renamed", v[1], v[2], v[3], new_path))
            except Exception as e: messagebox.showerror("Error", str(e))

    def handle_move(self):
        selected = self.tree.selection()
        if not selected: return
        for item_id in selected:
            v = self.tree.item(item_id)['values']
            if v[0] != "Renamed": continue
            dest = DESTINATIONS.get(v[2], "G:/Unsorted")
            try:
                dt = datetime.strptime(str(v[3]), "%d-%m-%Y")
                sub = os.path.join(dest, f"{dt.strftime('%B')}-{dt.strftime('%Y')}") if v[2] in ["SO", "DC"] else (os.path.join(dest, f"{dt.strftime('%b').upper()}-{dt.strftime('%Y')}") if v[2] == "Invoice" else dest)
                if not os.path.exists(sub): os.makedirs(sub)
                shutil.move(v[4], os.path.join(sub, v[1]))
                self.tree.item(item_id, values=("Moved", v[1], v[2], v[3], os.path.join(sub, v[1])))
            except Exception as e: messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = AutordrifyApp(root)
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".pdf"):
        file_queue.put(sys.argv[1])
        app.show_window()
    else:
        if "--minimized" in sys.argv:
            app.hide_window()
    if os.path.exists(ICON_FILE):
        try: root.iconphoto(False, tk.PhotoImage(file=ICON_FILE))
        except: pass
    root.mainloop()