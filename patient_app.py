import csv
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, datetime

# ================= SETTINGS =================
DATA_DIR = "PMS_Data"          # all Excel files are stored in this folder
USERNAME = "admin"
PASSWORD = "admin123"
BG = "#eaf4f4"
GREEN = "#0b6e4f"

os.makedirs(DATA_DIR, exist_ok=True)


# ================= EXCEL FILE STORAGE =================
class Store:
    """One Excel (.csv) file = one module. Handles add/update/delete/search."""

    def __init__(self, filename, fields):
        self.path = os.path.join(DATA_DIR, filename)
        self.counter_path = self.path.replace(".csv", ".counter")
        self.header = ["ID"] + fields
        if not os.path.exists(self.path):
            self._write([])

    def read(self):
        with open(self.path, newline="", encoding="utf-8-sig") as f:
            rows = [r for r in csv.reader(f)][1:]
        n = len(self.header)
        rows = [(r + [""] * n)[:n] for r in rows if r]
        rows.sort(key=lambda r: int(r[0]))
        return rows

    def _write(self, rows, quiet=False):
        rows.sort(key=lambda r: int(r[0]))
        try:
            with open(self.path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(self.header)
                w.writerows(rows)
            return True
        except PermissionError:
            if not quiet:
                messagebox.showerror("Excel file is open",
                                     os.path.basename(self.path) +
                                     " is open in Excel.\nClose it and try again.")
            return False

    def next_id(self, rows):
        # IDs only go up; a deleted ID is never used again
        last = max([int(r[0]) for r in rows], default=0)
        try:
            with open(self.counter_path) as f:
                last = max(last, int(f.read().strip()))
        except (FileNotFoundError, ValueError):
            pass
        return last + 1

    def add(self, values):
        rows = self.read()
        new_id = self.next_id(rows)
        rows.append([str(new_id)] + values)
        if self._write(rows):
            self.last_id = new_id
            with open(self.counter_path, "w") as f:
                f.write(str(new_id))
            return True
        return False

    def update(self, rid, values):
        rows = self.read()
        for r in rows:
            if r[0] == rid:
                r[1:] = values
        return self._write(rows)

    def delete(self, rid):
        return self._write([r for r in self.read() if r[0] != rid])


patients = Store("patients.csv",
                 ["Name", "Age", "Gender", "Phone", "Address", "Blood Group", "Disease",
                  "Doctor", "Tests"])
doctors = Store("doctors.csv", ["Name", "Specialization", "Phone", "Fee"])
tests = Store("tests.csv", ["Test Name", "Price"])
appointments = Store("appointments.csv",
                     ["Patient", "Doctor", "Date", "Time", "Reason", "Status"])
billing = Store("billing.csv",
                ["Patient", "Description", "Amount", "Paid", "Date", "Status"])


def patient_choices():
    return [r[0] + " - " + r[1] for r in patients.read()]


def doctor_choices():
    return [r[0] + " - Dr. " + r[1] + " (" + r[2] + ")" for r in doctors.read()]


def test_choices():
    return [r[1] for r in tests.read()]


def appt_status(date_str, time_str):
    """Automatic status: future = Scheduled, past = Completed."""
    try:
        when = datetime.strptime(date_str + " " + time_str, "%Y-%m-%d %H:%M")
    except ValueError:
        return "Scheduled"
    return "Scheduled" if when >= datetime.now() else "Completed"


def sync_appointments():
    """Refreshes the Status of every appointment (Cancelled stays Cancelled).
    Returns True if anything changed."""
    rows = appointments.read()
    changed = False
    for r in rows:
        if r[6] == "Cancelled":
            continue
        new = appt_status(r[3], r[4])
        if r[6] != new:
            r[6] = new
            changed = True
    if changed:
        appointments._write(rows, quiet=True)
    return changed


def bill_status(amount, paid):
    """Automatic status: nothing paid = Pending, part paid = Partial, all paid = Paid."""
    if paid <= 0:
        return "Pending"
    if paid >= amount:
        return "Paid"
    return "Partial"


def finalize_bill(v):
    """v = [Patient, Description, Amount, Paid, Date]. Adds the automatic Status."""
    amount = float(v[2])
    paid = float(v[3] or 0)
    if paid > amount:
        messagebox.showerror("Invalid input", "Amount Paid cannot be more than Amount.")
        return None
    if not v[3]:
        v[3] = "0"
    return v + [bill_status(amount, paid)]


def money(x):
    return ("%.2f" % x).rstrip("0").rstrip(".")


class MultiVar:
    """Lets a multi-select list box behave like a normal text variable."""

    def __init__(self, listbox):
        self.lb = listbox

    def get(self):
        return "; ".join(self.lb.get(i) for i in self.lb.curselection())

    def set(self, value):
        self.lb.selection_clear(0, "end")
        names = [x.strip() for x in value.split(";") if x.strip()]
        for i in range(self.lb.size()):
            if self.lb.get(i) in names:
                self.lb.selection_set(i)


def make_bill(doctor_text, tests_text):
    """Bill = doctor's consultation fee + price of every test performed."""
    items, total = [], 0.0
    doctor_id = doctor_text.split(" - ")[0]
    for d in doctors.read():
        if d[0] == doctor_id:
            fee = float(d[4] or 0)
            items.append("Dr. %s (%s)" % (d[1], money(fee)))
            total += fee
    prices = {t[1]: float(t[2] or 0) for t in tests.read()}
    for name in [x.strip() for x in tests_text.split(";") if x.strip()]:
        price = prices.get(name, 0.0)
        items.append("%s (%s)" % (name, money(price)))
        total += price
    if not items:
        items.append("No charges")
    return " + ".join(items), total


def bill_prefix(patient_id):
    return "Pt #%s - " % patient_id


def auto_bill_add(patient_id, v):
    """After a patient is added: generate the bill.
    v = [Name, Age, Gender, Phone, Address, Blood, Disease, Doctor, Tests]"""
    desc, total = make_bill(v[7], v[8])
    bill = finalize_bill([str(patient_id) + " - " + v[0],
                          bill_prefix(patient_id) + desc, money(total), "0",
                          str(date.today())])
    if bill and billing.add(bill):
        return "Bill generated automatically: Rs %s\n%s" % (money(total), desc)
    return None


def auto_bill_update(patient_id, v):
    """After a patient is edited: refresh the unpaid bill from the new doctor/tests."""
    desc, total = make_bill(v[7], v[8])
    prefix = bill_prefix(patient_id)
    mine = [b for b in billing.read() if b[2].startswith(prefix)]
    if not mine:
        return auto_bill_add(patient_id, v)
    for b in mine:
        if float(b[4] or 0) <= 0:
            bill = finalize_bill([str(patient_id) + " - " + v[0], prefix + desc,
                                  money(total), "0", b[5]])
            if bill and billing.update(b[0], bill):
                return "Bill updated automatically: Rs %s\n%s" % (money(total), desc)
            return None
    return "Existing bill already has payments, so it was not changed."


def remove_unpaid_bill(patient_id):
    """When a patient is deleted, the unpaid bill is removed (paid bills are kept)."""
    prefix = bill_prefix(patient_id)
    rows = billing.read()
    keep = [b for b in rows if not (b[2].startswith(prefix) and float(b[4] or 0) <= 0)]
    if len(keep) != len(rows):
        billing._write(keep, quiet=True)


# ================= GENERIC MODULE SCREEN =================
class CrudTab(tk.Frame):
    """A full screen (form + buttons + search + table) for one module."""

    def __init__(self, parent, store, fields, on_change=None, finalize=None,
                 extra_buttons=None, note=None, after_add=None, after_update=None, after_delete=None):
        super().__init__(parent, bg=BG)
        self.store = store
        self.fields = fields
        self.on_change = on_change
        self.finalize = finalize
        self.after_add = after_add
        self.after_update = after_update
        self.after_delete = after_delete
        self.vars = []
        self.combos = []
        self.multis = []
        self.selected = None
        self.search_var = tk.StringVar()

        form = tk.Frame(self, bg=BG)
        form.pack(pady=10)
        for i, f in enumerate(fields):
            r, c = divmod(i, 2)
            text = f["label"] + (" *" if f.get("req") else "")
            tk.Label(form, text=text, bg=BG, font=("Arial", 10)).grid(
                row=r, column=c * 2, padx=8, pady=6, sticky="e")
            if f["kind"] == "multi":
                w = tk.Listbox(form, selectmode="multiple", height=5, width=30,
                               exportselection=False, font=("Arial", 10))
                var = MultiVar(w)
                self.multis.append((f, var, w))
            else:
                var = tk.StringVar()
            if f["kind"] == "multi":
                pass
            elif f["kind"] in ("combo", "ref"):
                w = ttk.Combobox(form, textvariable=var, width=28, state="readonly")
                if f["kind"] == "combo":
                    w["values"] = f["options"]
                self.combos.append((f, w))
            else:
                w = tk.Entry(form, textvariable=var, width=30, font=("Arial", 10))
            w.grid(row=r, column=c * 2 + 1, padx=8, pady=6, sticky="w")
            self.vars.append(var)

        if note:
            tk.Label(self, text=note, bg=BG, fg="gray",
                     font=("Arial", 9, "italic")).pack()

        btns = tk.Frame(self, bg=BG)
        btns.pack(pady=6)
        btn_list = [("Add", self.add), ("Update", self.update),
                    ("Delete", self.delete), ("Clear", self.clear),
                    ("Open in Excel", self.open_excel)]
        for text, cmd in (extra_buttons or []):
            btn_list.append((text, lambda c=cmd: c(self)))
        for text, cmd in btn_list:
            tk.Button(btns, text=text, width=max(13, len(text) + 2), command=cmd,
                      bg=GREEN, fg="white",
                      font=("Arial", 10, "bold")).pack(side="left", padx=5)

        sf = tk.Frame(self, bg=BG)
        sf.pack(pady=4)
        tk.Label(sf, text="Search:", bg=BG, font=("Arial", 10)).pack(side="left")
        tk.Entry(sf, textvariable=self.search_var, width=26).pack(side="left", padx=5)
        tk.Button(sf, text="Search", command=self.search, bg=GREEN, fg="white").pack(side="left", padx=3)
        tk.Button(sf, text="Show All", command=self.show_all, bg=GREEN, fg="white").pack(side="left", padx=3)

        cols = store.header
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100, anchor="center")
        self.tree.column("ID", width=40)
        self.status_idx = cols.index("Status") if "Status" in cols else None
        for tag, color in [("Scheduled", "#1565c0"), ("Completed", "#2e7d32"),
                           ("Cancelled", "#c62828"), ("Pending", "#c62828"),
                           ("Paid", "#2e7d32"), ("Partial", "#ef6c00")]:
            self.tree.tag_configure(tag, foreground=color)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", pady=10)
        self.tree.pack(fill="both", expand=True, padx=(10, 0), pady=10)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        self.refresh_refs()
        self.clear()
        self.load()

    # ----- helpers -----
    def refresh_refs(self):
        for f, w in self.combos:
            if f["kind"] == "ref":
                w["values"] = f["source"]()
        for f, var, w in self.multis:
            keep = var.get()
            w.delete(0, "end")
            for name in f["source"]():
                w.insert("end", name)
            var.set(keep)

    def load(self, rows=None):
        if rows is None:
            rows = self.store.read()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in rows:
            tags = (r[self.status_idx],) if self.status_idx is not None else ()
            self.tree.insert("", "end", values=r, tags=tags)

    def clear(self):
        self.selected = None
        for f, var in zip(self.fields, self.vars):
            if f["kind"] == "date":
                var.set(str(date.today()))
            elif f["kind"] == "combo":
                var.set(f["options"][0])
            else:
                var.set(f.get("default", ""))
        self.tree.selection_remove(self.tree.selection())

    def validate(self):
        vals = []
        for f, var in zip(self.fields, self.vars):
            v = var.get().strip()
            label = f["label"]
            if f.get("req") and not v:
                messagebox.showwarning("Missing data", label + " is required.")
                return None
            if v:
                k = f["kind"]
                if k == "int":
                    if not v.isdigit() or not (0 < int(v) <= f.get("max", 10 ** 9)):
                        messagebox.showerror("Invalid input", "Enter a valid " + label + ".")
                        return None
                elif k == "number":
                    try:
                        float(v)
                    except ValueError:
                        messagebox.showerror("Invalid input", label + " must be a number.")
                        return None
                elif k == "phone":
                    if not (v.isdigit() and len(v) == 10):
                        messagebox.showerror("Invalid input", "Phone number must be 10 digits.")
                        return None
                elif k == "date":
                    try:
                        datetime.strptime(v, "%Y-%m-%d")
                    except ValueError:
                        messagebox.showerror("Invalid input", label + " must be YYYY-MM-DD.")
                        return None
                elif k == "time":
                    try:
                        datetime.strptime(v, "%H:%M")
                    except ValueError:
                        messagebox.showerror("Invalid input", label + " must be HH:MM (24-hour).")
                        return None
            vals.append(v)
        return vals

    def changed(self):
        self.clear()
        self.load()
        if self.on_change:
            self.on_change()

    # ----- actions -----
    def add(self):
        vals = self.validate()
        if vals is not None and self.finalize:
            vals = self.finalize(vals)
        if vals is not None and self.store.add(vals):
            msg = "Record added and saved in Excel file."
            if self.after_add:
                extra = self.after_add(self.store.last_id, vals)
                if extra:
                    msg += "\n\n" + extra
            messagebox.showinfo("Success", msg)
            self.changed()

    def on_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        rid = str(self.tree.item(sel[0])["values"][0])
        for r in self.store.read():
            if r[0] == rid:
                self.selected = rid
                for var, val in zip(self.vars, r[1:]):
                    var.set(val)
                break

    def update(self):
        if self.selected is None:
            messagebox.showwarning("Select record", "Click a row in the table first.")
            return
        vals = self.validate()
        if vals is not None and self.finalize:
            vals = self.finalize(vals)
        if vals is not None and self.store.update(self.selected, vals):
            msg = "Record updated in Excel file."
            if self.after_update:
                extra = self.after_update(self.selected, vals)
                if extra:
                    msg += "\n\n" + extra
            messagebox.showinfo("Success", msg)
            self.changed()

    def delete(self):
        if self.selected is None:
            messagebox.showwarning("Select record", "Click a row in the table first.")
            return
        if messagebox.askyesno("Confirm", "Delete this record?"):
            rid = self.selected
            if self.store.delete(rid):
                if self.after_delete:
                    self.after_delete(rid)
                self.changed()

    def search(self):
        term = self.search_var.get().strip().lower()
        self.load([r for r in self.store.read() if term in " ".join(r[1:]).lower()])

    def show_all(self):
        self.search_var.set("")
        self.load()

    def open_excel(self):
        os.startfile(self.store.path)


# ================= DASHBOARD =================
class Dashboard(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        tk.Label(self, text="Hospital Overview", font=("Arial", 18, "bold"),
                 bg=BG, fg=GREEN).pack(pady=15)
        grid = tk.Frame(self, bg=BG)
        grid.pack(pady=10)
        self.values = {}
        cards = [("patients", "Total Patients", "#0b6e4f"),
                 ("doctors", "Total Doctors", "#1565c0"),
                 ("appts", "Total Appointments", "#6a1b9a"),
                 ("today", "Today's Appointments", "#ef6c00"),
                 ("pending", "Pending Bills", "#c62828"),
                 ("amount", "Pending Amount (Rs)", "#00838f")]
        for i, (key, title, color) in enumerate(cards):
            card = tk.Frame(grid, bg=color, width=250, height=110)
            card.grid(row=i // 3, column=i % 3, padx=15, pady=15)
            card.pack_propagate(False)
            lbl = tk.Label(card, text="0", font=("Arial", 28, "bold"), bg=color, fg="white")
            lbl.pack(pady=(15, 0))
            tk.Label(card, text=title, font=("Arial", 11), bg=color, fg="white").pack()
            self.values[key] = lbl
        self.refresh()

    def refresh(self):
        today = str(date.today())
        appts = appointments.read()
        bills = [b for b in billing.read() if b[6] != "Paid"]
        total = 0.0
        for b in bills:
            try:
                total += float(b[3]) - float(b[4] or 0)
            except ValueError:
                pass
        self.values["patients"].config(text=str(len(patients.read())))
        self.values["doctors"].config(text=str(len(doctors.read())))
        self.values["appts"].config(text=str(len(appts)))
        self.values["today"].config(text=str(len([a for a in appts if a[3] == today and a[6] != "Cancelled"])))
        self.values["pending"].config(text=str(len(bills)))
        self.values["amount"].config(text=("%.2f" % total).rstrip("0").rstrip("."))


# ================= MAIN APPLICATION =================
root = tk.Tk()
root.title("Patient Management System")
root.configure(bg=BG)

style = ttk.Style()
style.theme_use("clam")
style.configure("TNotebook", background=BG)
style.configure("TNotebook.Tab", font=("Arial", 11, "bold"), padding=(16, 7))
style.map("TNotebook.Tab", background=[("selected", GREEN)],
          foreground=[("selected", "white")])
style.configure("Treeview", rowheight=24, font=("Arial", 10))
style.configure("Treeview.Heading", font=("Arial", 10, "bold"))


def clear_root():
    for w in root.winfo_children():
        w.destroy()


def show_login():
    clear_root()
    root.geometry("440x360")
    box = tk.Frame(root, bg=BG)
    box.pack(expand=True)
    tk.Label(box, text="Patient Management System", font=("Arial", 18, "bold"),
             bg=BG, fg=GREEN).pack(pady=(0, 5))
    tk.Label(box, text="Please login to continue", bg=BG).pack(pady=(0, 20))
    user_var, pass_var = tk.StringVar(), tk.StringVar()
    tk.Label(box, text="Username", bg=BG).pack()
    tk.Entry(box, textvariable=user_var, width=28).pack(pady=4)
    tk.Label(box, text="Password", bg=BG).pack()
    pw = tk.Entry(box, textvariable=pass_var, width=28, show="*")
    pw.pack(pady=4)

    def login(event=None):
        if user_var.get() == USERNAME and pass_var.get() == PASSWORD:
            build_main()
        else:
            messagebox.showerror("Login failed", "Wrong username or password.")

    tk.Button(box, text="Login", width=16, command=login, bg=GREEN, fg="white",
              font=("Arial", 10, "bold")).pack(pady=15)
    tk.Label(box, text="Default login:  admin / admin123", bg=BG, fg="gray").pack()
    root.bind("<Return>", login)


def build_main():
    clear_root()
    root.unbind("<Return>")
    root.geometry("1150x740")

    header = tk.Frame(root, bg=GREEN)
    header.pack(fill="x")
    tk.Label(header, text="Patient Management System", font=("Arial", 16, "bold"),
             bg=GREEN, fg="white").pack(side="left", padx=15, pady=8)
    tk.Button(header, text="Logout", command=show_login, bg="white", fg=GREEN,
              font=("Arial", 10, "bold")).pack(side="right", padx=15)

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=8, pady=8)

    sync_appointments()
    dash = Dashboard(nb)
    patient_tab = CrudTab(nb, patients, [
        {"label": "Patient Name", "kind": "text", "req": True},
        {"label": "Age", "kind": "int", "req": True, "max": 120},
        {"label": "Gender", "kind": "combo", "options": ["Male", "Female", "Other"]},
        {"label": "Phone", "kind": "phone", "req": True},
        {"label": "Address", "kind": "text"},
        {"label": "Blood Group", "kind": "combo",
         "options": ["A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"]},
        {"label": "Disease / Description", "kind": "text"},
        {"label": "Doctor Consulted", "kind": "ref", "source": doctor_choices, "req": True},
        {"label": "Tests Performed (select)", "kind": "multi", "source": test_choices}],
        on_change=dash.refresh,
        after_add=auto_bill_add, after_update=auto_bill_update,
        after_delete=remove_unpaid_bill,
        note="Bill = doctor's fee + selected tests (generated automatically). Ctrl+click to select many tests.")
    doctor_tab = CrudTab(nb, doctors, [
        {"label": "Doctor Name", "kind": "text", "req": True},
        {"label": "Specialization", "kind": "text", "req": True},
        {"label": "Phone", "kind": "phone", "req": True},
        {"label": "Consultation Fee", "kind": "number", "req": True}],
        on_change=dash.refresh)
    tests_tab = CrudTab(nb, tests, [
        {"label": "Test Name", "kind": "text", "req": True},
        {"label": "Price (Rs)", "kind": "number", "req": True}],
        on_change=dash.refresh)

    def cancel_appt(tab):
        if tab.selected is None:
            messagebox.showwarning("Select record", "Click an appointment in the table first.")
            return
        if messagebox.askyesno("Confirm", "Cancel this appointment?"):
            rows = appointments.read()
            for r in rows:
                if r[0] == tab.selected:
                    r[6] = "Cancelled"
            if appointments._write(rows):
                tab.changed()

    appt_tab = CrudTab(nb, appointments, [
        {"label": "Patient", "kind": "ref", "source": patient_choices, "req": True},
        {"label": "Doctor", "kind": "ref", "source": doctor_choices, "req": True},
        {"label": "Date (YYYY-MM-DD)", "kind": "date", "req": True},
        {"label": "Time (HH:MM)", "kind": "time", "req": True},
        {"label": "Reason", "kind": "text"}],
        on_change=dash.refresh,
        finalize=lambda v: v + [appt_status(v[2], v[3])],
        extra_buttons=[("Cancel Appointment", cancel_appt)],
        note="Status is automatic: future = Scheduled, past = Completed")
    bill_tab = CrudTab(nb, billing, [
        {"label": "Patient", "kind": "ref", "source": patient_choices, "req": True},
        {"label": "Description", "kind": "text", "req": True},
        {"label": "Amount (Rs)", "kind": "number", "req": True},
        {"label": "Amount Paid (Rs)", "kind": "number", "default": "0"},
        {"label": "Date (YYYY-MM-DD)", "kind": "date", "req": True}],
        on_change=dash.refresh, finalize=finalize_bill,
        note="Bills are generated automatically when a patient is added. Update Amount Paid to settle a bill.")

    nb.add(dash, text="Dashboard")
    nb.add(patient_tab, text="Patients")
    nb.add(doctor_tab, text="Doctors")
    nb.add(tests_tab, text="Tests")
    nb.add(appt_tab, text="Appointments")
    nb.add(bill_tab, text="Billing")

    def on_tab_change(event):
        if sync_appointments():
            appt_tab.load()
        dash.refresh()
        patient_tab.refresh_refs()
        appt_tab.refresh_refs()
        bill_tab.refresh_refs()
        bill_tab.load()

    nb.bind("<<NotebookTabChanged>>", on_tab_change)

    def tick():
        # every minute: update appointment status automatically
        if not nb.winfo_exists():
            return
        if sync_appointments():
            appt_tab.load()
            dash.refresh()
        root.after(60000, tick)

    root.after(60000, tick)


show_login()
root.mainloop()
