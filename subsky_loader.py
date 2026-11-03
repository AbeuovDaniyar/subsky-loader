"""Sub Sky Loader: turns a GT Trips file (.xlsx or .csv) into the Sub Sky loader workbook.

Flow: pick the trips file -> Save As (change folder and/or name) -> loader .xlsx is saved there,
plus a CSV copy in the same folder when the input is .xlsx -> option to open it.
"""
import csv
import datetime as dt
import json
import os
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

COLUMNS = ["Has Note", "Booking Id", "Provider Name", "Client Name", "Pickup Time", "Appt Time",
           "Pickup Polygon", "Phone Pickup", "Origin", "State (orig)", "Dropoff Polygon", "Phone Dropoff",
           "Destination", "State (dest)", "Subtype Abbr", "Direct Distance", "Passenger Types", "Space Types",
           "Provider Cost Dropoff", "Total Adjustments", "Total Provider Cost", "Mobility Aids", "Comments",
           "Pickup Comments", "Dropoff Comments", "Schedule Status"]
WIDTHS = [10, 11, 15, 22, 11, 10, 22, 16, 42, 10, 22, 16, 42, 10, 12, 14, 15, 12, 19, 16, 17, 14, 30, 16, 17, 15]
POLYGON = "HARTFORD COUNTY, CT"


# ---------- reading ----------

def norm(h):
    """Header match like the Excel query: ignore case, spaces, / . _ - and stray characters."""
    s = "" if h is None else str(h)
    s = s.replace("﻿", "").replace(" ", " ")
    s = re.sub(r"[\x00-\x1f\x7f]", "", s).strip().lower()
    return re.sub(r"[ /._-]", "", s)


def _pattern_section(fmt, v):
    """Pick the format section that applies, e.g. '[<=9999999]###-####;(###) ###-####'."""
    parts = fmt.split(";")
    for p in parts:
        m = re.match(r"\[([<>=]+)(-?[\d.]+)\]", p)
        if not m:
            continue
        op, lim = m.group(1), float(m.group(2))
        ok = {"<=": v <= lim, "<": v < lim, ">=": v >= lim, ">": v > lim, "=": v == lim}.get(op, False)
        if ok:
            return p[m.end():]
    rest = [p for p in parts if not re.match(r"\[[<>=]", p)]
    return rest[0] if rest else parts[-1]


def format_number(v, fmt):
    """Enough of Excel's number formats for trip exports: digit patterns (phones, ZIPs) and decimals."""
    if not fmt or fmt == "General":
        return f"{v:.15g}"
    sec = _pattern_section(fmt, v)
    sec = re.sub(r"\[[^\]]*\]", "", sec)          # colours / locale tags
    sec = re.sub(r'"([^"]*)"', r"\1", sec)        # quoted literals
    sec = re.sub(r"\\(.)", r"\1", sec)            # escaped literals
    sec = sec.replace("_)", "").replace("*", "")
    if "." not in sec and "," not in sec and isinstance(v, (int, float)) and float(v).is_integer() and v >= 0:
        slots = sum(ch in "#0" for ch in sec)
        if slots:
            digits = str(int(v))
            if len(digits) <= slots and re.search(r"[^#0]", sec):   # phone, SSN, ZIP+4 style
                min_len = sum(ch == "0" for ch in sec)
                digits = digits.zfill(min_len) if min_len > len(digits) else digits
                out, di = [], len(digits) - 1
                for ch in reversed(sec):
                    if ch in "#0":
                        if di >= 0:
                            out.append(digits[di]); di -= 1
                        elif ch == "0":
                            out.append("0")
                    else:
                        out.append(ch)
                return "".join(reversed(out)).strip()
            if re.fullmatch(r"0+", sec):                              # ZIP like 00000
                return digits.zfill(len(sec))
    m = re.search(r"\.(0+)", sec)
    if m:
        return f"{v:.{len(m.group(1))}f}"
    return f"{v:.15g}"


def cell_text(v, fmt="General"):
    """Show a cell the way Excel would in a CSV."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return format_number(v, fmt)
    if isinstance(v, dt.datetime):
        if v.time() == dt.time(0, 0):
            return v.strftime("%m/%d/%Y")
        return v.strftime("%m/%d/%Y %H:%M")
    if isinstance(v, dt.date):
        return v.strftime("%m/%d/%Y")
    if isinstance(v, dt.time):
        return v.strftime("%H:%M")
    return str(v)


def read_csv(path):
    raw = open(path, "rb").read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1252", errors="replace")
    return [row for row in csv.reader(text.splitlines())]


def read_xlsx(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    grid = [[cell_text(c.value, getattr(c, "number_format", "General") or "General") for c in row]
            for row in ws.iter_rows()]
    wb.close()
    while grid and not any(c.strip() for c in grid[-1]):
        grid.pop()
    width = max((len(r) for r in grid), default=0)
    # drop empty trailing columns
    while width and not any((r[width - 1] if len(r) >= width else "").strip() for r in grid):
        width -= 1
    return [r[:width] for r in grid]


def write_csv(grid, path):
    with open(path, "w", newline="", encoding="cp1252", errors="replace") as f:
        csv.writer(f).writerows(grid)


# ---------- converting ----------

def blank(v):
    return v is None or str(v).strip() == ""


def field(rec, names):
    for n in names:
        k = norm(n)
        if k in rec and not blank(rec[k]):
            return str(rec[k]).strip()
    return None


def fmt_time(v):
    """'12:00 EST' -> '12:00', '2:30 PM' -> '14:30', '9:40' -> '09:40'. Anything else stays as typed."""
    if blank(v):
        return None
    s = str(v).replace('"', "").strip()
    m = re.search(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp]\.?[Mm]\.?)?", s)
    if not m:
        return s
    h, mi = int(m.group(1)), int(m.group(2))
    if m.group(3):
        pm = m.group(3)[0] in "pP"
        if pm and h < 12:
            h += 12
        if not pm and h == 12:
            h = 0
    if h > 23 or mi > 59:
        return s
    return f"{h:02d}:{mi:02d}"


def clean_name(v):
    return None if blank(v) else re.sub(r"\s+", " ", str(v).replace(",", "")).strip()


def num(v):
    if blank(v):
        return None
    try:
        n = float(re.sub(r"[^0-9.\-]", "", str(v)))
        return int(n) if n.is_integer() else n
    except ValueError:
        return str(v).strip()


def convert(grid):
    """Returns (rows, format_name, warnings)."""
    known = {"legidleg#", "passengername", "name", "putime", "passenger", "pickupaddress", "contractname"}
    hi = next((i for i, r in enumerate(grid) if any(norm(c) in known for c in r)), 0)
    headers = [norm(c) for c in (grid[hi] if grid else [])]
    records = []
    for r in grid[hi + 1:]:
        if not any(not blank(c) for c in r):
            continue
        rec = {}
        for i, h in enumerate(headers):
            if h and h not in rec:
                rec[h] = r[i] if i < len(r) else None
        records.append(rec)

    def has(n):
        return norm(n) in headers

    is_leg = has("Leg ID/Leg #") or has("Passenger Name") or has("Name") or has("P/U Time")
    need = ([["Passenger Name", "Name"], ["Requested PickUp Time", "P/U Time"],
             ["Pickup Address", "P/U Address"], ["Dropoff Address", "D/O Address"]] if is_leg
            else [["Passenger"], ["Estimated Pickup Time"], ["Pickup Address"], ["Dropoff Address"]])
    warnings = [f'Column "{alts[0]}" not found, so that field is empty.' for alts in need if not any(has(a) for a in alts)]

    rows = []
    for i, r in enumerate(records):
        if is_leg:
            phone = field(r, ["Passenger Phone"])
            rows.append([0, i + 1, field(r, ["Vehicle Plate"]), clean_name(field(r, ["Passenger Name", "Name"])),
                         fmt_time(field(r, ["Requested PickUp Time", "P/U Time"])), None, POLYGON, phone,
                         field(r, ["Pickup Address", "P/U Address"]), "NH", POLYGON, phone,
                         field(r, ["Dropoff Address", "D/O Address"]), "NH", "SUB",
                         num(field(r, ["Estimated Miles"])), "CLI1", "TL1", 0, 0, 0,
                         field(r, ["Seating Need"]), field(r, ["Leg Notes"]), None, None, 10])
        else:
            notes = field(r, ["VA Notes"])
            rows.append([0, i + 1, field(r, ["Contract Name"]), field(r, ["Passenger"]),
                         fmt_time(field(r, ["Estimated Pickup Time"])), fmt_time(field(r, ["Appointment Time"])),
                         POLYGON, field(r, ["Passenger Phone"]), field(r, ["Pickup Address"]), "NH", POLYGON,
                         field(r, ["Phone Number"]), field(r, ["Dropoff Address"]), "NH", "SUB",
                         num(field(r, ["Estimated Distance (miles)"])), "CLI1", "TL1", 0, 0, 0,
                         field(r, ["Mobility"]), notes, notes, notes, 10])
    return rows, ("Leg export" if is_leg else "VA manifest"), warnings


# ---------- writing ----------

def sheet_name_for(path):
    base = os.path.splitext(os.path.basename(path))[0]
    base = re.sub(r"csc$", "", base, flags=re.I).strip()
    m = re.search(r"sub\s*sky.*$", base, flags=re.I)
    name = re.sub(r"[\[\]:*?/\\]", " ", m.group(0) if m else base)
    return re.sub(r"\s+", " ", name).strip()[:31] or "Sheet1"


def table_name_for(sheet):
    t = re.sub(r"_+", "_", re.sub(r"[^A-Za-z0-9_]", "_", sheet))
    if not re.match(r"^[A-Za-z_]", t) or re.match(r"^[A-Za-z]{1,3}\d+$", t):
        t = "T_" + t
    return t


def write_loader(rows, sheet, path):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(COLUMNS)
    for r in rows:
        ws.append(r)
    end = f"Z{len(rows) + 1}"
    tbl = Table(displayName=table_name_for(sheet), ref=f"A1:{end}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium7", showRowStripes=True)
    ws.add_table(tbl)
    for i, w in enumerate(WIDTHS):
        ws.column_dimensions[ws.cell(1, i + 1).column_letter].width = w
    wb.save(path)


def free_path(path):
    """Don't overwrite a file that's open in Excel: fall back to 'name (2).xlsx' etc."""
    root, ext = os.path.splitext(path)
    n = 2
    while True:
        try:
            if os.path.exists(path):
                with open(path, "a"):
                    pass
            return path
        except PermissionError:
            path = f"{root} ({n}){ext}"
            n += 1


def load(path):
    """Reads and converts one trips file. Returns (grid, rows, format, warnings)."""
    grid = read_csv(path) if path.lower().endswith(".csv") else read_xlsx(path)
    rows, fmt, warnings = convert(grid)
    if not rows:
        raise ValueError("No trips were found in this file.")
    return grid, rows, fmt, warnings


def save(src_path, grid, rows, out_path):
    """Writes the loader to out_path; for .xlsx input also writes a CSV copy into the same folder.
    The sheet is named after the saved file. Returns the CSV path (or None)."""
    if os.path.abspath(out_path).lower() == os.path.abspath(src_path).lower():
        raise ValueError("That's the trips file you opened. Pick a different name so it isn't overwritten.")
    sheet = re.sub(r"[\[\]:*?/\\]", " ", os.path.splitext(os.path.basename(out_path))[0]).strip()[:31] or "Sheet1"
    write_loader(rows, sheet, out_path)
    csv_path = None
    if not src_path.lower().endswith(".csv"):
        csv_path = os.path.join(os.path.dirname(out_path), os.path.splitext(os.path.basename(src_path))[0] + ".csv")
        csv_path = free_path(csv_path)
        write_csv(grid, csv_path)
    return csv_path


def process(path, out_path=None):
    """Whole job with default names, next to the input (used for testing)."""
    grid, rows, fmt, warnings = load(path)
    if out_path is None:
        out_path = free_path(os.path.join(os.path.dirname(path), sheet_name_for(path) + ".xlsx"))
    csv_path = save(path, grid, rows, out_path)
    return {"csv": csv_path, "xlsx": out_path, "trips": len(rows), "format": fmt, "warnings": warnings}


# ---------- settings (remembers folders between runs) ----------

SETTINGS = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "SubSkyLoader", "settings.json")


def load_settings():
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_settings(s):
    try:
        os.makedirs(os.path.dirname(SETTINGS), exist_ok=True)
        with open(SETTINGS, "w", encoding="utf-8") as f:
            json.dump(s, f)
    except OSError:
        pass


# ---------- window ----------

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    settings = load_settings()
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")

    def existing(d):
        return d if d and os.path.isdir(d) else (downloads if os.path.isdir(downloads) else None)

    while True:
        path = filedialog.askopenfilename(
            title="Sub Sky Loader - choose today's GT Trips file",
            initialdir=existing(settings.get("open_dir")),
            filetypes=[("Trips files", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")])
        if not path:
            break
        path = os.path.normpath(path)
        settings["open_dir"] = os.path.dirname(path)
        try:
            grid, rows, fmt, warnings = load(path)
        except Exception as e:  # show any problem to the user instead of crashing
            if not messagebox.askretrycancel("Sub Sky Loader", f"Couldn't read that file:\n\n{e}\n\nTry another file?"):
                break
            continue

        # Save As: suggested name and folder, both can be changed
        last = settings.get("save_dir")
        save_dir = last if last and os.path.isdir(last) else os.path.dirname(path)
        suggested = sheet_name_for(path) + ".xlsx"
        out_path = None
        while True:
            out_path = filedialog.asksaveasfilename(
                title=f"Save loader ({len(rows)} trips) - choose folder and file name",
                initialdir=save_dir, initialfile=suggested,
                defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
            if not out_path:
                break
            out_path = os.path.normpath(out_path)
            if not out_path.lower().endswith(".xlsx"):
                out_path += ".xlsx"
            try:
                csv_path = save(path, grid, rows, out_path)
                break
            except PermissionError:
                messagebox.showwarning("Sub Sky Loader",
                                       f"{os.path.basename(out_path)} is open in Excel.\n\nClose it, or choose a different name.")
            except Exception as e:
                messagebox.showwarning("Sub Sky Loader", f"Couldn't save:\n\n{e}")
            save_dir, suggested = os.path.dirname(out_path), os.path.basename(out_path)
        if not out_path:  # cancelled the save
            save_settings(settings)
            if not messagebox.askyesno("Sub Sky Loader", "Nothing was saved.\n\nConvert another file?"):
                break
            continue

        settings["save_dir"] = os.path.dirname(out_path)
        save_settings(settings)
        lines = [f"Done: {len(rows)} trips ({fmt}).", "",
                 f"Loader saved as:\n{out_path}"]
        if csv_path:
            lines += ["", f"CSV saved as:\n{csv_path}"]
        if warnings:
            lines += ["", "Check:"] + ["- " + w for w in warnings]
        lines += ["", "Open the loader now?"]
        if messagebox.askyesno("Sub Sky Loader", "\n".join(lines)):
            os.startfile(out_path)
        if not messagebox.askyesno("Sub Sky Loader", "Convert another file?"):
            break
    save_settings(settings)
    root.destroy()


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--cli":  # used for testing: --cli input [output.xlsx]
        print(process(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None))
    else:
        run_gui()
