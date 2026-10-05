"""
SAP Sales & Logistics Workflow Automation
---------------------------------------
This script automates the extraction of sales reports from SAP ERP via GUI Scripting,
processes the exported Excel file (calculating net EUR values, checking IncoTerms), 
dynamically prompts the user for missing freight costs via a Tkinter UI, 
and prepares the final structured management report.

Author: Nándor Magyar
Disclaimer: All company-specific data, T-codes, user IDs, and internal paths 
have been anonymized for public sharing. 
"""

import win32com.client, sys, time, os, subprocess, datetime, glob, shutil, tempfile
import tkinter as tk
from tkinter import ttk, messagebox
import logging

# Log to file AND console: if a row fails in production, there is a trace of it
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(os.path.expanduser("~"), "sap_automation.log"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)

TEST_FORCE_MISSING_FREIGHT = False

# --- Anonymized File Paths ---
PATH_SAVE_FOLDER = os.path.expanduser(r"~\Documents\Logistics_Reports") 
current_year = datetime.datetime.now().year
PATH_DAILY_PRICES_FOLDER = rf"C:\Users\Public\Documents\Daily_Prices\{current_year}- Base Price" 

# --- Anonymized SAP System Configuration ---
SAP_CFG_SYSTEM_NAME    = "SAP_ERP_PRODUCTION"  
SAP_CFG_CLIENT         = "400"
SAP_CFG_LANGUAGE       = "EN"
SAP_TRANSACTION_CODE   = "Z_CUSTOM_SALES_REP"  
SAP_PLANT_ENTITY       = "P1000"               

SAP_EXE_PATHS = [
    r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 
    r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"       
]

FINAL_COLUMN_ORDER = ["MC num", "Doc-date", "Ship start", "Ctr/nom. Qty", "Flat Price", "Flat Price Currency", "R-Truck freight", "R-Truck freight CUR", "EUR/HUF", "Flat EUR", "freight EUR", "Net Flat EUR", "Commodity desc", "Partner name", "Person responsible", "Type", "Base loc name", "Inco1", "Inco2"]

LIST_VALID_USERS = ["USER01", "USER02", "USER03", "USER04", "USER05", "USER06", "USER07"]
LIST_IGNORE_PARTNERS = ["INTERNAL_COMPANY_A", "INTERNAL_COMPANY_B"]

def get_credentials():
    config_file = os.path.join(os.path.expanduser("~"), "sap_config.txt")
    if not os.path.exists(config_file):
        with open(config_file, "w", encoding="utf-8") as f:
            f.write("USERNAME=Your_SAP_Username\nPASSWORD=Your_SAP_Password\n")
        print(f"\n[!] INFO: Created config template at: {config_file}\nPlease fill it and restart.\n")
        sys.exit()
    user, pwd = "", ""
    with open(config_file, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip().startswith("USERNAME="): user = line.strip().split("=", 1)[1]
            elif line.strip().startswith("PASSWORD="): pwd = line.strip().split("=", 1)[1]
    if not user or user == "Your_SAP_Username" or not pwd or pwd == "Your_SAP_Password":
        print(f"\n[!] ERROR: Please update credentials in {config_file}")
        sys.exit()
    return user, pwd

MY_SAP_USER, MY_SAP_PASSWORD = get_credentials()

def ask_freight_popup(mc_num, partner, incoterm_full, default_curr="EUR"):
    result = {"val": None, "curr": default_curr}
    root = tk.Tk()
    root.title("Missing Freight Rate")
    root.attributes('-topmost', True) 
    root.geometry("380x240")
    root.eval('tk::PlaceWindow . center')

    tk.Label(root, text=f"Freight rate required!", font=("Arial", 11, "bold"), fg="red").pack(pady=5)
    tk.Label(root, text=f"IncoTerm: {incoterm_full}", font=("Arial", 10, "bold")).pack(pady=2)
    tk.Label(root, text=f"Contract/MC: {mc_num}").pack()
    tk.Label(root, text=f"Partner: {partner}").pack()
    
    frame = tk.Frame(root)
    frame.pack(pady=15)
    tk.Label(frame, text="Rate:").grid(row=0, column=0, padx=5)
    entry_val = tk.Entry(frame, width=12)
    entry_val.grid(row=0, column=1, padx=5)
    entry_val.focus()
    
    curr_var = tk.StringVar(value=default_curr if default_curr in ["EUR", "HUF", "USD"] else "EUR")
    cb_curr = ttk.Combobox(frame, textvariable=curr_var, values=["EUR", "HUF", "USD"], width=6, state="readonly")
    cb_curr.grid(row=0, column=2, padx=5)
    
    def on_submit(event=None):
        raw_val = entry_val.get().replace(',', '.') 
        if not raw_val.strip():
            messagebox.showwarning("Warning", "Please enter a valid numeric rate!")
            return
        try:
            result["val"] = float(raw_val)
            result["curr"] = curr_var.get()
            root.destroy()
        except ValueError:
            messagebox.showerror("Error", "Invalid number format!")
            
    root.bind('<Return>', on_submit)
    tk.Button(root, text="Save & Continue", command=on_submit, bg="lightblue", font=("Arial", 9, "bold")).pack()
    root.mainloop()
    return result["val"], result["curr"]

def get_sap_application():
    try: return win32com.client.GetObject("SAPGUI").GetScriptingEngine
    except: pass
    sap_exe = next((p for p in SAP_EXE_PATHS if os.path.exists(p)), None)
    if not sap_exe: return None
    subprocess.Popen(sap_exe)
    for _ in range(15): 
        time.sleep(1)
        try: return win32com.client.GetObject("SAPGUI").GetScriptingEngine
        except: pass 
    return None

def get_run_configs():
    today = datetime.date.today()
    if today.weekday() == 0:
        s = today - datetime.timedelta(days=7); e = today - datetime.timedelta(days=3)
        return [{"sap_start": s.strftime("%d.%m.%Y"), "sap_end": e.strftime("%d.%m.%Y"), "filename": f"weekly_report_{today.strftime('%Y-%m-%d')}"}]
    else: 
        s = today - datetime.timedelta(days=1); e = today - datetime.timedelta(days=1)
        return [{"sap_start": s.strftime("%d.%m.%Y"), "sap_end": e.strftime("%d.%m.%Y"), "filename": f"daily_report_{today.strftime('%Y-%m-%d')}"}]

def get_column_mapping(ws):
    m, m_lower = {}, {}
    for col in range(1, 200):
        val = ws.Cells(1, col).Value
        if val: m[str(val).strip()] = col; m_lower[str(val).strip().lower()] = col
        elif col > 80: break
    return m, m_lower

def find_col_index(m, m_lower, t_name):
    if t_name in m: return m[t_name]
    t_lower = t_name.lower()
    if t_lower in m_lower: return m_lower[t_lower]
    for h, c in m_lower.items():
        if t_lower in h: return c
    return None

def add_months(source_date, months):
    month = source_date.month - 1 + months
    year = source_date.year + month // 12
    month = month % 12 + 1
    day = min(source_date.day, [31, 29 if year % 4 == 0 and not year % 100 == 0 or year % 400 == 0 else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
    return datetime.datetime(year, month, day)

def parse_excel_date(value):
    if value is None: return None
    if hasattr(value, 'year') and hasattr(value, 'month') and hasattr(value, 'day'): return datetime.datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        val_str = value.strip().replace('-', '.').replace('/', '.')
        try:
            parts = val_str.split('.')
            if len(parts) >= 3:
                p0, p1, p2 = int(parts[0]), int(parts[1]), int(parts[2][:4])
                if p0 > 1000: return datetime.datetime(p0, p1, p2)
                elif p2 > 1000: return datetime.datetime(p2, p1, p0)
        except: pass
    return None

def lookup_rate(target_date, rate_table, col_index):
    best_match = None
    if target_date: target_date = target_date.replace(tzinfo=None)
    for row in rate_table:
        row_date = parse_excel_date(row[0]) 
        if not row_date: continue
        if row_date <= target_date: best_match = row
        else: break
    if best_match:
        try: return float(best_match[col_index])
        except: pass
    return 0.0

def reorder_columns_absolute(ws):
    try: ws.AutoFilterMode = False
    except: pass
    name_map = {"Flat Price Currency": "Flat Price CUR", "R-Truck freight CUR": "R-Truck freight CUR", "R-Truck freight": "R-Truck freight"}
    curr_idx = 1
    for desired_name in FINAL_COLUMN_ORDER:
        search_name = name_map.get(desired_name, desired_name).lower().strip()
        found_idx = next((col for col in range(1, 250) if ws.Cells(1, col).Value and str(ws.Cells(1, col).Value).strip().lower() == search_name), None)
        if found_idx:
            if found_idx != curr_idx: ws.Columns(found_idx).Cut(); ws.Columns(curr_idx).Insert()
            ws.Cells(1, curr_idx).Value = desired_name
        else:
            ws.Columns(curr_idx).Insert()
            ws.Cells(1, curr_idx).Value = desired_name
            ws.Cells(1, curr_idx).Interior.Color = 65535 
        curr_idx += 1

def get_crop_year(ship_start_date, comm_short):
    if not ship_start_date: return "No Date"
    try:
        y = ship_start_date.year
        threshold = datetime.datetime(y, 7, 15) if comm_short == "RSM" else datetime.datetime(y, 9, 1)
        return f"{y}/{str(y+1)[-2:]}" if ship_start_date >= threshold else f"{y-1}/{str(y)[-2:]}"
    except: return "Invalid Date"

def draw_summary_table(ws, start_row, start_col, title, stats_data, unique_years):
    ws.Cells(start_row - 1, start_col).Value = title
    ws.Cells(start_row - 1, start_col).Font.Bold = True
    for i, h_text in enumerate(["Base location", "Commodity", "Crop year", "Quantity", "Price"]):
        cell = ws.Cells(start_row, start_col + i)
        cell.Value, cell.Interior.Color, cell.Font.Color, cell.Font.Bold = h_text, 4600064, 16777215, True
        cell.HorizontalAlignment, cell.VerticalAlignment, cell.Borders.LineStyle = -4108, -4108, 1

    current_row = start_row + 1
    for loc in ["Location_A", "Location_B"]:
        loc_start_row = current_row
        ws.Cells(loc_start_row, start_col).Value = loc
        for comm in ["RSM", "SFM"]:
            comm_start_row = current_row
            ws.Cells(comm_start_row, start_col + 1).Value = comm
            for cy in unique_years:
                total_qty, total_val = stats_data.get((loc, comm, cy), [0.0, 0.0])
                ws.Cells(current_row, start_col + 2).Value = cy
                if total_qty > 0:
                    ws.Cells(current_row, start_col + 3).Value = round(total_qty)
                    ws.Cells(current_row, start_col + 4).Value = round(total_val / total_qty)
                for c in range(5):
                    cell = ws.Cells(current_row, start_col + c)
                    cell.Borders.LineStyle = 1; cell.VerticalAlignment = -4108
                    if c in [0, 1]: cell.Interior.Color = 52479; cell.Font.Bold = True; cell.HorizontalAlignment = -4108
                    if c in [2, 3, 4]: cell.HorizontalAlignment = -4108
                    if c == 3: cell.NumberFormat = "0"
                    if c == 4: cell.NumberFormat = '0" €"'
                current_row += 1
            if current_row > comm_start_row + 1:
                try: ws.Range(ws.Cells(comm_start_row, start_col + 1), ws.Cells(current_row - 1, start_col + 1)).Merge()
                except: pass
        if current_row > loc_start_row + 1:
            try: lr = ws.Range(ws.Cells(loc_start_row, start_col), ws.Cells(current_row - 1, start_col)); lr.Merge(); lr.Font.Italic = True
            except: pass

def process_excel_modifications(xl_app, wb_report, filename_name):
    if wb_report is None: return
    print(f"\n[INFO] Excel processing started for: {wb_report.Name}")
    try: wb_report.Activate(); ws = wb_report.Sheets(1); xl_app.DisplayAlerts = False
    except Exception as e: print(f"[ERROR] Excel init failed: {e}"); return 

    try: ws.Unprotect()
    except: pass
    try: ws.AutoFilterMode = False
    except: pass

    col_map, col_map_lower = get_column_mapping(ws)
    last_row = ws.Cells(ws.Rows.Count, "C").End(3).Row

    if type_col := find_col_index(col_map, col_map_lower, "Type"):
        for r in range(last_row, 1, -1):
            if str(ws.Cells(r, type_col).Value).strip().upper() == 'P': ws.Rows(r).Delete()
        last_row = ws.Cells(ws.Rows.Count, "C").End(3).Row 
        col_map, col_map_lower = get_column_mapping(ws) 

    ws.Columns("G:H").Insert(); ws.Columns("K:K").Insert(); ws.Columns("L:L").Insert()
    ws.Cells(1, 7).Value = "EUR/HUF"; ws.Cells(1, 8).Value = "Flat EUR"; ws.Cells(1, 11).Value = "freight EUR"; ws.Cells(1, 12).Value = "Net Flat EUR" 
    col_map, col_map_lower = get_column_mapping(ws) 

    mc_col, inco_col, inco2_col = find_col_index(col_map, col_map_lower, "MC num"), find_col_index(col_map, col_map_lower, "Inco1"), find_col_index(col_map, col_map_lower, "Inco2")
    fr_col, curr_col = find_col_index(col_map, col_map_lower, "R-Truck freight"), find_col_index(col_map, col_map_lower, "R-Truck freight CUR")
    rw_col, partner_col = find_col_index(col_map, col_map_lower, "R-RW freight"), find_col_index(col_map, col_map_lower, "Partner name")
    price_curr_col = find_col_index(col_map, col_map_lower, "Flat Price Currency") or find_col_index(col_map, col_map_lower, "Flat Price CUR")

    freight_cache = {}
    def is_empty_val(v):
        if v is None or str(v).strip() == "": return True
        try: return float(v) == 0.0
        except: return False

    if mc_col and fr_col and curr_col:
        for r in range(2, last_row + 1):
            m, f, c = str(ws.Cells(r, mc_col).Value or "").strip(), ws.Cells(r, fr_col).Value, str(ws.Cells(r, curr_col).Value or "").strip()
            base_m = m.split('-')[0].strip() if m else "" 
            if base_m and not is_empty_val(f): freight_cache[base_m] = (float(f), c)

    test_triggered = False
    if mc_col and inco_col and fr_col and curr_col:
        for r in range(2, last_row + 1):
            inco = str(ws.Cells(r, inco_col).Value or "").strip().upper()
            if inco and inco not in ["FCA", "NONE"]:
                f_truck, f_rw = ws.Cells(r, fr_col).Value, ws.Cells(r, rw_col).Value if rw_col else None
                m = str(ws.Cells(r, mc_col).Value or "").strip()
                base_m = m.split('-')[0].strip() if m else "" 
                incoterm_full = f"{inco} {str(ws.Cells(r, inco2_col).Value or '').strip() if inco2_col else ''}".strip()
                contract_curr = str(ws.Cells(r, price_curr_col).Value).strip().upper() if price_curr_col and ws.Cells(r, price_curr_col).Value else "EUR"
                
                if TEST_FORCE_MISSING_FREIGHT and not test_triggered and base_m:
                    f_truck = None 
                    if base_m in freight_cache: del freight_cache[base_m] 
                    test_triggered = True 

                if is_empty_val(f_truck) and is_empty_val(f_rw):
                    if not base_m: continue 
                    partner = str(ws.Cells(r, partner_col).Value or "").strip() if partner_col else ""
                    if base_m in freight_cache: val, curr = freight_cache[base_m]
                    else: 
                        val, curr = ask_freight_popup(m, partner, incoterm_full, contract_curr)
                        if val is not None:
                            freight_cache[base_m] = (val, curr)
                            for prev_r in range(2, r):
                                prev_m = str(ws.Cells(prev_r, mc_col).Value or "").strip()
                                if (prev_m.split('-')[0].strip() if prev_m else "") == base_m:
                                    ws.Cells(prev_r, fr_col).Value, ws.Cells(prev_r, curr_col).Value = val, curr
                        else: continue 
                    ws.Cells(r, fr_col).Value, ws.Cells(r, curr_col).Value = val, curr

    doc_date_col, person_col = find_col_index(col_map, col_map_lower, "Doc-date") or 2, find_col_index(col_map, col_map_lower, "Person responsible")
    needed_dates = set()
    for i in range(2, last_row + 1):
        if person_col and str(ws.Cells(i, person_col).Value or "").strip().upper() not in LIST_VALID_USERS: continue
        if d := parse_excel_date(ws.Cells(i, doc_date_col).Value): needed_dates.add(d)

    rate_cache, last_gyari_date = {}, datetime.datetime(1900, 1, 1)
    for d in needed_dates:
        files = glob.glob(os.path.join(PATH_DAILY_PRICES_FOLDER, f"*{d.strftime('%Y %m %d')}*.xls*")) + glob.glob(os.path.join(PATH_DAILY_PRICES_FOLDER, f"*{d.strftime('%Y.%m.%d')}*.xls*")) + glob.glob(os.path.join(PATH_DAILY_PRICES_FOLDER, f"*{d.strftime('%Y-%m-%d')}*.xls*"))
        if not files: continue
        fp = max(files, key=os.path.getctime)
        if d > last_gyari_date: last_gyari_date = d
        if d.strftime("%Y-%m-%d") in rate_cache: continue

        temp_fp = os.path.join(tempfile.gettempdir(), f"TEMP_{int(time.time())}_{os.path.basename(fp)}")
        try:
            shutil.copy2(fp, temp_fp)
            wb_temp = xl_app.Workbooks.Open(temp_fp, ReadOnly=True, UpdateLinks=False)
            ws_indication = next((s for s in wb_temp.Worksheets if "INDICATION" in s.Name.upper()), None)
            if ws_indication: rate_cache[d.strftime("%Y-%m-%d")] = [row for row in ws_indication.Range("M2:R19").Value if row[0]]
            wb_temp.Close(False)
        except: pass
        finally:
            if os.path.exists(temp_fp): 
                try: os.remove(temp_fp)
                except: pass

    stats, stats_friday = {}, {}
    def calc_eur(val, curr, e_rate, u_rate):
        if val is None or str(val).strip() == "": return 0.0
        try: v = float(val)
        except: return 0.0
        if v == 0.0: return 0.0
        c = str(curr).strip().upper() if curr else "EUR"
        if c == "EUR": return v
        elif c == "HUF" and e_rate > 0: return v / e_rate
        elif c == "USD" and e_rate > 0 and u_rate > 0: return (v * u_rate) / e_rate
        return 0.0

    for i in range(2, last_row + 1):
        try: 
            def get_val(tn): return ws.Cells(i, find_col_index(col_map, col_map_lower, tn)).Value if find_col_index(col_map, col_map_lower, tn) else None
            doc_date, ship_start = parse_excel_date(ws.Cells(i, doc_date_col).Value), parse_excel_date(get_val("Ship start"))
            qty, val_mat, curr_mat = get_val("Ctr/nom. Qty") or 0, get_val("Flat Price"), get_val("Flat Price CUR")
            val_truck, curr_truck, val_rw, curr_rw, val_ins, curr_ins = get_val("R-Truck freight"), get_val("R-Truck freight CUR"), get_val("R-RW freight"), get_val("R-RW freight CUR"), get_val("R-Insurance"), get_val("R-Insurance CUR")
            partner_name, user_name, incoterm = str(get_val("Partner name")).strip().upper(), str(get_val("Person responsible")).strip().upper(), str(get_val("Inco1")).strip().upper()

            eur_rate, usd_rate = 0.0, 0.0
            if doc_date and ship_start and (cache_key := doc_date.strftime("%Y-%m-%d")) in rate_cache:
                lkp = add_months(ship_start, 2)
                usd_rate, eur_rate = lookup_rate(lkp, rate_cache[cache_key], 3), lookup_rate(lkp, rate_cache[cache_key], 5)

            if eur_col_idx := find_col_index(col_map, col_map_lower, "EUR/HUF"): ws.Cells(i, eur_col_idx).Value = eur_rate
            row_is_red = (str(curr_mat).strip().upper() == "HUF" and val_mat is not None and val_mat < 50000) or any(bp in partner_name for bp in LIST_IGNORE_PARTNERS) or (incoterm in ["DDP", "CPT"] and is_empty_val(val_truck) and is_empty_val(val_rw))
            if row_is_red: ws.Cells(i, 200).Value = "RED"

            flat_eur = calc_eur(val_mat, curr_mat, eur_rate, usd_rate)
            if fe_idx := find_col_index(col_map, col_map_lower, "Flat EUR"): ws.Cells(i, fe_idx).Value = flat_eur
            freight_eur = calc_eur(val_truck, curr_truck, eur_rate, usd_rate) + calc_eur(val_rw, curr_rw, eur_rate, usd_rate) + calc_eur(val_ins, curr_ins, eur_rate, usd_rate)
            if fre_idx := find_col_index(col_map, col_map_lower, "freight EUR"): ws.Cells(i, fre_idx).Value = freight_eur
            net_eur = flat_eur - freight_eur
            if nfe_idx := find_col_index(col_map, col_map_lower, "Net Flat EUR"): ws.Cells(i, nfe_idx).Value = net_eur

            if not row_is_red and (user_name in LIST_VALID_USERS):
                raw_comm, plant_val = str(get_val("Commodity desc")).strip().upper(), str(get_val("Base loc name")).strip().upper()
                loc_short = "Location_A" if "LOC_A" in plant_val else ("Location_B" if "LOC_B" in plant_val else None) # Anonymized
                comm_short = "RSM" if "RAPE" in raw_comm or "RSM" in raw_comm else ("SFM" if "SUN" in raw_comm or "SFM" in raw_comm else None)
                if loc_short and comm_short:
                    stat_key = (loc_short, comm_short, get_crop_year(ship_start, comm_short))
                    if stat_key not in stats: stats[stat_key] = [0.0, 0.0] 
                    stats[stat_key][0] += float(qty or 0); stats[stat_key][1] += float(qty or 0) * float(net_eur or 0)
                    if "weekly" in filename_name.lower() and doc_date and doc_date.weekday() == 4:
                        if stat_key not in stats_friday: stats_friday[stat_key] = [0.0, 0.0]
                        stats_friday[stat_key][0] += float(qty or 0); stats_friday[stat_key][1] += float(qty or 0) * float(net_eur or 0)
        except Exception as row_err:
            logging.warning(f"Row {i} skipped due to error: {row_err}")
            continue

    try: reorder_columns_absolute(ws)
    except: pass
    col_map_final, col_map_lower_final = get_column_mapping(ws)
    for col_name in ["Flat EUR", "freight EUR", "Net Flat EUR", "EUR/HUF"]:
        if idx := find_col_index(col_map_final, col_map_lower_final, col_name): ws.Range(f"{ws.Cells(1, idx).Address.split('$')[1]}2:{ws.Cells(1, idx).Address.split('$')[1]}{last_row}").NumberFormatLocal = "0.00"
    for i in range(2, last_row + 1):
        if ws.Cells(i, 200).Value == "RED": ws.Rows(i).Interior.Color = 255 + (200 * 256) + (200 * 65536); ws.Cells(i, 200).Value = "" 

    if filter_col := find_col_index(col_map_final, col_map_lower_final, "Person responsible"): ws.Range(f"A1:AZ{last_row}").AutoFilter(Field=filter_col, Criteria1=LIST_VALID_USERS, Operator=7)
    ws.Columns("A:AZ").EntireColumn.AutoFit()

    unique_years = sorted(list(set(["2025/26", "2026/27"]) | {k[2] for k in stats.keys() if k[2] not in ["No Date", "Invalid Date"]}))
    draw_summary_table(ws, last_row + 4, 2, "WEEKLY SUMMARY" if "weekly" in filename_name.lower() else "DAILY SUMMARY", stats, unique_years)
    if "weekly" in filename_name.lower(): draw_summary_table(ws, last_row + 4, 8, "FRIDAY SALES ONLY", stats_friday, unique_years)
    ws.Columns("B:L").AutoFit()
    for col in ["B", "H"]: ws.Columns(col).ColumnWidth = 14
    for col in ["C", "D", "E", "F", "I", "J", "K", "L"]: ws.Columns(col).ColumnWidth = 12

    if not os.path.exists(PATH_SAVE_FOLDER): os.makedirs(PATH_SAVE_FOLDER)
    base_filename = os.path.join(PATH_SAVE_FOLDER, f"{filename_name}.xlsx")
    try:
        if os.path.exists(base_filename): os.remove(base_filename)
        wb_report.SaveAs(base_filename, FileFormat=51)
    except:
        wb_report.SaveAs(os.path.join(PATH_SAVE_FOLDER, f"{filename_name}_{datetime.datetime.now().strftime('%H%M%S')}.xlsx"), FileFormat=51)
    xl_app.DisplayAlerts = True

def sap_login_and_run():
    application = get_sap_application()
    if application is None: return
    try:
        if application.Children.Count > 0: session = application.Children(0).Children(0)
        else: session = application.OpenConnection(SAP_CFG_SYSTEM_NAME, True).Children(0)
    except: return

    try:
        if session.findById("wnd[0]/usr/txtRSYST-BNAME").text == "" or session.findById("wnd[0]/usr/pwdRSYST-BCODE").text == "":
            session.findById("wnd[0]/usr/txtRSYST-MANDT").text, session.findById("wnd[0]/usr/txtRSYST-BNAME").text, session.findById("wnd[0]/usr/pwdRSYST-BCODE").text, session.findById("wnd[0]/usr/txtRSYST-LANGU").text = SAP_CFG_CLIENT, MY_SAP_USER, MY_SAP_PASSWORD, SAP_CFG_LANGUAGE
            session.findById("wnd[0]").sendVKey(0) 
            time.sleep(3) 
    except:
        try:
            if session.ActiveWindow.Text.startswith("Information") or session.ActiveWindow.Text.startswith("Warning"):
               session.findById("wnd[1]/tbar[0]/btn[0]").press(); time.sleep(1)
        except: pass

    for config in get_run_configs():
        try:
            try: session.findById("wnd[0]/tbar[0]/okcd").text = "/n"; session.findById("wnd[0]").sendVKey(0); time.sleep(1)
            except: pass
            session.findById("wnd[0]/tbar[0]/okcd").text = SAP_TRANSACTION_CODE
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/cmbP_ENTITY").key = SAP_PLANT_ENTITY
            session.findById("wnd[0]/usr/ctxtP_FROMDATE").text, session.findById("wnd[0]/usr/ctxtP_TODATE").text = config["sap_start"], config["sap_end"]
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            # Note: Layout selection relies on SAP default user settings for security compliance.
            
            # Remember already-open workbooks so the freshly exported one can be identified
            try: existing_wbs = [wb.Name for wb in win32com.client.GetActiveObject("Excel.Application").Workbooks]
            except Exception: existing_wbs = []

            session.findById("wnd[0]/usr/cntlGV_ALV_CONT/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
            session.findById("wnd[0]/usr/cntlGV_ALV_CONT/shellcont/shell").selectContextMenuItem("&XXL")
            try: session.findById("wnd[1]/tbar[0]/btn[12]").press()
            except: pass

            time.sleep(6) 
            
            excel_ready, xl_app, wb_report = False, None, None
            for _ in range(15): 
                try:
                    xl_app = win32com.client.GetActiveObject("Excel.Application")
                    xl_app.Visible = True
                    for wb in xl_app.Workbooks:
                        if wb.Name not in existing_wbs: wb_report = wb; break
                    if wb_report is None: wb_report = xl_app.ActiveWorkbook
                    if wb_report is None: raise Exception("Empty")
                    xl_app.DisplayAlerts = False 
                    xl_app.DisplayAlerts = True 
                    excel_ready = True
                    break 
                except: time.sleep(2) 
            
            if not excel_ready or wb_report is None: continue 
            process_excel_modifications(xl_app, wb_report, config["filename"])
            
            # --- Map Generation Prompt ---
            base_filename = os.path.join(PATH_SAVE_FOLDER, f"{config['filename']}.xlsx")
            root_map = tk.Tk()
            root_map.withdraw() 
            root_map.attributes('-topmost', True)
            
            if messagebox.askyesno(title="Generate Logistics Map", message=f"Report {config['filename']} is ready!\n\nWould you like to generate an interactive HTML logistics map?"):
                try: 
                    import map_generator 
                    map_generator.generate_map(base_filename)
                except Exception as e: print(f"[WARNING] Map gen failed: {e}")
            
            root_map.destroy()
        except Exception as run_err:
            logging.error(f"Run failed for config {config.get('filename', '?')}: {run_err}")
            continue

if __name__ == "__main__":
    sap_login_and_run()
