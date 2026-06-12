"""
Logistics Routing Map Generator
-------------------------------
Generates an interactive HTML map visualizing sales and delivery volumes.
Draws routes from fulfillment bases to destinations and places dynamic 
pie-chart markers to represent commodity breakdowns.

Uses Folium (Leaflet.js) and Geopy (Nominatim API).
Author: [Your Name]
"""

import pandas as pd
import folium
from geopy.geocoders import Nominatim
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import webbrowser
import shutil
import tempfile
import math
from html import escape

# --- Anonymized Base Locations ---
BASES = {
    "BASE_A": {"lat": 47.0000, "lon": 20.2833}, # Example coords
    "BASE_B": {"lat": 48.0245, "lon": 16.7784}  # Example coords
}

CACHE_FILE = os.path.join(os.path.expanduser("~"), "city_coordinates.json")

PRODUCT_CONFIG = {
    "RAPESEED": {"label": "Rapeseed", "color": "#FFD400", "enabled": True, "aliases": ["RAPESEED", "RAPE"]},
    "SUNFLOWER": {"label": "Sunflower", "color": "#2E8B57", "enabled": True, "aliases": ["SUNFLOWER", "SUN"]}
}

MIN_RADIUS, MAX_RADIUS = 7, 28
CAP_PERCENTILE = 0.95
ROUTE_WEIGHT = 3

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f: return json.load(f)
    return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f: json.dump(cache, f, ensure_ascii=False, indent=4)

def parse_number(value):
    if pd.isna(value): return None
    if isinstance(value, (int, float)): return float(value)
    text = str(value).strip().replace("\xa0", "").replace(" ", "")
    if text == "" or text.upper() == "NAN": return None
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."): text = text.replace(".", "").replace(",", ".")
        else: text = text.replace(",", "")
    else: text = text.replace(",", ".")
    try: return float(text)
    except: return None

def normalize_text(value): return "" if pd.isna(value) else str(value).strip().upper()

def detect_base_key(base_text):
    text = normalize_text(base_text)
    if "KEYWORD_A" in text: return "BASE_A" # Anonymized mapping logic
    if "KEYWORD_B" in text: return "BASE_B"
    return None

def detect_product_group(commodity_desc):
    text = normalize_text(commodity_desc)
    for product_key, config in PRODUCT_CONFIG.items():
        if any(alias in text for alias in config["aliases"]): return product_key
    return None

def format_qty(value): return "N/A" if value is None or pd.isna(value) else f"{value:,.0f} MT"

def get_scale_limits(values, cap_percentile=CAP_PERCENTILE):
    cleaned = [float(v) for v in values if v is not None and not pd.isna(v) and float(v) > 0]
    if not cleaned: return 0, 0
    min_q = min(cleaned)
    max_q = float(pd.Series(cleaned).quantile(cap_percentile))
    return min_q, max(cleaned) if max_q <= min_q else max_q

def scale_radius(qty, min_q, max_q, min_r=MIN_RADIUS, max_r=MAX_RADIUS):
    if qty is None or pd.isna(qty) or qty <= 0: return min_r
    if max_q is None or min_q is None or pd.isna(max_q) or pd.isna(min_q) or max_q <= min_q: return (min_r + max_r) / 2
    capped_qty = min(float(qty), float(max_q))
    sqrt_min, sqrt_max, sqrt_qty = math.sqrt(float(min_q)), math.sqrt(float(max_q)), math.sqrt(capped_qty)
    if sqrt_max == sqrt_min: return (min_r + max_r) / 2
    return min_r + (sqrt_qty - sqrt_min) * (max_r - min_r) / (sqrt_max - sqrt_min)

def make_pie_slice_svg(radius, start_angle, end_angle, color):
    radius = int(max(4, radius))
    size = radius * 2 + 6
    center, r = size / 2, radius
    if end_angle - start_angle >= 359.99:
        return f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}" style="pointer-events:none; overflow:visible;"><circle cx="{center}" cy="{center}" r="{r}" fill="{color}" fill-opacity="0.82" stroke="#ffffff" stroke-width="1.2" style="pointer-events:auto;" /></svg>'
    start_rad, end_rad = math.radians(start_angle - 90), math.radians(end_angle - 90)
    x1, y1 = center + r * math.cos(start_rad), center + r * math.sin(start_rad)
    x2, y2 = center + r * math.cos(end_rad), center + r * math.sin(end_rad)
    large_arc = 1 if (end_angle - start_angle) > 180 else 0
    path = f"M {center},{center} L {x1},{y1} A {r},{r} 0 {large_arc},1 {x2},{y2} Z"
    return f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}" style="pointer-events:none; overflow:visible;"><path d="{path}" fill="{color}" fill-opacity="0.82" stroke="#ffffff" stroke-width="1.2" style="pointer-events:auto;" /></svg>'

def add_pie_slice_marker(map_layer, location, radius, start_angle, end_angle, color, tooltip_html):
    radius = int(max(4, radius))
    size = radius * 2 + 6
    svg = make_pie_slice_svg(radius, start_angle, end_angle, color)
    icon = folium.DivIcon(html=f'<div style="width:{size}px; height:{size}px; pointer-events:none;">{svg}</div>', icon_size=(size, size), icon_anchor=(size / 2, size / 2))
    folium.Marker(location=location, icon=icon, tooltip=folium.Tooltip(tooltip_html, sticky=True)).add_to(map_layer)

def geocode_with_ui(city_name):
    geolocator = Nominatim(user_agent="agri_logistics_mapper")
    try: results = geolocator.geocode(city_name, exactly_one=False, limit=5)
    except: results = None

    if results and len(results) == 1: return results[0].latitude, results[0].longitude

    result_coords = {"lat": None, "lon": None, "resolved": False}
    current_options_data = []

    root = tk.Tk()
    root.title("Geocoding Resolution Required")
    root.attributes("-topmost", True)
    root.geometry("600x350")
    try: root.eval("tk::PlaceWindow . center")
    except: pass

    tk.Label(root, text=f"Multiple/No matches for: {city_name}", font=("Arial", 12, "bold")).pack(pady=10)
    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=20)
    tk.Label(frame, text="Select from results:").pack(anchor=tk.W)

    combo = ttk.Combobox(frame, state="readonly", width=85)

    def update_dropdown(res_list):
        current_options_data.clear()
        if res_list:
            opts = []
            for r in res_list:
                opts.append(r.address)
                current_options_data.append((r.latitude, r.longitude))
            combo["values"] = opts
            combo.current(0)
        else:
            combo.set("No results found. Please refine search.")
            combo["values"] = []

    update_dropdown(results)
    combo.pack(pady=5)
    tk.Label(frame, text="Or enter manually (Zip, City, Country):").pack(anchor=tk.W, pady=(15, 0))
    search_entry = tk.Entry(frame, width=50)
    search_entry.pack(pady=5)

    def on_search():
        if search_entry.get().strip():
            try:
                new_res = geolocator.geocode(search_entry.get().strip(), exactly_one=False, limit=5)
                if new_res: update_dropdown(new_res)
                else: messagebox.showinfo("No Results", "No locations found.")
            except: messagebox.showerror("Error", "Network error during search.")

    tk.Button(frame, text="Search Again", command=on_search).pack()

    def on_save():
        idx = combo.current()
        if 0 <= idx < len(current_options_data):
            result_coords["lat"], result_coords["lon"] = current_options_data[idx]
            result_coords["resolved"] = True
            root.destroy()
        else: messagebox.showwarning("Warning", "Select a valid location!")

    tk.Button(root, text="Save & Continue", command=on_save, bg="lightgreen", font=("Arial", 10, "bold")).pack(pady=20)
    root.mainloop()

    if result_coords["resolved"]: return result_coords["lat"], result_coords["lon"]
    return None, None

def build_summary_panel(df_products):
    summary_rows = []
    for base_key in ["BASE_A", "BASE_B"]:
        base_df = df_products[df_products["BaseKey"] == base_key]
        for product_key in ["RAPESEED", "SUNFLOWER"]:
            sub = base_df[base_df["ProductGroup"] == product_key]
            if sub.empty: continue
            total_qty = sub["Ctr/nom. Qty"].sum()
            delivered_sub, fca_sub = sub[sub["Inco1"].isin(["CPT", "DAP", "DDP"])], sub[sub["Inco1"] == "FCA"]
            summary_rows.append({
                "base": base_key, "product_key": product_key, "product_label": PRODUCT_CONFIG[product_key]["label"],
                "color": PRODUCT_CONFIG[product_key]["color"], "total_qty": total_qty,
                "delivered_qty": delivered_sub["Ctr/nom. Qty"].sum() if not delivered_sub.empty else 0,
                "fca_qty": fca_sub["Ctr/nom. Qty"].sum() if not fca_sub.empty else 0
            })

    grand_total = sum(r["total_qty"] for r in summary_rows)
    html = f"""<div id="summary-panel" style="position: fixed; left: 14px; bottom: 24px; width: 330px; max-height: 55vh; overflow-y: auto; background: white; padding: 10px 12px; border-radius: 12px; box-shadow: 0 4px 18px rgba(0,0,0,0.22); z-index: 9999; font-family: Arial, sans-serif; font-size: 12px;"><details><summary style="font-size:15px; font-weight:bold; cursor:pointer;">Logistics Summary</summary><div style="margin-top:10px; padding:8px; background:#f3f3f3; border-radius:8px;"><b>Total Sales:</b> {format_qty(grand_total)}</div>"""
    
    for base_key in ["BASE_A", "BASE_B"]:
        html += f'<div style="margin-top: 12px; padding: 8px; border-radius: 8px; background: #f6f6f6; font-weight: bold; font-size: 14px;">{escape(base_key)}</div>'
        base_rows = [r for r in summary_rows if r["base"] == base_key]
        if not base_rows:
            html += '<div style="padding:8px; color:#777;">No data available.</div>'
            continue
        for r in base_rows:
            html += f'<div style="margin-top: 8px; padding: 8px; border-left: 6px solid {r["color"]}; background: #ffffff; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.08);"><div style="font-weight:bold; font-size:13px;">{escape(r["product_label"])}</div><div>Total Qty: <b>{format_qty(r["total_qty"])}</b></div><div>Delivered Qty: {format_qty(r["delivered_qty"])}</div><div>FCA Qty: {format_qty(r["fca_qty"])}</div></div>'
    html += "</details></div>"
    return html

def generate_map(excel_path):
    print("\n--- Map Generation Started ---")
    try:
        temp_file = os.path.join(tempfile.gettempdir(), "temp_map_read.xlsx")
        try: shutil.copy2(excel_path, temp_file)
        except: return
        df = pd.read_excel(temp_file, sheet_name=0)
        try: os.remove(temp_file)
        except: pass

        required_cols = ["Base loc name", "Inco1", "Inco2", "Ctr/nom. Qty", "freight EUR", "Commodity desc"]
        if not all(col in df.columns for col in required_cols): return

        df["Base loc name"], df["Inco1"], df["Inco2"], df["Commodity desc"] = df["Base loc name"].apply(normalize_text), df["Inco1"].apply(normalize_text), df["Inco2"].apply(normalize_text), df["Commodity desc"].apply(normalize_text)
        df["Ctr/nom. Qty"], df["freight EUR"] = df["Ctr/nom. Qty"].apply(parse_number), df["freight EUR"].apply(parse_number)
        df = df.dropna(subset=["Base loc name", "Ctr/nom. Qty"])
        df = df[df["Ctr/nom. Qty"] > 0]
        df["BaseKey"], df["ProductGroup"] = df["Base loc name"].apply(detect_base_key), df["Commodity desc"].apply(detect_product_group)
        
        df_products = df[df["ProductGroup"].notna() & df["BaseKey"].notna()].copy()
        if df_products.empty: return

        cache = load_cache()
        m = folium.Map(location=[48.0, 14.0], zoom_start=5, tiles="CartoDB positron")
        base_layer, route_layer = folium.FeatureGroup(name="Bases", show=True), folium.FeatureGroup(name="Routes", show=True)
        product_layers = {pk: folium.FeatureGroup(name=c["label"], show=True) for pk, c in PRODUCT_CONFIG.items() if c["enabled"] and pk in df_products["ProductGroup"].values}

        for base_name, base_coords in BASES.items():
            folium.Marker(location=[base_coords["lat"], base_coords["lon"]], popup=f"Base: {base_name}", icon=folium.Icon(color="darkblue", icon="industry", prefix="fa")).add_to(base_layer)

        delivered_data = df_products[df_products["Inco1"].isin(["CPT", "DDP", "DAP"]) & (df_products["Inco2"] != "NAN") & (df_products["Inco2"] != "")].copy()
        destinations = delivered_data.groupby(["Inco2", "BaseKey"]).agg({"Ctr/nom. Qty": "sum"}).reset_index()
        fca_data_all = df_products[df_products["Inco1"] == "FCA"].copy()
        fca_totals = fca_data_all.groupby(["BaseKey"]).agg({"Ctr/nom. Qty": "sum"}).reset_index()

        scale_values = (destinations["Ctr/nom. Qty"].tolist() if not destinations.empty else []) + (fca_totals["Ctr/nom. Qty"].tolist() if not fca_totals.empty else [])
        min_q, max_q = get_scale_limits(scale_values)

        for _, row in destinations.iterrows():
            city, base_key, total_qty = row["Inco2"], row["BaseKey"], row["Ctr/nom. Qty"]
            if total_qty <= 0 or base_key not in BASES: continue
            
            start_coords = BASES[base_key]
            if city not in cache:
                lat, lon = geocode_with_ui(city)
                if lat is not None and lon is not None: cache[city] = {"lat": lat, "lon": lon}; save_cache(cache)
                else: continue
            dest_coords = cache[city]

            folium.PolyLine(locations=[[start_coords["lat"], start_coords["lon"]], [dest_coords["lat"], dest_coords["lon"]]], color="#808080", weight=ROUTE_WEIGHT, opacity=0.60, tooltip=f"{base_key} → {city} | Total: {format_qty(total_qty)}").add_to(route_layer)
            
            city_base_data = delivered_data[(delivered_data["Inco2"] == city) & (delivered_data["BaseKey"] == base_key)]
            product_split = city_base_data.groupby("ProductGroup").agg({"Ctr/nom. Qty": "sum"}).reset_index()
            radius, start_angle = scale_radius(total_qty, min_q, max_q), 0

            for _, p_row in product_split.iterrows():
                product_key, product_qty = p_row["ProductGroup"], p_row["Ctr/nom. Qty"]
                if product_key not in product_layers: continue
                end_angle = start_angle + (product_qty / total_qty) * 360
                
                comm_text = "<br>".join([f"- {safe_html(k)}: {format_qty(v)}" for k, v in city_base_data[city_base_data["ProductGroup"] == product_key].groupby("Commodity desc")["Ctr/nom. Qty"].sum().items()])
                tooltip_html = f'<div style="font-family: Arial; font-size: 12px;"><b>Destination:</b> {safe_html(city)}<br><b>Product:</b> {safe_html(PRODUCT_CONFIG[product_key]["label"])}<br><b>Qty:</b> {format_qty(product_qty)}<hr style="margin: 3px 0;">{comm_text}</div>'
                
                add_pie_slice_marker(product_layers[product_key], [dest_coords["lat"], dest_coords["lon"]], radius, start_angle, end_angle, PRODUCT_CONFIG[product_key]["color"], tooltip_html)
                start_angle = end_angle

        for base_name, base_coords in BASES.items():
            fca_data = fca_data_all[fca_data_all["BaseKey"] == base_name].copy()
            if fca_data.empty: continue
            total_fca_qty = fca_data["Ctr/nom. Qty"].sum()
            if total_fca_qty <= 0: continue
            
            radius, start_angle = scale_radius(total_fca_qty, min_q, max_q), 0
            for _, p_row in fca_data.groupby("ProductGroup").agg({"Ctr/nom. Qty": "sum"}).reset_index().iterrows():
                product_key, product_qty = p_row["ProductGroup"], p_row["Ctr/nom. Qty"]
                if product_key not in product_layers: continue
                end_angle = start_angle + (product_qty / total_fca_qty) * 360
                
                customer_text = "<br>".join([f"- {safe_html(c if c and c!='NAN' else 'N/A')} | {safe_html(comm)}: {format_qty(q)}" for (c, comm), q in fca_data[fca_data["ProductGroup"] == product_key].groupby(["Inco2", "Commodity desc"])["Ctr/nom. Qty"].sum().items()])
                tooltip_html = f'<div style="font-family: Arial; font-size: 12px;"><b>FCA Base:</b> {safe_html(base_name)}<br><b>Product:</b> {safe_html(PRODUCT_CONFIG[product_key]["label"])}<br><b>Qty:</b> {format_qty(product_qty)}<hr style="margin: 3px 0;">{customer_text}</div>'
                
                add_pie_slice_marker(product_layers[product_key], [base_coords["lat"], base_coords["lon"]], radius, start_angle, end_angle, PRODUCT_CONFIG[product_key]["color"], tooltip_html)
                start_angle = end_angle

        m.get_root().html.add_child(folium.Element(build_summary_panel(df_products)))
        base_layer.add_to(m); route_layer.add_to(m)
        for _, layer in product_layers.items(): layer.add_to(m)
        folium.LayerControl(collapsed=False, position="topright").add_to(m)

        map_filename = os.path.join(os.path.dirname(excel_path), "Logistics_Map.html")
        m.save(map_filename)
        save_cache(cache)
        print(f"--- Map generated successfully: {map_filename} ---")
        webbrowser.open(map_filename)

    except Exception as e: print(f"Map generation error: {e}")

if __name__ == "__main__":
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    initial_dir = os.path.expanduser(r"~\Documents")
    selected_file = filedialog.askopenfilename(title="Select Report for Mapping", initialdir=initial_dir, filetypes=[("Excel files", "*.xlsx;*.xls")])
    root.destroy()
    if selected_file: generate_map(selected_file)