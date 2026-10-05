"""
Logistics Routing Map Generator
-------------------------------
Generates an interactive HTML map visualizing sales and delivery volumes.
Draws routes from fulfillment bases to destinations and places dynamic 
pie-chart markers to represent commodity breakdowns.

Features mutually exclusive layer filtering via JavaScript injection.

Uses Folium (Leaflet.js) and Geopy (Nominatim API).
Author: Nándor Magyar
Disclaimer: Internal location names and paths have been anonymized.
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

# ==============================================================================
# BASE LOCATIONS & CACHE CONFIGURATION (Anonymized)
# ==============================================================================

BASES = {
    "BASE_A": {"lat": 47.0000, "lon": 20.2833}, # Example coordinates
    "BASE_B": {"lat": 48.0245, "lon": 16.7784}  # Example coordinates
}

CACHE_FILE = os.path.join(os.path.expanduser("~"), "city_coordinates.json")

# ==============================================================================
# PRODUCT CONFIGURATION
# ==============================================================================

PRODUCT_CONFIG = {
    "RAPESEED": {
        "label": "RSM",
        "color": "#FFD400",
        "enabled": True,
        "aliases": ["RAPESEED", "RAPE"]
    },
    "SUNFLOWER": {
        "label": "SFM",
        "color": "#2E8B57",
        "enabled": True,
        "aliases": ["SUNFLOWER", "SUNFLOWERSEED", "SUN FLOWER", "SFM"]
    },
    "SOY": {
        "label": "SOY",
        "color": "#D62728",
        "enabled": False,
        "aliases": ["ANY SOYABEANMEAL FEED, DEHULLED"]
    }
}

# ==============================================================================
# VISUALIZATION SETTINGS
# ==============================================================================

MIN_RADIUS = 7
MAX_RADIUS = 22
CAP_PERCENTILE = 0.95
ROUTE_WEIGHT = 3

SHOW_BASE_ICONS = True

BASE_ICON_OFFSETS = {
    "BASE_A": {"lat": 0.10, "lon": -0.12},
    "BASE_B": {"lat": 0.10, "lon": 0.12}
}

PARTNER_COLUMN_CANDIDATES = [
    "Sold-to party name", "Sold-to pt name", "Sold-to name", "Customer name",
    "Customer", "Partner", "Partner name", "Name 1", "Sold-to party",
    "Ship-to party name", "Sold-to party description", "Ship-to party description"
]

# ==============================================================================
# CACHE MANAGEMENT
# ==============================================================================

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=4)

# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def parse_number(value):
    if pd.isna(value): return None
    if isinstance(value, (int, float)): return float(value)
    text = str(value).strip().replace("\xa0", "").replace(" ", "")
    if text == "" or text.upper() == "NAN": return None
    if "." in text and "," in text:
        if text.rfind(",") > text.rfind("."): text = text.replace(".", "").replace(",", ".")
        else: text = text.replace(",", "")
    else:
        text = text.replace(",", ".")
    try: return float(text)
    except ValueError: return None

def normalize_text(value):
    if pd.isna(value): return ""
    return str(value).strip().upper()

def safe_html(value):
    return escape(str(value))

def format_qty(value):
    if value is None or pd.isna(value): return "N/A"
    return f"{value:,.0f} MT"

def format_price(value):
    if value is None or pd.isna(value): return "N/A"
    return f"{value:,.2f} €/MT"

def detect_base_key(base_text):
    text = normalize_text(base_text)
    if "KEYWORD_A" in text: return "BASE_A" # Anonymized logic
    if "KEYWORD_B" in text: return "BASE_B" # Anonymized logic
    return None

def detect_product_group(commodity_desc):
    text = normalize_text(commodity_desc)
    for product_key, config in PRODUCT_CONFIG.items():
        for alias in config["aliases"]:
            if alias.upper() in text: return product_key
    return None

def product_enabled(product_key):
    return product_key in PRODUCT_CONFIG and PRODUCT_CONFIG[product_key].get("enabled", False)

def product_label(product_key):
    return PRODUCT_CONFIG[product_key]["label"]

def product_color(product_key):
    return PRODUCT_CONFIG[product_key]["color"]

def weighted_average(group, value_col, weight_col):
    valid = group[[value_col, weight_col]].copy()
    valid[value_col] = pd.to_numeric(valid[value_col], errors="coerce")
    valid[weight_col] = pd.to_numeric(valid[weight_col], errors="coerce")
    valid = valid.dropna(subset=[value_col, weight_col])
    valid = valid[valid[weight_col] > 0]
    if valid.empty: return None
    total_weight = valid[weight_col].sum()
    if total_weight == 0: return None
    return (valid[value_col] * valid[weight_col]).sum() / total_weight

def get_scale_limits(values, cap_percentile=CAP_PERCENTILE):
    cleaned = [float(v) for v in values if v is not None and not pd.isna(v) and float(v) > 0]
    if not cleaned: return 0, 0
    min_q = min(cleaned)
    max_q = float(pd.Series(cleaned).quantile(cap_percentile))
    if max_q <= min_q: max_q = max(cleaned)
    return min_q, max_q

def scale_radius(qty, min_q, max_q, min_r=MIN_RADIUS, max_r=MAX_RADIUS):
    if qty is None or pd.isna(qty) or qty <= 0: return min_r
    if max_q is None or min_q is None or pd.isna(max_q) or pd.isna(min_q) or max_q <= min_q: return (min_r + max_r) / 2
    capped_qty = min(float(qty), float(max_q))
    log_min, log_max, log_qty = math.log(float(min_q) + 1), math.log(float(max_q) + 1), math.log(capped_qty + 1)
    if log_max == log_min: return (min_r + max_r) / 2
    return min_r + (log_qty - log_min) * (max_r - min_r) / (log_max - log_min)

def find_partner_column(df):
    normalized_map = {str(c).strip().upper(): c for c in df.columns}
    for candidate in PARTNER_COLUMN_CANDIDATES:
        key = candidate.strip().upper()
        if key in normalized_map: return normalized_map[key]
    return None

def get_partner_text(group, fallback_value):
    partner_col = group.attrs.get("partner_col")
    if partner_col and partner_col in group.columns:
        vals = []
        for v in group[partner_col].dropna().astype(str).tolist():
            clean = v.strip()
            if clean and clean.upper() != "NAN" and clean not in vals: vals.append(clean)
        if vals: return ", ".join(vals[:5])
    fallback = str(fallback_value).strip()
    if fallback and fallback.upper() != "NAN": return fallback
    return "N/A"

def html_tooltip(lines):
    return "<div style='font-family: Arial; font-size: 12px; line-height: 1.35;'>" + "<br>".join(lines) + "</div>"

def plain_from_html_lines(lines):
    clean_lines = []
    for line in lines:
        clean = str(line).replace("<b>", "").replace("</b>", "").replace("<br>", "\n")
        clean_lines.append(clean)
    return "\n".join(clean_lines)

# ==============================================================================
# CUSTOM SVG PIE CHART MARKER
# ==============================================================================

def polar_to_cartesian(center, radius, angle_degrees):
    angle_radians = math.radians(angle_degrees - 90)
    return (center + radius * math.cos(angle_radians), center + radius * math.sin(angle_radians))

def make_svg_pie(radius, slices, center_label=""):
    radius = int(max(4, radius))
    size = radius * 2 + 8
    center = size / 2

    if not slices: return ""
    if len(slices) == 1:
        s = slices[0]
        color, title = product_color(s["product_key"]), safe_html(s["tooltip_plain"])
        return f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}" xmlns="http://www.w3.org/2000/svg" style="overflow:visible;"><circle cx="{center}" cy="{center}" r="{radius}" fill="{color}" fill-opacity="0.84" stroke="#ffffff" stroke-width="1.4"><title>{title}</title></circle></svg>'

    svg_parts = [f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}" xmlns="http://www.w3.org/2000/svg" style="overflow:visible;">']
    start_angle = 0.0

    for idx, s in enumerate(slices):
        if s["total_qty"] <= 0 or s["qty"] <= 0: continue
        slice_angle = (s["qty"] / s["total_qty"]) * 360.0
        end_angle = 360.0 if idx == len(slices) - 1 else start_angle + slice_angle
        x1, y1 = polar_to_cartesian(center, radius, start_angle)
        x2, y2 = polar_to_cartesian(center, radius, end_angle)
        large_arc = 1 if (end_angle - start_angle) > 180 else 0
        path = f"M {center},{center} L {x1},{y1} A {radius},{radius} 0 {large_arc},1 {x2},{y2} Z"
        color, title = product_color(s["product_key"]), safe_html(s["tooltip_plain"])
        svg_parts.append(f'<path d="{path}" fill="{color}" fill-opacity="0.84" stroke="#ffffff" stroke-width="1.4"><title>{title}</title></path>')
        start_angle = end_angle

    svg_parts.append("</svg>")
    return "".join(svg_parts)

def add_svg_marker(layer, location, svg_html, radius):
    radius = int(max(4, radius))
    size = radius * 2 + 8
    icon = folium.DivIcon(html=f'<div style="width:{size}px; height:{size}px;">{svg_html}</div>', icon_size=(size, size), icon_anchor=(size / 2, size / 2))
    folium.Marker(location=location, icon=icon).add_to(layer)

# ==============================================================================
# GEOCODING RESOLUTION UI
# ==============================================================================

def geocode_with_ui(city_name):
    """Fetches coordinates. If multiple results are found, prompts the user to select."""
    geolocator = Nominatim(user_agent="logistics_mapper_portfolio")
    try: results = geolocator.geocode(city_name, exactly_one=False, limit=5)
    except Exception as e: print(f"    [!] Search error ({city_name}): {e}"); results = None

    if results and len(results) == 1:
        print(f"    [+] {city_name} auto-resolved: {results[0].address}")
        return results[0].latitude, results[0].longitude

    result_coords = {"lat": None, "lon": None, "resolved": False}
    current_options_data = []

    root = tk.Tk()
    root.title("Geocoding Resolution")
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
        new_query = search_entry.get().strip()
        if new_query:
            try:
                new_res = geolocator.geocode(new_query, exactly_one=False, limit=5)
                if new_res: update_dropdown(new_res)
                else: messagebox.showinfo("No match", "Could not find location.")
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

# ==============================================================================
# TOOLTIP / TEXT BUILDERS
# ==============================================================================

def build_delivery_product_tooltip(city, base_key, product_key, product_qty, total_qty, product_data):
    weighted_freight = weighted_average(product_data, "freight EUR", "Ctr/nom. Qty")
    partner = get_partner_text(product_data, city)
    lines = [
        f"<b>Partner:</b> {safe_html(partner)}",
        f"<b>Destination:</b> {safe_html(city)}",
        f"<b>Product:</b> {safe_html(product_label(product_key))}",
        f"<b>Total quantity:</b> {format_qty(product_qty)}",
        f"<b>Freight:</b> {format_price(weighted_freight)}"
    ]
    return html_tooltip(lines), plain_from_html_lines(lines)

def build_fca_product_tooltip(base_key, product_key, product_qty, total_qty, product_data):
    weighted_freight = weighted_average(product_data, "freight EUR", "Ctr/nom. Qty")
    partner = get_partner_text(product_data, "N/A")
    lines = [
        f"<b>Partner:</b> {safe_html(partner)}",
        f"<b>Destination:</b> FCA {safe_html(base_key)}",
        f"<b>Product:</b> {safe_html(product_label(product_key))}",
        f"<b>Total quantity:</b> {format_qty(product_qty)}",
        f"<b>Freight:</b> {format_price(weighted_freight)}"
    ]
    return html_tooltip(lines), plain_from_html_lines(lines)

# ==============================================================================
# JAVASCRIPT INJECTION FOR MUTUALLY EXCLUSIVE LAYERS
# ==============================================================================

def add_mutually_exclusive_layer_script(m, all_layer, product_layers):
    map_name = m.get_name()
    all_layer_var = all_layer.get_name()

    only_layer_items = [{"label": f"{PRODUCT_CONFIG[pk]['label']} only", "var_name": l.get_name()} for pk, l in product_layers.items()]
    only_layers_js = ",\n            ".join([f'{{ label: "{i["label"]}", layer: {i["var_name"]} }}' for i in only_layer_items])

    script = f"""
    setTimeout(function() {{
        var map = {map_name};
        var allLayer = {all_layer_var};
        var onlyLayers = [{only_layers_js}];

        function normalizeText(text) {{ return (text || '').replace(/\\s+/g, ' ').trim(); }}

        function getOverlayRows() {{
            var rows = [];
            var labels = document.querySelectorAll('.leaflet-control-layers-overlays label');
            labels.forEach(function(label) {{
                var input = label.querySelector("input[type='checkbox']");
                var text = normalizeText(label.textContent);
                if (input) rows.push({{labelElement: label, inputElement: input, text: text}});
            }});
            return rows;
        }}

        function setCheckbox(labelText, checked) {{
            getOverlayRows().forEach(function(row) {{ if (row.text === labelText) row.inputElement.checked = checked; }});
        }}

        function syncCheckboxesAll() {{
            setCheckbox('All Sales Mix', true);
            onlyLayers.forEach(function(item) {{ setCheckbox(item.label, false); }});
        }}

        function syncCheckboxesOnly(activeLabel) {{
            setCheckbox('All Sales Mix', false);
            onlyLayers.forEach(function(item) {{ setCheckbox(item.label, item.label === activeLabel); }});
        }}

        function removeLayerIfVisible(layer) {{ if (map.hasLayer(layer)) map.removeLayer(layer); }}
        function addLayerIfHidden(layer) {{ if (!map.hasLayer(layer)) map.addLayer(layer); }}

        function activateAll() {{
            onlyLayers.forEach(function(item) {{ removeLayerIfVisible(item.layer); }});
            addLayerIfHidden(allLayer);
            syncCheckboxesAll();
        }}

        function activateOnly(activeItem) {{
            removeLayerIfVisible(allLayer);
            onlyLayers.forEach(function(item) {{ if (item.label !== activeItem.label) removeLayerIfVisible(item.layer); }});
            addLayerIfHidden(activeItem.layer);
            syncCheckboxesOnly(activeItem.label);
        }}

        function wireCheckboxes() {{
            getOverlayRows().forEach(function(row) {{
                if (row.text === 'All Sales Mix') {{
                    row.inputElement.addEventListener('click', function(e) {{ e.preventDefault(); e.stopPropagation(); activateAll(); return false; }}, true);
                }}
                onlyLayers.forEach(function(item) {{
                    if (row.text === item.label) {{
                        row.inputElement.addEventListener('click', function(e) {{ e.preventDefault(); e.stopPropagation(); activateOnly(item); return false; }}, true);
                    }}
                }});
            }});
        }}
        wireCheckboxes();
        activateAll();
    }}, 1000);
    """
    m.get_root().script.add_child(folium.Element(script))

# ==============================================================================
# MAIN MAP GENERATION
# ==============================================================================

def generate_map(excel_path):
    print("\n--- Generating Map... ---")
    try:
        temp_file = os.path.join(tempfile.gettempdir(), "temp_map_read.xlsx")
        try: shutil.copy2(excel_path, temp_file)
        except Exception as e: print(f"[ERROR] Copy failed: {e}"); return
        
        df = pd.read_excel(temp_file, sheet_name=0)
        try: os.remove(temp_file)
        except: pass

        required_cols = ["Base loc name", "Inco1", "Inco2", "Ctr/nom. Qty", "freight EUR", "Commodity desc"]
        for col in required_cols:
            if col not in df.columns:
                print(f"[ERROR] Missing column '{col}'. Map cannot be generated.")
                return

        partner_col = find_partner_column(df)

        df["Base loc name"], df["Inco1"], df["Inco2"], df["Commodity desc"] = df["Base loc name"].apply(normalize_text), df["Inco1"].apply(normalize_text), df["Inco2"].apply(normalize_text), df["Commodity desc"].apply(normalize_text)
        df["Ctr/nom. Qty"], df["freight EUR"] = df["Ctr/nom. Qty"].apply(parse_number), df["freight EUR"].apply(parse_number)
        
        df = df.dropna(subset=["Base loc name", "Ctr/nom. Qty"])
        df = df[df["Ctr/nom. Qty"] > 0]
        df["BaseKey"], df["ProductGroup"] = df["Base loc name"].apply(detect_base_key), df["Commodity desc"].apply(detect_product_group)

        df_products = df[df["ProductGroup"].notna() & df["ProductGroup"].apply(product_enabled) & df["BaseKey"].notna()].copy()
        df_products.attrs["partner_col"] = partner_col

        if df_products.empty:
            print("[INFO] No valid product data found.")
            return

        cache = load_cache()
        m = folium.Map(location=[48.0, 14.0], zoom_start=5, tiles="CartoDB positron")

        base_layer = folium.FeatureGroup(name="Bases", show=True)
        all_sales_layer = folium.FeatureGroup(name="All Sales Mix", show=True)
        product_layers = {pk: folium.FeatureGroup(name=f"{config['label']} only", show=False) for pk, config in PRODUCT_CONFIG.items() if config["enabled"] and pk in df_products["ProductGroup"].values}

        delivered_data = df_products[df_products["Inco1"].isin(["CPT", "DDP", "DAP"]) & (df_products["Inco2"] != "NAN") & (df_products["Inco2"] != "")].copy()
        delivered_data.attrs["partner_col"] = partner_col
        destinations = delivered_data.groupby(["Inco2", "BaseKey"]).agg({"Ctr/nom. Qty": "sum"}).reset_index()

        fca_data_all = df_products[df_products["Inco1"] == "FCA"].copy()
        fca_data_all.attrs["partner_col"] = partner_col
        fca_totals = fca_data_all.groupby(["BaseKey"]).agg({"Ctr/nom. Qty": "sum"}).reset_index()

        if SHOW_BASE_ICONS:
            for base_name, base_coords in BASES.items():
                has_fca = not fca_totals.empty and not fca_totals[fca_totals["BaseKey"] == base_name].empty and fca_totals[fca_totals["BaseKey"] == base_name]["Ctr/nom. Qty"].iloc[0] > 0
                offset = BASE_ICON_OFFSETS.get(base_name, {"lat": 0, "lon": 0}) if has_fca else {"lat": 0, "lon": 0}
                icon_location = [base_coords["lat"] + offset["lat"], base_coords["lon"] + offset["lon"]]
                if has_fca: folium.PolyLine(locations=[[base_coords["lat"], base_coords["lon"]], icon_location], color="#1f4e79", weight=1, opacity=0.45, dash_array="4,4").add_to(base_layer)
                folium.Marker(location=icon_location, popup=f"Facility: {base_name}", tooltip=f"Facility: {base_name}", icon=folium.Icon(color="darkblue", icon="industry", prefix="fa")).add_to(base_layer)

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

            city_base_data = delivered_data[(delivered_data["Inco2"] == city) & (delivered_data["BaseKey"] == base_key)].copy()
            city_base_data.attrs["partner_col"] = partner_col
            
            slices = []
            for _, p_row in city_base_data.groupby("ProductGroup").agg({"Ctr/nom. Qty": "sum"}).reset_index().iterrows():
                if p_row["ProductGroup"] not in product_layers: continue
                p_data = city_base_data[city_base_data["ProductGroup"] == p_row["ProductGroup"]].copy()
                p_data.attrs["partner_col"] = partner_col
                _, tooltip_plain = build_delivery_product_tooltip(city, base_key, p_row["ProductGroup"], p_row["Ctr/nom. Qty"], total_qty, p_data)
                slices.append({"product_key": p_row["ProductGroup"], "qty": p_row["Ctr/nom. Qty"], "total_qty": total_qty, "tooltip_plain": tooltip_plain})

            if slices:
                folium.PolyLine(locations=[[start_coords["lat"], start_coords["lon"]], [dest_coords["lat"], dest_coords["lon"]]], color="#808080", weight=ROUTE_WEIGHT, opacity=0.60, tooltip=f"{base_key} → {city} | Total: {format_qty(total_qty)}").add_to(all_sales_layer)
                add_svg_marker(all_sales_layer, [dest_coords["lat"], dest_coords["lon"]], make_svg_pie(scale_radius(total_qty, min_q, max_q), slices, f"{city} | Total: {format_qty(total_qty)}"), scale_radius(total_qty, min_q, max_q))

            for _, p_row in city_base_data.groupby("ProductGroup").agg({"Ctr/nom. Qty": "sum"}).reset_index().iterrows():
                pk = p_row["ProductGroup"]
                if pk not in product_layers: continue
                p_data = city_base_data[city_base_data["ProductGroup"] == pk].copy()
                p_data.attrs["partner_col"] = partner_col
                tooltip_html, _ = build_delivery_product_tooltip(city, base_key, pk, p_row["Ctr/nom. Qty"], total_qty, p_data)
                folium.PolyLine(locations=[[start_coords["lat"], start_coords["lon"]], [dest_coords["lat"], dest_coords["lon"]]], color=product_color(pk), weight=ROUTE_WEIGHT, opacity=0.55, tooltip=f"{product_label(pk)} | {base_key} → {city} | {format_qty(p_row['Ctr/nom. Qty'])}").add_to(product_layers[pk])
                folium.CircleMarker(location=[dest_coords["lat"], dest_coords["lon"]], radius=scale_radius(p_row["Ctr/nom. Qty"], min_q, max_q), color="#ffffff", weight=1.2, fill=True, fill_color=product_color(pk), fill_opacity=0.82, tooltip=folium.Tooltip(tooltip_html, sticky=True)).add_to(product_layers[pk])

        for base_name, base_coords in BASES.items():
            fca_data = fca_data_all[fca_data_all["BaseKey"] == base_name].copy()
            fca_data.attrs["partner_col"] = partner_col
            if fca_data.empty or fca_data["Ctr/nom. Qty"].sum() <= 0: continue
            
            total_fca_qty = fca_data["Ctr/nom. Qty"].sum()
            slices = []
            for _, p_row in fca_data.groupby("ProductGroup").agg({"Ctr/nom. Qty": "sum"}).reset_index().iterrows():
                if p_row["ProductGroup"] not in product_layers: continue
                p_data = fca_data[fca_data["ProductGroup"] == p_row["ProductGroup"]].copy()
                p_data.attrs["partner_col"] = partner_col
                _, tooltip_plain = build_fca_product_tooltip(base_name, p_row["ProductGroup"], p_row["Ctr/nom. Qty"], total_fca_qty, p_data)
                slices.append({"product_key": p_row["ProductGroup"], "qty": p_row["Ctr/nom. Qty"], "total_qty": total_fca_qty, "tooltip_plain": tooltip_plain})

            if slices: add_svg_marker(all_sales_layer, [base_coords["lat"], base_coords["lon"]], make_svg_pie(scale_radius(total_fca_qty, min_q, max_q), slices, f"FCA {base_name} | Total: {format_qty(total_fca_qty)}"), scale_radius(total_fca_qty, min_q, max_q))

            for _, p_row in fca_data.groupby("ProductGroup").agg({"Ctr/nom. Qty": "sum"}).reset_index().iterrows():
                pk = p_row["ProductGroup"]
                if pk not in product_layers: continue
                p_data = fca_data[fca_data["ProductGroup"] == pk].copy()
                p_data.attrs["partner_col"] = partner_col
                tooltip_html, _ = build_fca_product_tooltip(base_name, pk, p_row["Ctr/nom. Qty"], total_fca_qty, p_data)
                folium.CircleMarker(location=[base_coords["lat"], base_coords["lon"]], radius=scale_radius(p_row["Ctr/nom. Qty"], min_q, max_q), color="#ffffff", weight=1.2, fill=True, fill_color=product_color(pk), fill_opacity=0.82, tooltip=folium.Tooltip(tooltip_html, sticky=True)).add_to(product_layers[pk])

        if SHOW_BASE_ICONS: base_layer.add_to(m)
        all_sales_layer.add_to(m)
        for _, layer in product_layers.items(): layer.add_to(m)
        folium.LayerControl(collapsed=False, position="topright").add_to(m)
        add_mutually_exclusive_layer_script(m, all_sales_layer, product_layers)

        map_filename = os.path.join(os.path.dirname(excel_path), "Logistics_Map.html")
        m.save(map_filename)
        save_cache(cache)
        print(f"--- Map generated: {map_filename} ---")
        webbrowser.open(map_filename)
    except Exception as e: print(f"Map Error: {e}")

if __name__ == "__main__":
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    initial_dir = os.path.expanduser(r"~\Documents\Logistics_Reports") # Anonymized path
    if not os.path.exists(initial_dir): initial_dir = os.path.expanduser("~")
    selected_file = filedialog.askopenfilename(title="Select Excel File for Map", initialdir=initial_dir, filetypes=[("Excel files", "*.xlsx;*.xls")])
    root.destroy()
    if selected_file: generate_map(selected_file)
