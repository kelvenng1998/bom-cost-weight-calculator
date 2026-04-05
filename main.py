import os
import re
import unicodedata
import pandas as pd
from tkinter import Tk, filedialog, messagebox
import shutil
import time
import subprocess
import platform

# ---------------------------------------------------------
# FIXED PATH CONFIGURATION
# ---------------------------------------------------------
# This finds the folder where this script is currently saved, let's say C:\Users\Ng Kel Ven\Desktop\BOM
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 1. Source: Defaults to the script's folder, which is C:\Users\Ng Kel Ven\Desktop\BOM
DEFAULT_SOURCE_FOLDER = BASE_DIR

# 2. Output of the calculation: Fixed to \BOM\output
OUTPUT_BASE_FOLDER = os.path.join(BASE_DIR, "output")

# 3. Database for the raw material calculation: Fixed to \BOM\sample_data\raw_material_data.xlsx
RAW_MATERIAL_FILE = os.path.join(BASE_DIR, "sample_data", "raw_material_data.xlsx")

# 4. Memory of the last folder: Saves the last used folder location
LAST_FOLDER_FILE = os.path.join(BASE_DIR, "last_source_folder.txt")

def reset_and_restart():
    """Deletes memory and restarts the script."""
    if os.path.exists(LAST_FOLDER_FILE):
        os.remove(LAST_FOLDER_FILE)
    messagebox.showinfo("Reset", "Folder memory cleared. Restarting...")
    os.execl(sys.executable, sys.executable, *sys.argv)

# ---------------------------------------------------------
# FOLDER LOGIC (STICKY SOURCE)
# ---------------------------------------------------------
# If there is a .txt file, get from this as the last save file
SAMPLE_DATA_FOLDER = os.path.join(BASE_DIR, "sample_data")
SOURCE_FOLDER = SAMPLE_DATA_FOLDER
    
root = Tk()
root.withdraw()

# ---------------------------------------------------------
# 🧼 Normalize
# ---------------------------------------------------------
# This is to make all the text readable, including blank cells which is call nAN
def normalize(text):
    if pd.isna(text): return ""
    text = str(text)
    text = unicodedata.normalize('NFKD', text)
    text = text.strip().upper()
    text = re.sub(r'\s+', ' ', text)
    text = text.replace("°", "DEG").replace("–", "-")
    return text

# ---------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------
print(f"📍 Current Source: {SOURCE_FOLDER}")
print("💡 Processing BOM data...")

# ---------------------------------------------------------
# 🛡️ Safe filename
# ---------------------------------------------------------
def safe_filename(text):
    return re.sub(r'[<>:"/\\|?*]', '_', str(text))

# ---------------------------------------------------------
# Load list of input BOM files
# ---------------------------------------------------------
def load_file_list(file_path):
    df = pd.read_excel(file_path)
    required = {'Filename', 'Quantity'}
    if not required.issubset(df.columns):
        raise ValueError("Excel must contain 'Filename' and 'Quantity' columns.")
    return df

# ---------------------------------------------------------
# Tkinter BOM File Selection (always open at IOUT)
# ---------------------------------------------------------
root = Tk()
root.withdraw()
file_list_excel = os.path.join(BASE_DIR, "sample_data", "input.xlsx")

if not file_list_excel:
    raise FileNotFoundError("❌ No BOM list selected.")

file_list_df = load_file_list(file_list_excel)
project_name = os.path.splitext(os.path.basename(file_list_excel))[0]

# Force the output to be a subfolder inside the IOUT directory
project_name_path = os.path.join(OUTPUT_BASE_FOLDER, project_name)
print(f"📁 Output folder: {project_name_path}")
if not os.path.exists(project_name_path):
    os.makedirs(project_name_path)
    print(f"✅ Output folder auto-created: {project_name_path}")

# ---------------------------------------------------------
# 🛑 Check if any Excel file is open before clearing
# ---------------------------------------------------------
def normalize_input_filename(value):
    name = str(value).strip()
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name

def check_and_wait_for_open_files(folder_path):
    while True:
        open_files = []

        for file in os.listdir(folder_path):
            if file.endswith(".xlsx"):
                file_path = os.path.join(folder_path, file)
                try:
                    with open(file_path, "a"):
                        pass
                except PermissionError:
                    open_files.append(file)

        if not open_files:
            return True  # Safe to continue

        msg = "The following Excel files are currently open:\n\n"
        msg += "\n".join(open_files)
        msg += "\n\nPlease close them.\n\nClick RETRY after closing.\nClick CANCEL to stop the program."

        retry = messagebox.askretrycancel("Excel File Open", msg)

        if not retry:
            messagebox.showinfo("Operation Cancelled", "Process stopped by user.")
            return False

        time.sleep(1)

if os.path.exists(project_name_path):

    # Ask user to close open Excel files
    if not check_and_wait_for_open_files(project_name_path):
        raise SystemExit("❌ User cancelled due to open Excel files.")

    # Now safe to delete files
    for item in os.listdir(project_name_path):
        item_path = os.path.join(project_name_path, item)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.unlink(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
        except PermissionError:
            messagebox.showerror(
                "File Locked",
                f"Cannot delete file:\n{item}\n\nPlease close it and run again."
            )
            raise

    print(f"🧹 Output folder cleared: {project_name_path}")

else:
    os.makedirs(project_name_path)
    print(f"✅ Output folder created: {project_name_path}")

# ---------------------------------------------------------
# Load raw material data (fixed location)
# ---------------------------------------------------------
if not os.path.exists(RAW_MATERIAL_FILE):
    raise FileNotFoundError("❌ Raw material file not found at fixed location.")

material_df = pd.read_excel(RAW_MATERIAL_FILE)

required_columns = {'Type', 'Specification', 'Unit', 'Unit Cost', 'Unit Weight'}
if not required_columns.issubset(material_df.columns):
    raise ValueError(f"Material file must have columns: {required_columns}")

material_df['__Type_Norm'] = material_df['Type'].apply(normalize)
material_df['__Spec_Norm'] = material_df['Specification'].apply(normalize)

missing_files = set()
missing_weighting = []

# ---------------------------------------------------------
# Constants
# ---------------------------------------------------------
ignore_keywords = 'EXPANDED|MESH'
fitting_keywords = "NUT|FLANGE|ELBOW|REDUCER|WASHER"
DEFAULT_BAR_LENGTH = 6000
KERF = 3  # mm per cut
STUD_BOLT_BAR_LENGTH = 2000

# ---------------------------------------------------------
# BAR NESTING
# ---------------------------------------------------------
def bar_nesting():
    aggregated = {}

    for _, file_row in file_list_df.iterrows():
        file_name = normalize_input_filename(file_row['Filename'])
        if not file_name.lower().endswith(".xlsx"):
            file_name += ".xlsx"
        multiplier = int(file_row['Quantity'])

        file_path = os.path.join(SOURCE_FOLDER, "database", file_name)
        if not os.path.exists(file_path):
            print(f"⚠️ File not found: {file_name}")
            missing_files.add(file_name)
            continue

        df = pd.read_excel(file_path)

        df['Length'] = pd.to_numeric(df['Length'], errors='coerce')
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')

        df = df.dropna(subset=['Type', 'Specification', 'Length', 'Quantity'])
        df = df[(df['Length'] > 0) & (df['Quantity'] > 0)]
        df['Quantity'] = df['Quantity'].astype(int) * multiplier

        df = df[
            ~df['Type'].str.upper().str.contains(ignore_keywords + "|" + fitting_keywords, na=False) &
            ~df['Specification'].str.upper().str.contains(ignore_keywords + "|" + fitting_keywords, na=False) &
            ~df['Type'].str.upper().str.contains("FLAT PLATE", na=False) &
            ~df['Specification'].str.upper().str.contains("FLAT PLATE", na=False)
        ]

        grouped = df.groupby(['Type', 'Specification'])
        for (bar_type, spec), df_group in grouped:
            key = (str(bar_type).strip(), str(spec).strip())
            if key not in aggregated:
                aggregated[key] = {'cuts': [], 'bar_type': bar_type, 'spec': spec, 'total_pieces': 0}
            for _, row in df_group.iterrows():
                length = int(row['Length'])
                qty = int(row['Quantity'])
                aggregated[key]['cuts'].extend([length] * qty)
                aggregated[key]['total_pieces'] += qty

    # Nesting
    for (bar_type, spec), info in aggregated.items():
        cuts = info['cuts']
        if not cuts:
            continue

        cuts.sort(reverse=True)
        bar_length = STUD_BOLT_BAR_LENGTH if "STUD BOLT" in str(bar_type).upper() else DEFAULT_BAR_LENGTH
        bars = []

        for cut in cuts:
            cut_with_kerf = cut + KERF
            
            best_bar = None
            min_waste = bar_length + 1

            for bar in bars:
                used = sum(bar)
                remaining = bar_length - used

                if cut_with_kerf <= remaining:
                    waste = remaining - cut_with_kerf
                    if waste < min_waste:
                        min_waste = waste
                        best_bar = bar

            if best_bar is not None:
                best_bar.append(cut_with_kerf)
            else:
                bars.append([cut_with_kerf])

        total_used = 0
        total_waste = 0

        rows = []
        for i, bar in enumerate(bars, start=1):
            used = sum(bar)
            waste = bar_length - used

            total_used += used
            total_waste += waste

            rows.append({
                "Bar #": i,
                "Cuts": ", ".join(str(c - KERF) for c in bar),
                "Used (mm)": used,
                "Waste (mm)": waste
            })

        df_summary = pd.DataFrame(rows)

        summary_row = pd.DataFrame([{
            "Bar #": "TOTAL",
            "Cuts": "",
            "Used (mm)": total_used,
            "Waste (mm)": total_waste
        }])

        df_summary = pd.concat([df_summary, summary_row], ignore_index=True)

        safe_type = safe_filename(bar_type)
        safe_spec = safe_filename(spec)
        out_file = f"{safe_type}_{safe_spec}_{info['total_pieces']} pcs_{len(bars)} bars.xlsx"
        df_summary.to_excel(os.path.join(project_name_path, out_file), index=False)

        print(f"✅ {safe_type} - {safe_spec}: {info['total_pieces']} pcs → {len(bars)} bars")

# ---------------------------------------------------------
# WEIGHT SUMMARY
# ---------------------------------------------------------
def weight_summary():
    results = []

    for _, file_row in file_list_df.iterrows():
        file_name = normalize_input_filename(file_row['Filename'])
        multiplier = int(file_row['Quantity'])

        file_path = os.path.join(SOURCE_FOLDER, "database",file_name)

        if not os.path.exists(file_path):
            missing_files.add(file_name)
            continue

        df = pd.read_excel(file_path)

        df['Length'] = pd.to_numeric(df['Length'], errors='coerce')
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')

        df = df.dropna(subset=['Type', 'Specification', 'Length', 'Quantity'])
        df = df[(df['Length'] > 0) & (df['Quantity'] > 0)]
        df['Quantity'] = df['Quantity'].astype(int) * multiplier

        total_weight = 0

        for _, row in df.iterrows():
            t = normalize(row['Type'])
            s = normalize(row['Specification'])

            ref = material_df[
                (material_df['__Type_Norm'] == t) &
                (material_df['__Spec_Norm'] == s)
            ]

            if ref.empty:
                missing_weighting.append([
                    file_name,
                    row['Type'],
                    row['Specification'],
                    "Material not found"
                ])
                continue

            unit = ref.iloc[0]['Unit']
            uw = ref.iloc[0]['Unit Weight']

            if pd.isna(uw) or uw == 0:
                missing_weighting.append([
                    file_name,
                    row['Type'],
                    row['Specification'],
                    f"Invalid Unit Weight: {uw}"
                ])
                continue

            uw = float(uw)
            qty = int(row['Quantity'])
            length = float(row['Length'])

            if unit in ["mm", "mm2"]:
                total_weight += length * uw * qty
            elif unit == "piece":
                total_weight += qty * uw

        results.append({
            "Filename": os.path.splitext(os.path.basename(file_name))[0],
            "Quantity": multiplier,
            "Total Weight (kg)": round(total_weight, 3)
        })

    if results:
        df_summary = pd.DataFrame(results)

        total_row = pd.DataFrame([{
            "Filename": "TOTAL",
            "Quantity": "",
            "Total Weight (kg)": round(df_summary["Total Weight (kg)"].sum(), 3)
        }])

        df_summary = pd.concat([df_summary, total_row], ignore_index=True)

        out_path = os.path.join(project_name_path, f"{project_name}_weight.xlsx")
        df_summary.to_excel(out_path, index=False)

        print("⚖️ Weight summary saved:", out_path)


# ---------------------------------------------------------
# COST SUMMARY
# ---------------------------------------------------------
def cost_summary():
    results = []
    missing_costing = []

    for _, file_row in file_list_df.iterrows():
        file_name = normalize_input_filename(file_row['Filename'])
        multiplier = int(file_row['Quantity'])

        # file_name from Excel is already a full path → use directly
        file_path = os.path.join(SOURCE_FOLDER, "database", file_name)

        if not os.path.exists(file_path):
            missing_files.add(file_name)
            continue

        df = pd.read_excel(file_path)

        df['Length'] = pd.to_numeric(df['Length'], errors='coerce')
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')

        df = df.dropna(subset=['Type', 'Specification', 'Length', 'Quantity'])
        df = df[(df['Length'] > 0) & (df['Quantity'] > 0)]
        df['Quantity'] = df['Quantity'].astype(int) * multiplier

        total_cost = 0
        for _, row in df.iterrows():
            t = normalize(row['Type'])
            s = normalize(row['Specification'])

            ref = material_df[(material_df['__Type_Norm'] == t) &
                              (material_df['__Spec_Norm'] == s)]

            if ref.empty:
                missing_costing.append([file_name, row['Type'], row['Specification'], "Not found"])
                continue

            unit = ref.iloc[0]['Unit']
            uc = ref.iloc[0]['Unit Cost']

            if pd.isna(uc) or uc == 0:
                missing_costing.append([file_name, row['Type'], row['Specification'], f"Invalid cost {uc}"])
                continue

            uc = float(uc)
            qty = int(row['Quantity'])
            length = float(row['Length'])

            if unit in ["mm", "mm2"]:
                total_cost += length * uc * qty
            elif unit == "piece":
                total_cost += qty * uc

        results.append({
            # strip folder path → leave only filename
            "Filename": os.path.splitext(os.path.basename(file_name))[0],
            "Quantity": multiplier,
            "Total Cost": round(total_cost, 2)
        })

    if results:
        df_summary = pd.DataFrame(results)

        # Add TOTAL row
        total_value = round(df_summary["Total Cost"].sum(), 2)
        total_row = pd.DataFrame([{
            "Filename": "TOTAL",
            "Quantity": "",
            "Total Cost": total_value
        }])

        df_summary = pd.concat([df_summary, total_row], ignore_index=True)

        out_path = os.path.join(project_name_path, f"{project_name}_cost.xlsx")
        df_summary.to_excel(out_path, index=False)

        print("💰 Cost summary saved:", out_path)

    out_missing = os.path.join(project_name_path, "missing_costing.xlsx")

    if missing_costing:
        df_missing = pd.DataFrame(
            missing_costing,
            columns=["File", "Type", "Spec", "Issue"]
        )
        df_missing.to_excel(out_missing, index=False)
        print("⚠️ Missing COST data report saved.")

    else:
        if os.path.exists(out_missing):
            os.remove(out_missing)
            print("🧹 No missing cost data — removed old report.")
        else:
            print("✅ No missing cost data.")

# ---------------------------------------------------------
# FITTING OUTPUT
# ---------------------------------------------------------
def fitting_output():
    aggregated = {}

    for _, file_row in file_list_df.iterrows():
        file_name = file_row['Filename']
        multiplier = int(file_row['Quantity'])
        file_path = os.path.join(SOURCE_FOLDER, "database",file_name)

        if not os.path.exists(file_path):
            continue

        df = pd.read_excel(file_path)
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
        df = df.dropna(subset=['Type', 'Specification', 'Quantity'])
        df['Quantity'] = df['Quantity'].astype(int) * multiplier

        df_fit = df[
            df['Type'].str.upper().str.contains(fitting_keywords, na=False) |
            df['Specification'].str.upper().str.contains(fitting_keywords, na=False)
        ]

        for _, row in df_fit.iterrows():
            key = (row['Type'], row['Specification'])
            aggregated[key] = aggregated.get(key, 0) + row['Quantity']

    if aggregated:
        rows = []
        for (t, s), qty in aggregated.items():
            rows.append({"Type": t, "Specification": s, "Total Quantity": qty})

            sf_t = safe_filename(t)
            sf_s = safe_filename(s)

            df_out = pd.DataFrame([{"Item": f"{t} - {s}", "Total Quantity": qty}])
            df_out.to_excel(os.path.join(project_name_path, f"{sf_t}_{sf_s}_{qty} pcs.xlsx"), index=False)

        pd.DataFrame(rows).to_excel(os.path.join(project_name_path, "fittings_summary.xlsx"), index=False)
        print("📦 Fittings output completed.")
    else:
        print("ℹ️ No fittings found.")

# ---------------------------------------------------------
# MISSING FILE REPORT
# ---------------------------------------------------------
def save_missing_files_report():
    out_name = f"{project_name}_missing_files.xlsx"
    out_path = os.path.join(project_name_path, out_name)

    if missing_files:
        # Keep order as in the input BOM
        ordered_missing = [
            normalize_input_filename(f)
            for f in file_list_df['Filename']
            if normalize_input_filename(f) in missing_files
        ]

        pd.DataFrame(
            ordered_missing,
            columns=["Missing File"]
        ).to_excel(out_path, index=False)

        print(f"⚠️ Missing files report saved (ordered as BOM): {out_name}")

    else:
        if os.path.exists(out_path):
            os.remove(out_path)
            print("🧹 No missing files found — old missing files report removed.")
        else:
            print("✅ No missing files found.")

def save_missing_weighting_report():
    out_path = os.path.join(project_name_path, "missing_weighting.xlsx")

    if missing_weighting:
        df = pd.DataFrame(
            missing_weighting,
            columns=["File", "Type", "Spec", "Issue"]
        )
        df.to_excel(out_path, index=False)
        print("⚠️ Missing WEIGHT data report saved.")

    else:
        if os.path.exists(out_path):
            os.remove(out_path)
            print("🧹 No missing weight data — removed old report.")
        else:
            print("✅ No missing weight data.")

# ---------------------------------------------------------
# RUN EVERYTHING
# ---------------------------------------------------------
bar_nesting()
weight_summary()
cost_summary()
fitting_output()
save_missing_files_report()
save_missing_weighting_report()

# ---------------------------------------------------------
# FINAL AUTO OPEN LOGIC
# ---------------------------------------------------------
def open_file(file):
    if platform.system() == "Windows":
        os.startfile(file)
    elif platform.system() == "Darwin":
        subprocess.call(["open", file])
    else:
        subprocess.call(["xdg-open", file])

def final_auto_open():
    files_to_open = []

    # Missing reports
    missing_files_path = os.path.join(project_name_path, f"{project_name}_missing_files.xlsx")
    missing_weight_path = os.path.join(project_name_path, "missing_weighting.xlsx")
    missing_cost_path = os.path.join(project_name_path, "missing_costing.xlsx")

    # Normal outputs
    cost_path = os.path.join(project_name_path, f"{project_name}_cost.xlsx")
    weight_path = os.path.join(project_name_path, f"{project_name}_weight.xlsx")

    # Priority: Missing reports first
    if os.path.exists(missing_files_path):
        files_to_open.append(missing_files_path)

    if os.path.exists(missing_weight_path):
        files_to_open.append(missing_weight_path)

    if os.path.exists(missing_cost_path):
        files_to_open.append(missing_cost_path)

    # If NO missing reports → open cost & weight instead
    if not files_to_open:
        if os.path.exists(cost_path):
            files_to_open.append(cost_path)
        if os.path.exists(weight_path):
            files_to_open.append(weight_path)

    for file in files_to_open:
        try:
            open_file(file)
        except Exception as e:
            print(f"⚠️ Could not open file: {file} → {e}")

final_auto_open()

print("\n🎉 All processing completed successfully!")
