import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def contains_map(text, mapping, default=""):
    if text is None or pd.isna(text) or str(text).strip() == "":
        return default
    target = str(text)
    # Match longer keys first to avoid partial-match mistakes (e.g., 150 before 1500).
    for key in sorted(mapping.keys(), key=len, reverse=True):
        if key in target:
            return mapping[key]
    return default


def extract_after_dash(text, index):
    if text is None or pd.isna(text) or str(text).strip() == "":
        return ""
    parts = str(text).split("-")
    return parts[index] if len(parts) > index else ""



# =======================
# STREAMLIT UI
# =======================
st.set_page_config(page_title="PIN Code Generator", layout="wide")
st.title("PIN Code Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
run_process = st.button("Run PIN Generation")

if uploaded_file and run_process:
    progress = st.progress(0)
    status = st.empty()

    # Save uploaded file to temp location
    temp_dir = tempfile.mkdtemp()
    file_path = os.path.join(temp_dir, uploaded_file.name)

    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    progress.progress(10)


# =======================# ===== DO NOT MODIFY START =====
#1 MODEL NUMBER - code
# =======================
model_map = {
    "-5": "05", "-10": "10", "-18": "18", "-21": "21",
    "-33": "33", "-35": "35", "-41": "41",
    "-77": "77", "-78": "78", "-80": "80", "-28": "28",
}
df["Model Number-Code"] = df["Model Number"].apply(lambda x: contains_map(x, model_map))

# =======================
#2 SIZE - code
# =======================
size_map = {
    "0.5 x": "05", "0.7 x": "75", "1 x": "01", "1.5 x": "15",
    "2 x": "02", "3 x": "03", "4 x": "04", "6 x": "06",
    "8 x": "08", "10 x": "10", "12 x": "12", "14 x": "14",
    "16 x": "16", "18 x": "18", "20 x": "20", "24 x": "24",
    "26 x": "26", "28 x": "28", "30 x": "30",
    "36 x": "36", "40 x": "40", "42 x": "42", "48 x": "48"
}
df["In x Body x Out Size-Code"] = df["In x Body x Out Size"].apply(lambda x: contains_map(x, size_map))

# =======================
#3 RATING CLASS  - code
# =======================
rating_map = {
    "150": "1", "300": "2", "600": "3",
    "900": "4", "1500": "5", "2500": "6"
}
df["Rating Class-Code"] = df["Rating Class"].apply(lambda x: contains_map(x, rating_map))

# =======================
#4 END CONNECTION - code
# =======================
end_conn_map = {
    "RF": "RF", "FF": "FF", "RTJ": "RJ",
    "Lugged": "LG", "BW": "BW", "SW": "SW"
}
df["End Connection-Code"] = df["End Connection"].apply(lambda x: contains_map(x, end_conn_map))

# =======================
#5 BODY MATERIAL - code
# =======================
body_mat_map = {
    "WCC": "A", "LCC": "B", "A105": "C", "LF2": "D",
    "CF8 ": "E", "CF3 ": "F", "CF8M": "G", "CF3M": "H",
    "Duplex": "I", "Super Duplex": "J", "Aluminum Bronze": "K" , "12MW" : "L" , "C95800" : "M"
}
df["Body Material-Code"] = df["Body Material"].apply(lambda x: contains_map(x, body_mat_map))

# =======================
#6 BODY STUDS    - code
# =======================
df["Body Studs-Code"] = df["Body Studs"].apply(
    lambda x: "2" if pd.notna(x) and "coat" in str(x).lower() else "1"
)

# =======================
#7 BONNET TYPE   - code
# =======================
bonnet_map = {
    "Standard": "ST",
    "Extended": "EB",
    "Finned": "FB"
}
df["Bonnet Type-Code"] = df["Bonnet Type"].apply(lambda x: contains_map(x, bonnet_map, "NA"))

# =======================
#8 ACTUATOR MODEL - code
# =======================
act_model_map = {
    "Top Mounted Handwheel": "20",
    "87": "87", "88": "88",
    "51": "51", "52": "52", "53": "53",
    "37": "37", "38": "38",
    "Electrical Linear": "EL",
    "Electrical Rotary": "ER"
}
df["Actuator Model-Code"] = df["Actuator Model"].apply(lambda x: contains_map(x, act_model_map))

# =======================
#9 ACTUATOR SIZE - code
# =======================
act_size_map = {
    "6": "A", "12": "B", "16": "C", "20": "D",
    "23L": "F", "23": "E", "11": "G", "13": "H",
    "15": "I", "18": "J", "24": "K",
    "Electric": "L", "10": "M"
}
df["Actuator Size-Code"] = df["Actuator Size"].apply(lambda x: contains_map(x, act_size_map))

# =======================
#10 PLUG MATERIAL - code
# =======================
def plug_material_code(text):
    if pd.isna(text):
        return ""
    t = str(text)
    if "316" in t and ("Hard" in t or "HF" in t):
        return "3"
    if "316" in t:
        return "2"
    if "410" in t:
        return "1"
    if "CA6NM" in t and ("Plating" in t or "coat" in t):
        return "5"
    if "CA6NM" in t:
        return "4"
    if "31254" in t:
        return "6"
    if "C276" in t:
        return "7"
    if "Monel" in t:
        return "8"  
    if "stellite" in t:
        return "9"
    return ""

df["Plug Material-Code"] = df["Plug Material"].apply(plug_material_code)

# =======================
#11 Plug Type-Code
# =======================
df["Plug Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 2))

# =======================
#12 Trim Type-Code
# =======================
df["Trim Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 3))

# =======================
#13 Seat Type-Code
# =======================
df["Seat Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 4))

# =======================
#13 TRIM CHARACTERISTIC - code
# =======================
trim_char_map = {
    "Contoured"	: "A",
    "Linear" : "B",
    "Equal Percent" : "C",
    "Modified Percentage" :"D",
    "Quick Opening"	:	"E",
    "Anti-Cavitation 1 Stage - Linear"	: "F",
    "Anti-Cavitation 1 Stage - Equal Percentage"	: "G",
    "Anti-Cavitation 2 Stage - Linear"	:"H",
    "Anti-Cavitation 2 Stage - Equal Percentage"	:"I",
    "50 % LODB 1 Stage Equal % + 50% Contoured"	:"J",
    "LoDB 1 Stage - Linear"	:"K",
    "LoDB 2 Stage - Linear"	:"L",
    "LoDB 1 Stage - Equal Percentage"	:"M",
    "LoDB 2 Stage - Equal Percentage"	:"N",
    "Antisurge Lo dB 1 Stage"	:"O",
    "Antisurge Lo dB 2 Stage"	:"P",
    "LoDB 1 Stage - Close clearance Linear"	:"Q",
    "LoDB 2 Stage - Close clearance Linear"	:"R"   ,
}
df["Trim Characteristic-Code"] = df["Trim Characteristic"].apply(lambda x: contains_map(x, trim_char_map))

# =======================
#14 PLUG TYPE DESCRIPTION
# =======================
def plug_type_desc(model):
    if pd.isna(model):
        return ""
    
    m = str(model)
    x = extract_after_dash(m, 2)
    y = extract_after_dash(m, 3)

    if "-41" in m:
        mapping = {
            "0": "Undefined",
            "3": "Pressure energized PTFE seal ring",
            "4": "With pilot",
            "5": "Metal seal ring",
            "6": "PTFE seal ring",
            "7": "HT metal seal ring",
            "9": "Graphite seal ring",
        }
        return mapping.get(x, "")

    if "-21" in m:
        mapping = {
            "0": "Undefined",
            "1": "Contoured",
            "3": "Close Clearance Lo-dB/Anti-cavitation",
            "5": "Over Travel",
            "6": "Soft Seat",
            "7": "Single Stage Lo-dB/Anti-cavitation",
            "8": "Double Stage Anti-cavitation",
            "9": "Double Stage Lo-dB",
        }
        return mapping.get(x, "")

    if "-10" in m:
        mapping = {
            "0": "Undefined",
            "1": "Double Seat",

        }
        return mapping.get(x, "")   
    
    if "-18" in m:
        mapping = {
            "0": "Undefined",
            "1": "Axial Flow High Resistance (Downseating)",

        }
        return mapping.get(x, "") 
    
    if "-78" in m:
        mapping = {
            "0": "Undefined",
            "1": "Axial Flow High Resistance (Downseating)",

        }
        return mapping.get(x, "")
    
    if "-80" in m:
        mapping = {
            "0": "Undefined",
            "3": "Top and Port Guided",

        }
        return mapping.get(x, "")
    
    if "-33" in m:
        mapping = {
            "0": "Triple Offset",
            
        }
        return mapping.get(y, "")
    
    if "-35" in m:
        mapping = {
            "0": "Self-aligning eccentrically rotatingt",
            
        }
        return mapping.get(x, "")
    
    if "-28" in m:
        mapping = {
            "0": "3.8, Linear",
            "1": "2.3, Linear",
            "2": "1.2, Linear",
            "3": "0.6, Linear",
            "4": "0.25, Linear",
            "5": "0.1, Linear",
            "6": "0.05, Modified Linear",
            "7": "0.025, Modified Linear",
            "8": "0.01, Modified Linear",
            "9": "0.004, Modified Linear"
        }
        return mapping.get(y, "")
    
    if "-77" in m:
        mapping = {
            "0": "Undefined",
            "1": "Trim A: 9-stage unbalanced",
            "2": "Trim B: 5-stage unbalanced",
            "3": "Trim C: 1-stage unbalanced",
            "4": "Trim X: 5-stage partially balanced",
            "5": "Trim Y: 3-stage unbalanced",
            
        }
        return mapping.get(x, "")

    return ""

df["Plug Type-Des"] = df["Model Number"].apply(lambda x: plug_type_desc(x))


# =======================
#15 TRIM TYPE DESCRIPTION
# =======================
def trim_type_desc(model):
    if pd.isna(model):
        return ""

    m = str(model)
    x = extract_after_dash(m, 3)
    y = extract_after_dash(m, 4)

    if "-41" in m:
        mapping = {
            "0": "Undefined",
            "1": "Standard cage / Linear",
            "2": "Standard cage / Equal percentage",
            "3": "Lo-dB® / Anticavitation single stage / Linear",
            "4": "Lo-dB® single stage with diffuser / Linear",
            "5": "Lo-dB® double stage / Linear",
            "6": "VRT (stack) Type S / Linear",
            "7": "VRT (stack partial) Type S / modified percentage",
            "8": "VRT (cage) Type C / Linear",
            "9": "Anticavitation double stage / Linear (1)",
            "A": "High Capacity Linear",
            "B": "High Capacity Equal %",
            "C": "High Capacity Lo-DB / Anti-Cav",
        }
        return mapping.get(x, "")

    if "-21" in m:
        mapping = {
            "0": "Undefined",
            "4": "Quick Change",
            "5": "Threaded",
        }
        return mapping.get(y, "")
    
    if "-18" in m:
        mapping = {
            "0": "Optional Trim",
            "1": "Trim A, Balanced Hard Seat",
            "2": "Trim B, Balanced Hard Seat",
            "3": "Trim C, Balanced Hard Seat",
            "4": "Trim A, Balanced Soft Seat",
            "5": "Trim B, Balanced Soft Seat",
            "6": "Trim C, Balanced Soft Seat",
            "7": "Trim A, Unbalanced Hard Seat",
            "8": "Trim B, Unbalanced Hard Seat",
            "9": "Trim C, Unbalanced Hard Seat",
        }
        return mapping.get(y, "")
    
    if "-78" in m:
        mapping = {
            "0": "Optional Trim",
            "1": "Trim A, Balanced Hard Seat",
            "2": "Trim B, Balanced Hard Seat",
            "3": "Trim C, Balanced Hard Seat",
            "4": "Trim A, Balanced Soft Seat",
            "5": "Trim B, Balanced Soft Seat",
            "6": "Trim C, Balanced Soft Seat",
            "7": "Trim A, Unbalanced Hard Seat",
            "8": "Trim B, Unbalanced Hard Seat",
            "9": "Trim C, Unbalanced Hard Seat",
        }
        return mapping.get(y, "")
    
    if "-77" in m:
        mapping = {
            "0": "Optional Trim",
            "1": "Bottom entry; outlet spool",
            "2": "Top entry; bolted bonnet",
            "3": "Compact Top entry; bolted bonnet",
        }
        return mapping.get(y, "")
    
    if "-80" in m:
        mapping = {
            "0": "Combined design",
            "1": "Diverting design",
        
        }
        return mapping.get(y, "")
    
    if "-28" in m:
        mapping = {
            "0": "Metal seat",
                    
        }
        return mapping.get(y, "")
    
    if "-33" in m:
        mapping = {
            "0": "Metal + Graphite Laminated",
                    
        }
        return mapping.get(y, "")
    
    if "-35" in m:
        mapping = {
            "1": "Metal Seat",
            "2": "Soft Seat",
            "3": "Metal Seat w/ Differential Velocity Trim",
            "4": "Soft Seat w/ Differential Velocity Trim",
        }
        return mapping.get(y, "")

    return ""

df["Trim Type-Des"] = df["Model Number"].apply(lambda x: trim_type_desc(x))

# =======================
# 16 PIN CODE
# =======================
pin_columns = [
    "Model Number-Code",
    "In x Body x Out Size-Code",
    "Rating Class-Code",
    "End Connection-Code",
    "Body Material-Code",
    "Body Studs-Code",
    "Bonnet Type-Code",
    "Actuator Model-Code",
    "Actuator Size-Code",
    "Plug Material-Code",
    "Trim Type-Code",
    "Seat Type-Code",
    "Trim Characteristic-Code"
]

df["PIN-Code"] = df[pin_columns].fillna("").astype(str).agg("".join, axis=1)
df["PIN-Code-Length"] = df["PIN-Code"].astype(str).str.len()

# =======================
# 17 PIN DESCRIPTION
# =======================
desc_columns = [
    "Model Number",
    "In x Body x Out Size",
    "Rating Class",
    "End Connection",
    "Body Material",
    "Body Studs",
    "Bonnet Type",
    "Actuator Model",
    "Actuator Size",
    "Plug Material",
    "Trim Type-Des",
    "Plug Type-Des",
    "Seat Type",
    "Trim Characteristic"
]

df["PIN-Code description"] = df[desc_columns].fillna("").astype(str).agg(", ".join, axis=1)

# =======================# ===== DO NOT MODIFY END =====
# SAVE FILE
# =======================
output_file = file_path.with_name(file_path.stem + "_PIN_Generated.xlsx")
df.to_excel(output_file, index=False)

# =======================
# FORMATTING
# =======================
wb = load_workbook(output_file)
ws = wb.active

from openpyxl.styles import PatternFill

light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

# =======================
# 18 PIN Length
# =======================
header_map = {
    ws.cell(row=1, column=col).value: col
    for col in range(1, ws.max_column + 1)
}

if "PIN-Code-Length" in header_map:
    col_idx = header_map["PIN-Code-Length"]

    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_idx)
        try:
            if int(cell.value) < 18:
                cell.fill = light_blue_fill
        except:
            pass

green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

new_columns = [
    "Model Number-Code",
    "In x Body x Out Size-Code",
    "Rating Class-Code",
    "End Connection-Code",
    "Body Material-Code",
    "Body Studs-Code",
    "Bonnet Type-Code",
    "Actuator Model-Code",
    "Actuator Size-Code",
    "Plug Material-Code",
    "Plug Type-Code",
    "Trim Type-Code",
    "Seat Type-Code",
    "Trim Characteristic-Code",
    "Trim Type-Des",
    "Plug Type-Des",
    "PIN-Code",
    "PIN-Code description"
]

header_map = {
    ws.cell(row=1, column=col).value: col
    for col in range(1, ws.max_column + 1)
}

for col_name in new_columns:
    if col_name not in header_map:
        continue

    col_idx = header_map[col_name]
    ws.cell(row=1, column=col_idx).fill = green_fill

    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_idx)
        if cell.value is None or str(cell.value).strip() == "":
            cell.fill = yellow_fill

wb.save(output_path)  

progress.progress(100)
status.text("Completed")
st.success("PIN Generation Completed")

if Path(output_path).exists():
        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Generated File",
                data=f,
                file_name="PIN_Generated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )