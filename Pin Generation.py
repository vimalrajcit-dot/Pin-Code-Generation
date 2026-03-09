import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# =======================
# HELPER FUNCTIONS
# =======================
def contains_map(text, mapping, default="", method="new"):
    if pd.isna(text) or str(text).strip() == "":
        return default
    target = str(text)
    if method == "new":
        for key in sorted(mapping.keys(), key=len, reverse=True):
            if key in target:
                return mapping[key]
    elif method == "old":
        for key, value in mapping.items():
            if key in target:
                return value
    return default


def extract_after_dash(text, index, method="new"):
    if pd.isna(text) or str(text).strip() == "":
        return ""
    text = str(text)
    if method == "new":
        parts = text.split("-")
        return parts[index].strip() if index < len(parts) else ""
    elif method == "old":
        if "-" not in text:
            return ""
        part = text.split("-", 1)[1]
        return part[index] if index < len(part) else ""
    return ""


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

    # Save uploaded file
    temp_dir = tempfile.mkdtemp()
    file_path = Path(temp_dir) / uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Load Excel into DataFrame
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()
    progress.progress(10)

    # =======================
    # PIN generation mappings
    # =======================
    # 1 MODEL NUMBER - code
    model_map = {
        "-5": "05", "-10": "10", "-18": "18", "-21": "21",
        "-33": "33", "-35": "35", "-41": "41",
        "-77": "77", "-78": "78", "-80": "80", "-28": "28",
    }
    df["Model Number-Code"] = df["Model Number"].apply(lambda x: contains_map(x, model_map))

    # 2 SIZE - code
    size_map = {
        "0.5 x": "05", "0.7 x": "75", "1 x": "01", "1.5 x": "15",
        "2 x": "02", "3 x": "03", "4 x": "04", "6 x": "06",
        "8 x": "08", "10 x": "10", "12 x": "12", "14 x": "14",
        "16 x": "16", "18 x": "18", "20 x": "20", "24 x": "24",
        "26 x": "26", "28 x": "28", "30 x": "30",
        "36 x": "36", "40 x": "40", "42 x": "42", "48 x": "48"
    }
    df["In x Body x Out Size-Code"] = df["In x Body x Out Size"].apply(lambda x: contains_map(x, size_map))

    # 3 RATING CLASS - code
    rating_map = {"150": "1", "300": "2", "600": "3", "900": "4", "1500": "5", "2500": "6"}
    df["Rating Class-Code"] = df["Rating Class"].apply(lambda x: contains_map(x, rating_map))

    # 4 END CONNECTION - code
    end_conn_map = {"RF": "RF", "FF": "FF", "RTJ": "RJ", "Lugged": "LG", "BW": "BW", "SW": "SW"}
    df["End Connection-Code"] = df["End Connection"].apply(lambda x: contains_map(x, end_conn_map))

    # 5 BODY MATERIAL - code
    body_mat_map = {
        "WCC": "A", "LCC": "B", "A105": "C", "LF2": "D",
        "CF8 ": "E", "CF3 ": "F", "CF8M": "G", "CF3M": "H",
        "Duplex": "I", "Super Duplex": "J", "Aluminum Bronze": "K", "12MW": "L", "C95800": "M"
    }
    df["Body Material-Code"] = df["Body Material"].apply(lambda x: contains_map(x, body_mat_map))

    # 6 BODY STUDS - code
    df["Body Studs-Code"] = df["Body Studs"].apply(lambda x: "2" if pd.notna(x) and "coat" in str(x).lower() else "1")

    # 7 BONNET TYPE - code
    bonnet_map = {"Standard": "ST", "Extended": "EB", "Finned": "FB"}
    df["Bonnet Type-Code"] = df["Bonnet Type"].apply(lambda x: contains_map(x, bonnet_map, "NA"))

    # 8 ACTUATOR MODEL - code
    act_model_map = {
        "Top Mounted Handwheel": "20", "87": "87", "88": "88",
        "51": "51", "52": "52", "53": "53",
        "37": "37", "38": "38",
        "Electrical Linear": "EL", "Electrical Rotary": "ER"
    }
    df["Actuator Model-Code"] = df["Actuator Model"].apply(lambda x: contains_map(x, act_model_map))

    # 9 ACTUATOR SIZE - code
    act_size_map = {"6": "A", "12": "B", "16": "C", "20": "D", "23L": "F", "23": "E", "11": "G", "13": "H",
                    "15": "I", "18": "J", "24": "K", "Electric": "L", "10": "M"}
    df["Actuator Size-Code"] = df["Actuator Size"].apply(lambda x: contains_map(x, act_size_map))

    # 10 PLUG MATERIAL - code
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

    # 11 Plug Type-Code
    df["Plug Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 2))

    # 12 Trim Type-Code
    df["Trim Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 3))

    # 13 Seat Type-Code
    df["Seat Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 4))

    # 14 TRIM CHARACTERISTIC - code
    trim_char_map = {
        "Contoured": "A", "Linear": "B", "Equal Percent": "C", "Modified Percentage": "D",
        "Quick Opening": "E", "Anti-Cavitation 1 Stage - Linear": "F",
        "Anti-Cavitation 1 Stage - Equal Percentage": "G",
        "Anti-Cavitation 2 Stage - Linear": "H",
        "Anti-Cavitation 2 Stage - Equal Percentage": "I",
        "50 % LODB 1 Stage Equal % + 50% Contoured": "J",
        "LoDB 1 Stage - Linear": "K",
        "LoDB 2 Stage - Linear": "L",
        "LoDB 1 Stage - Equal Percentage": "M",
        "LoDB 2 Stage - Equal Percentage": "N",
        "Antisurge Lo dB 1 Stage": "O",
        "Antisurge Lo dB 2 Stage": "P",
        "LoDB 1 Stage - Close clearance Linear": "Q",
        "LoDB 2 Stage - Close clearance Linear": "R",
    }
    df["Trim Characteristic-Code"] = df["Trim Characteristic"].apply(lambda x: contains_map(x, trim_char_map))

    # 15 PLUG TYPE DESCRIPTION
    def plug_type_desc(model):
        if pd.isna(model):
            return ""
        m = str(model)
        x = extract_after_dash(m, 2)
        y = extract_after_dash(m, 3)
        if "-41" in m:
            mapping = {"0": "Undefined", "3": "Pressure energized PTFE seal ring", "4": "With pilot",
                       "5": "Metal seal ring", "6": "PTFE seal ring", "7": "HT metal seal ring", "9": "Graphite seal ring"}
            return mapping.get(x, "")
        # add other mappings here
        return ""
    df["Plug Type-Des"] = df["Model Number"].apply(lambda x: plug_type_desc(x))

    # 16 TRIM TYPE DESCRIPTION
    def trim_type_desc(model):
        if pd.isna(model):
            return ""
        m = str(model)
        x = extract_after_dash(m, 3)
        y = extract_after_dash(m, 4)
        if "-41" in m:
            mapping = {"0": "Undefined", "1": "Standard cage / Linear"}  # simplified example
            return mapping.get(x, "")
        return ""
    df["Trim Type-Des"] = df["Model Number"].apply(lambda x: trim_type_desc(x))

    # 17 PIN CODE and description
    pin_columns = [
        "Model Number-Code", "In x Body x Out Size-Code", "Rating Class-Code",
        "End Connection-Code", "Body Material-Code", "Body Studs-Code",
        "Bonnet Type-Code", "Actuator Model-Code", "Actuator Size-Code",
        "Plug Material-Code", "Trim Type-Code", "Seat Type-Code", "Trim Characteristic-Code"
    ]
    df["PIN-Code"] = df[pin_columns].fillna("").astype(str).agg("".join, axis=1)
    df["PIN-Code-Length"] = df["PIN-Code"].astype(str).str.len()

    desc_columns = [
        "Model Number", "In x Body x Out Size", "Rating Class", "End Connection",
        "Body Material", "Body Studs", "Bonnet Type", "Actuator Model", "Actuator Size",
        "Plug Material", "Trim Type-Des", "Plug Type-Des", "Seat Type", "Trim Characteristic"
    ]
    df["PIN-Code description"] = df[desc_columns].fillna("").astype(str).agg(", ".join, axis=1)

    # =======================
    # SAVE & FORMATTING
    # =======================
    output_file = file_path.with_name(file_path.stem + "_PIN_Generated.xlsx")
    df.to_excel(output_file, index=False)

    wb = load_workbook(output_file)
    ws = wb.active

    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    header_map = {ws.cell(row=1, column=col).value: col for col in range(1, ws.max_column + 1)}
    if "PIN-Code-Length" in header_map:
        col_idx = header_map["PIN-Code-Length"]
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            try:
                if int(cell.value) < 18:
                    cell.fill = light_blue_fill
            except:
                pass

    new_columns = [
        "Model Number-Code", "In x Body x Out Size-Code", "Rating Class-Code",
        "End Connection-Code", "Body Material-Code", "Body Studs-Code", "Bonnet Type-Code",
        "Actuator Model-Code", "Actuator Size-Code", "Plug Material-Code", "Plug Type-Code",
        "Trim Type-Code", "Seat Type-Code", "Trim Characteristic-Code",
        "Trim Type-Des", "Plug Type-Des", "PIN-Code", "PIN-Code description"
    ]

    for col_name in new_columns:
        if col_name not in header_map:
            continue
        col_idx = header_map[col_name]
        ws.cell(row=1, column=col_idx).fill = green_fill
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is None or str(cell.value).strip() == "":
                cell.fill = yellow_fill

    wb.save(output_file)

    progress.progress(100)
    status.text("Completed")
    st.success("PIN Generation Completed")

    with open(output_file, "rb") as f:
        st.download_button(
            label="Download Generated File",
            data=f,
            file_name="PIN_Generated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )