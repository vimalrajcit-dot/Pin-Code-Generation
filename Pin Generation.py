import streamlit as st
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile

st.title("PIN Code Generator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
run_btn = st.button("Run")

if uploaded_file and run_btn:

    progress = st.progress(0)
    status = st.empty()

    # =======================
    # FILE SETUP
    # =======================
    status.text("Reading Excel file...")
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    progress.progress(5)

    # =======================
    # HELPER FUNCTIONS
    # =======================
    def contains_map(text, mapping, default=""):
        if pd.isna(text):
            return default
        text = str(text)
        for key, value in mapping.items():
            if key in text:
                return value
        return default

    def extract_after_dash(text, index):
        if pd.isna(text) or "-" not in str(text):
            return ""
        part = str(text).split("-", 1)[1]
        return part[index] if len(part) > index else ""

    # =======================
    # MODEL NUMBER
    # =======================
    status.text("Processing Model Number...")
    model_map = {
        "-5": "05", "-10": "10", "-18": "18", "-21": "21",
        "-33": "33", "-35": "35", "-41": "41",
        "-77": "77", "-78": "78", "-80": "80"
    }
    df["Model Number-Code"] = df["Model Number"].apply(lambda x: contains_map(x, model_map))
    progress.progress(15)

    # =======================
    # SIZE
    # =======================
    status.text("Processing Size...")
    size_map = {
        "0.5 x": "05", "0.7 x": "75", "1 x": "01", "1.5 x": "15",
        "2 x": "02", "3 x": "03", "4 x": "04", "6 x": "06",
        "8 x": "08", "10 x": "10", "12 x": "12", "14 x": "14",
        "16 x": "16", "18 x": "18", "20 x": "20", "24 x": "24",
        "26 x": "26", "28 x": "28", "30 x": "30",
        "36 x": "36", "40 x": "40", "42 x": "42", "48 x": "48"
    }
    df["In x Body x Out Size-Code"] = df["In x Body x Out Size"].apply(lambda x: contains_map(x, size_map))
    progress.progress(25)

    # =======================
    # RATING CLASS
    # =======================
    status.text("Processing Rating Class...")
    rating_map = {
        "150": "1", "300": "2", "600": "3",
        "800": "4", "900": "5", "1500": "6", "2500": "7"
    }
    df["Rating Class-Code"] = df["Rating Class"].apply(lambda x: contains_map(x, rating_map))
    progress.progress(35)

    # =======================
    # END CONNECTION
    # =======================
    status.text("Processing End Connection...")
    end_conn_map = {
        "RF": "RF", "FF": "FF", "RTJ": "RJ",
        "Lugged": "LG", "BW": "BW", "SW": "SW"
    }
    df["End Connection-Code"] = df["End Connection"].apply(lambda x: contains_map(x, end_conn_map))
    progress.progress(45)

    # =======================
    # BODY MATERIAL
    # =======================
    status.text("Processing Body Material...")
    body_mat_map = {
        "WCC": "A", "LCC": "B", "A105": "C", "LF2": "D",
        "CF8 ": "E", "CF3 ": "F", "CF8M": "G", "CF3M": "H",
        "Duplex": "I", "Super Duplex": "J", "Aluminum Bronze": "K"
    }
    df["Body Material-Code"] = df["Body Material"].apply(lambda x: contains_map(x, body_mat_map))
    progress.progress(55)

    # =======================
    # BODY STUDS
    # =======================
    status.text("Processing Body Studs...")
    df["Body Studs-Code"] = df["Body Studs"].apply(
        lambda x: "2" if pd.notna(x) and "coat" in str(x).lower() else "1"
    )
    progress.progress(60)

    # =======================
    # BONNET TYPE
    # =======================
    status.text("Processing Bonnet Type...")
    bonnet_map = {"Standard": "ST", "Extended": "EB", "Finned": "FB"}
    df["Bonnet Type-Code"] = df["Bonnet Type"].apply(lambda x: contains_map(x, bonnet_map, "NA"))
    progress.progress(65)

    # =======================
    # ACTUATOR MODEL
    # =======================
    status.text("Processing Actuator Model...")
    act_model_map = {
        "Top Mounted Handwheel": "20",
        "87": "87", "88": "88",
        "51": "51", "52": "52", "53": "53",
        "37": "37", "38": "38",
        "Electrical Linear": "EL",
        "Electrical Rotary": "ER"
    }
    df["Actuator Model-Code"] = df["Actuator Model"].apply(lambda x: contains_map(x, act_model_map))
    progress.progress(70)

    # =======================
    # ACTUATOR SIZE
    # =======================
    status.text("Processing Actuator Size...")
    act_size_map = {
        "6": "A", "12": "B", "16": "C", "20": "D",
        "23L": "F", "23": "E", "11": "G", "13": "H",
        "15": "I", "18": "J", "24": "K",
        "Electric": "L", "10": "M"
    }
    df["Actuator Size-Code"] = df["Actuator Size"].apply(lambda x: contains_map(x, act_size_map))
    progress.progress(75)

    # =======================
    # PLUG MATERIAL
    # =======================
    status.text("Processing Plug Material...")
    def plug_material_code(text):
        if pd.isna(text):
            return ""
        t = str(text)
        if "316" in t and ("Hard" in t or "HF" in t):
            return "1"
        if "316" in t:
            return "2"
        if "410" in t:
            return "3"
        if "CA6NM" in t:
            return "4"
        return ""

    df["Plug Material-Code"] = df["Plug Material"].apply(plug_material_code)
    progress.progress(80)

    # =======================
    # TRIM / SEAT FROM MODEL
    # =======================
    status.text("Processing Trim Codes...")
    df["Plug Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 2))
    df["Trim Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 3))
    df["Seat Type-Code"] = ""
    progress.progress(85)

    # =======================
    # TRIM CHARACTERISTIC
    # =======================
    status.text("Processing Trim Characteristic...")
    trim_char_map = {
        "Linear": "1",
        "Equal Percentage": "2",
        "EQ": "2",
        "Modified Percentage": "3",
        "Quick Opening": "4"
    }
    df["Trim Characteristic-Code"] = df["Trim Characteristic"].apply(lambda x: contains_map(x, trim_char_map))
    progress.progress(90)

    # =======================
    # PIN CODE
    # =======================
    status.text("Generating PIN Code...")
    pin_columns = [
        "Model Number-Code","In x Body x Out Size-Code","Rating Class-Code",
        "End Connection-Code","Body Material-Code","Body Studs-Code",
        "Bonnet Type-Code","Actuator Model-Code","Actuator Size-Code",
        "Plug Material-Code","Trim Type-Code","Seat Type-Code",
        "Trim Characteristic-Code"
    ]

    df["PIN-Code"] = df[pin_columns].fillna("").astype(str).agg("".join, axis=1)
    progress.progress(95)

    # SAVE
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    output_file = tmp.name
    df.to_excel(output_file, index=False)

    # FORMATTING (UNCHANGED)
    wb = load_workbook(output_file)
    ws = wb.active

    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    header_map = {
        ws.cell(row=1, column=col).value: col
        for col in range(1, ws.max_column + 1)
    }

    for col_name in header_map:
        col_idx = header_map[col_name]
        ws.cell(row=1, column=col_idx).fill = green_fill

        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is None or str(cell.value).strip() == "":
                cell.fill = yellow_fill

    wb.save(output_file)
    progress.progress(100)
    status.text("Completed")

    with open(output_file, "rb") as f:
        st.download_button("Download PIN File", f, file_name="PIN_Generated.xlsx")