import streamlit as st
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile
import os

st.set_page_config(page_title="PIN Code Generator", layout="wide")
st.title("🚀 PIN Code Generator")


def contains_map(text, mapping, default=""):
    if pd.isna(text):
        return default
    t = str(text)
    for key, value in mapping.items():
        if key.strip() in t:
            return value
    return default


def extract_after_dash(text, position):
    if pd.isna(text):
        return ""
    parts = str(text).split("-")
    return parts[position] if len(parts) > position else ""


# UPLOAD FILE
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Read the uploaded file
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    st.info("File uploaded successfully ✅")
    
    # Show preview of data
    with st.expander("Preview Data"):
        st.write(df.head())

    if st.button("Run Processing"):
        progress = st.progress(0)
        status = st.empty()
        
        # =======================
        # MODEL NUMBER
        # =======================
        status.text("Processing Model Number...")
        model_map = {
            "-5": "05", "-10": "10", "-18": "18", "-21": "21",
            "-33": "33", "-35": "35", "-41": "41",
            "-77": "77", "-78": "78", "-80": "80", "-28": "28",
        }
        df["Model Number-Code"] = df["Model Number"].apply(lambda x: contains_map(x, model_map))
        progress.progress(10)

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
        progress.progress(20)

        # =======================
        # RATING CLASS
        # =======================
        status.text("Processing Rating Class...")
        rating_map = {
            "150": "1", "300": "2", "600": "3",
            "900": "4", "1500": "5", "2500": "6"
        }
        df["Rating Class-Code"] = df["Rating Class"].apply(lambda x: contains_map(x, rating_map))
        progress.progress(30)

        # =======================
        # END CONNECTION
        # =======================
        status.text("Processing End Connection...")
        end_conn_map = {
            "RF": "RF", "FF": "FF", "RTJ": "RJ",
            "Lugged": "LG", "BW": "BW", "SW": "SW"
        }
        df["End Connection-Code"] = df["End Connection"].apply(lambda x: contains_map(x, end_conn_map))
        progress.progress(40)

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
        progress.progress(45)

        # =======================
        # BODY STUDS
        # =======================
        status.text("Processing Body Studs...")
        df["Body Studs-Code"] = df["Body Studs"].apply(
            lambda x: "2" if pd.notna(x) and "coat" in str(x).lower() else "1"
        )
        progress.progress(50)

        # =======================
        # BONNET TYPE
        # =======================
        status.text("Processing Bonnet Type...")
        bonnet_map = {
            "Standard": "ST",
            "Extended": "EB",
            "Finned": "FB"
        }
        df["Bonnet Type-Code"] = df["Bonnet Type"].apply(lambda x: contains_map(x, bonnet_map, "NA"))
        progress.progress(55)

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
        progress.progress(60)

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
        progress.progress(65)

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
        progress.progress(70)

        # =======================
        # TRIM / SEAT FROM MODEL
        # =======================
        status.text("Processing Trim/Seat from Model...")
        df["Plug Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 2))
        df["Trim Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 3))
        df["Seat Type-Code"] = df["Model Number"].apply(lambda x: extract_after_dash(x, 4))
        progress.progress(75)

        # =======================
        # TRIM CHARACTERISTIC
        # =======================
        status.text("Processing Trim Characteristic...")
        trim_char_map = {
            "Linear": "1",
            "Lin": "1",
            "Equal Percentage": "2",
            "Equal": "2",
            "EQ": "2",
            "Modified Percentage": "3",
            "Quick Opening": "4"
        }
        df["Trim Characteristic-Code"] = df["Trim Characteristic"].apply(lambda x: contains_map(x, trim_char_map))
        progress.progress(80)

        # =======================
        # PLUG TYPE DESCRIPTION
        # =======================
        status.text("Processing Plug Type Description...")
        def plug_type_desc(model):
            if pd.isna(model):
                return ""
            
            m = str(model)
            x = extract_after_dash(m, 2)

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
        progress.progress(85)

        # =======================
        # TRIM TYPE DESCRIPTION
        # =======================
        status.text("Processing Trim Type Description...")
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

            return ""

        df["Trim Type-Des"] = df["Model Number"].apply(lambda x: trim_type_desc(x))
        progress.progress(90)

        # =======================
        # PIN CODE
        # =======================
        status.text("Generating PIN Code...")
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
        progress.progress(95)

        # =======================
        # PIN DESCRIPTION
        # =======================
        status.text("Generating PIN Description...")
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

        # =======================
        # SAVE FILE
        # =======================
        status.text("Saving file...")
        
        # Create a temporary file to save the Excel
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            output_path = tmp_file.name
        
        # Save the dataframe to the temporary file
        df.to_excel(output_path, index=False)

        # =======================
        # FORMATTING
        # =======================
        wb = load_workbook(output_path)
        ws = wb.active

        light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Get header mapping
        header_map = {}
        for col in range(1, ws.max_column + 1):
            cell_value = ws.cell(row=1, column=col).value
            if cell_value:
                header_map[cell_value] = col

        # Format PIN-Code-Length column
        if "PIN-Code-Length" in header_map:
            col_idx = header_map["PIN-Code-Length"]
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                try:
                    if cell.value and int(cell.value) < 18:
                        cell.fill = light_blue_fill
                except:
                    pass

        # Format new columns
        new_columns = [
            "Model Number-Code", "In x Body x Out Size-Code", "Rating Class-Code",
            "End Connection-Code", "Body Material-Code", "Body Studs-Code",
            "Bonnet Type-Code", "Actuator Model-Code", "Actuator Size-Code",
            "Plug Material-Code", "Plug Type-Code", "Trim Type-Code",
            "Seat Type-Code", "Trim Characteristic-Code", "Trim Type-Des",
            "Plug Type-Des", "PIN-Code", "PIN-Code description", "PIN-Code-Length"
        ]

        for col_name in new_columns:
            if col_name in header_map:
                col_idx = header_map[col_name]
                
                # Green fill for headers
                header_cell = ws.cell(row=1, column=col_idx)
                header_cell.fill = green_fill
                
                # Yellow fill for empty cells
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col_idx)
                    if cell.value is None or str(cell.value).strip() == "":
                        cell.fill = yellow_fill

        # Save the formatted workbook
        wb.save(output_path)

        progress.progress(100)
        status.success("✅ Processing complete!")

        # Read the file for download
        with open(output_path, "rb") as f:
            file_data = f.read()
        
        # Clean up temp file
        os.unlink(output_path)

        # Provide download button
        st.download_button(
            label="📥 Download Processed Excel",
            data=file_data,
            file_name=f"{Path(uploaded_file.name).stem}_PIN_Generated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )