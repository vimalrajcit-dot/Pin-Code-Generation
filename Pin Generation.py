import streamlit as st
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile
import os

# =======================
# STREAMLIT START
# =======================
st.set_page_config(page_title="PIN Code Generator", layout="wide")
st.title("🚀 PIN Code Generator")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Read the uploaded file
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    
    st.info("File uploaded successfully ✅")
    
    # Optional: Show preview of data
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
        # Functions needed for mapping are defined after plug/trim, so we'll define a placeholder lambda here
        df["Model Number-Code"] = df["Model Number"].astype(str).apply(lambda x: next((v for k,v in model_map.items() if k in str(x)), ""))
        progress.progress(5)

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
        df["In x Body x Out Size-Code"] = df["In x Body x Out Size"].astype(str).apply(lambda x: next((v for k,v in size_map.items() if k in str(x)), ""))
        progress.progress(10)

        # =======================
        # RATING CLASS
        # =======================
        status.text("Processing Rating Class...")
        rating_map = {"150": "1", "300": "2", "600": "3", "900": "4", "1500": "5", "2500": "6"}
        df["Rating Class-Code"] = df["Rating Class"].astype(str).apply(lambda x: next((v for k,v in rating_map.items() if k in str(x)), ""))
        progress.progress(15)

        # =======================
        # END CONNECTION
        # =======================
        status.text("Processing End Connection...")
        end_conn_map = {"RF": "RF", "FF": "FF", "RTJ": "RJ", "Lugged": "LG", "BW": "BW", "SW": "SW"}
        df["End Connection-Code"] = df["End Connection"].astype(str).apply(lambda x: next((v for k,v in end_conn_map.items() if k in str(x)), ""))
        progress.progress(20)

        # =======================
        # BODY MATERIAL
        # =======================
        status.text("Processing Body Material...")
        body_mat_map = {
            "WCC": "A", "LCC": "B", "A105": "C", "LF2": "D",
            "CF8 ": "E", "CF3 ": "F", "CF8M": "G", "CF3M": "H",
            "Duplex": "I", "Super Duplex": "J", "Aluminum Bronze": "K"
        }
        df["Body Material-Code"] = df["Body Material"].astype(str).apply(lambda x: next((v for k,v in body_mat_map.items() if k in str(x)), ""))
        progress.progress(25)

        # =======================
        # BODY STUDS
        # =======================
        status.text("Processing Body Studs...")
        df["Body Studs-Code"] = df["Body Studs"].apply(lambda x: "2" if pd.notna(x) and "coat" in str(x).lower() else "1")
        progress.progress(30)

        # =======================
        # BONNET TYPE
        # =======================
        status.text("Processing Bonnet Type...")
        bonnet_map = {"Standard": "ST", "Extended": "EB", "Finned": "FB"}
        df["Bonnet Type-Code"] = df["Bonnet Type"].apply(lambda x: next((v for k,v in bonnet_map.items() if k in str(x)), "NA"))
        progress.progress(35)

        # =======================
        # ACTUATOR MODEL
        # =======================
        status.text("Processing Actuator Model...")
        act_model_map = {
            "Top Mounted Handwheel": "20", "87": "87", "88": "88",
            "51": "51", "52": "52", "53": "53",
            "37": "37", "38": "38",
            "Electrical Linear": "EL", "Electrical Rotary": "ER"
        }
        df["Actuator Model-Code"] = df["Actuator Model"].apply(lambda x: next((v for k,v in act_model_map.items() if k in str(x)), ""))
        progress.progress(40)

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
        df["Actuator Size-Code"] = df["Actuator Size"].apply(lambda x: next((v for k,v in act_size_map.items() if k in str(x)), ""))
        progress.progress(45)

        # =======================
        # PLUG / TRIM
        # =======================
        # Trim / plug type from Model Number
        df["Plug Type-Code"] = df["Model Number"].apply(lambda x: str(x).split("-")[2] if len(str(x).split("-"))>2 else "")
        df["Trim Type-Code"] = df["Model Number"].apply(lambda x: str(x).split("-")[3] if len(str(x).split("-"))>3 else "")
        df["Seat Type-Code"] = df["Model Number"].apply(lambda x: str(x).split("-")[4] if len(str(x).split("-"))>4 else "")

        # =======================
        # FUNCTIONS (MUST BE HERE, ABOVE USAGE)
        # =======================
        def contains_map(text, mapping, default=""):
            if text is None or str(text).strip() == "":
                return default
            t = str(text)
            for key, value in mapping.items():
                if key.strip() in t:
                    return value
            return default

        def extract_after_dash(text, position):
            if text is None or str(text).strip() == "":
                return ""
            parts = str(text).split("-")
            return parts[position] if len(parts) > position else ""

        def plug_material_code(text):
            if text is None or str(text).strip() == "":
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

        def plug_type_desc(model):
            if model is None or str(model).strip() == "":
                return ""
            m = str(model)
            x = extract_after_dash(m, 2)
            mapping = {}
            if "-41" in m:
                mapping = {"0":"Undefined","3":"Pressure energized PTFE seal ring","4":"With pilot","5":"Metal seal ring","6":"PTFE seal ring","7":"HT metal seal ring","9":"Graphite seal ring"}
            elif "-21" in m:
                mapping = {"0":"Undefined","1":"Contoured","3":"Close Clearance Lo-dB/Anti-cavitation","5":"Over Travel","6":"Soft Seat","7":"Single Stage Lo-dB/Anti-cavitation","8":"Double Stage Anti-cavitation","9":"Double Stage Lo-dB"}
            return mapping.get(x,"")

        def trim_type_desc(model):
            if model is None or str(model).strip() == "":
                return ""
            m = str(model)
            y = extract_after_dash(m, 4)
            mapping = {}
            if "-41" in m:
                mapping = {"0":"Undefined","1":"Standard cage / Linear","2":"Standard cage / Equal percentage","3":"Lo-dB® / Anticavitation single stage / Linear"}
            elif "-18" in m:
                mapping = {"0":"Optional Trim","1":"Trim A, Balanced Hard Seat"}
            return mapping.get(y,"")

        # Apply plug/trim functions
        df["Plug Type-Des"] = df["Model Number"].apply(plug_type_desc)
        df["Trim Type-Des"] = df["Model Number"].apply(trim_type_desc)
        df["Plug Material-Code"] = df["Plug Material"].apply(plug_material_code)

        # =======================
        # TRIM CHARACTERISTIC
        # =======================
        status.text("Processing Trim Characteristic...")
        trim_char_map = {"Linear":"1","Lin":"1","Equal Percentage":"2","Equal":"2","EQ":"2","Modified Percentage":"3","Quick Opening":"4"}
        df["Trim Characteristic-Code"] = df["Trim Characteristic"].apply(lambda x: contains_map(x, trim_char_map))
        progress.progress(60)

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
        progress.progress(85)

        # =======================
        # PIN DESCRIPTION
        # =======================
        status.text("Generating PIN Description...")
        desc_columns = [
            "Model Number","In x Body x Out Size","Rating Class","End Connection",
            "Body Material","Body Studs","Bonnet Type","Actuator Model","Actuator Size",
            "Plug Material","Trim Type-Des","Plug Type-Des","Seat Type","Trim Characteristic"
        ]
        df["PIN-Code description"] = df[desc_columns].fillna("").astype(str).agg(", ".join, axis=1)
        progress.progress(90)

        # =======================
        # SAVE AND FORMAT EXCEL
        # =======================
        status.text("Saving file...")
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            output_path = tmp_file.name
        df.to_excel(output_path, index=False)

        wb = load_workbook(output_path)
        ws = wb.active
        light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Map headers
        header_map = {ws.cell(row=1, column=col).value: col for col in range(1, ws.max_column+1) if ws.cell(row=1, column=col).value}

        # Color PIN-Code-Length
        if "PIN-Code-Length" in header_map:
            col_idx = header_map["PIN-Code-Length"]
            for row in range(2, ws.max_row+1):
                cell = ws.cell(row=row, column=col_idx)
                try:
                    if cell.value and int(cell.value)<18:
                        cell.fill = light_blue_fill
                except: pass

        # Format headers and empty cells
        new_columns = pin_columns + ["Trim Type-Des","Plug Type-Des","PIN-Code","PIN-Code description","PIN-Code-Length"]
        for col_name in new_columns:
            if col_name in header_map:
                col_idx = header_map[col_name]
                ws.cell(row=1, column=col_idx).fill = green_fill
                for row in range(2, ws.max_row+1):
                    cell = ws.cell(row=row, column=col_idx)
                    if cell.value is None or str(cell.value).strip()=="":
                        cell.fill = yellow_fill

        wb.save(output_path)
        progress.progress(100)
        status.success("✅ Processing complete!")

        with open(output_path,"rb") as f:
            file_data = f.read()
        os.unlink(output_path)

        st.download_button(
            label="📥 Download Processed Excel",
            data=file_data,
            file_name=f"{Path(uploaded_file.name).stem}_PIN_Generated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )