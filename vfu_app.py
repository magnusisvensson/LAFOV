import streamlit as stimport streamlit as Font, PatternFill, Alignment
``
import pandas as pd

from openpyxl import Workbook



st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ]

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ✅ kap med ?
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = str(int(float(r["Antal platser"])))
        except:
            kap[r["Skolenhet"]] = "?"

    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    rows = []
    for _, r in skolor.iterrows():

        try:
            antal = int(float(r["Antal platser"]))
        except:
            antal = 0

        for _ in range(antal):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "", "År2": "", "År3": "", "År4": ""
            })

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    bold = Font(bold=True)
    header_font = Font(bold=True, size=14)

    center = Alignment(horizontal="center")
    left = Alignment(horizontal="left")

    fills = {
        "Kalmar": PatternFill("solid", fgColor="D9EAF7"),
        "Oskarshamn": PatternFill("solid", fgColor="DFF5DF"),
        "Karlskrona": PatternFill("solid", fgColor="FFF4CC"),
        "Skola": PatternFill("solid", fgColor="E7E7E7"),
        "Header": PatternFill("solid", fgColor="CCCCCC"),
        "Warning": PatternFill("solid", fgColor="FFC7CE"),
    }

    ws.append(["Skola","År1","År2","År3","År4"])

    for c in range(1,6):
        ws.cell(1,c).fill = fills["Header"]
        ws.cell(1,c).font = bold
        ws.cell(1,c).alignment = center

    row_idx = 2

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        row_idx += 1

        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
        cell = ws.cell(row_idx,1)
        cell.value = region.upper()
        cell.fill = fills[region]
        cell.font = header_font
        cell.alignment = center
        row_idx += 1

        ws.append([])
        row_idx += 1

        for skola in skolor[skolor["Region"] == region]["Skolenhet"]:

            ws.append([f"{skola} (max {kap[skola]})"])

            is_unknown = kap[skola] == "?"

            for c in range(1,6):
                cell = ws.cell(row_idx,c)
                cell.font = bold

                if is_unknown:
                    cell.fill = fills["Warning"]
                else:
                    cell.fill = fills["Skola"]

            row_idx += 1

    wb.save("kull_resultat.xlsx")

    with open("kull_resultat.xlsx","rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f)

else:
    st.info("Ladda upp båda filer")
import pandas as pd

