
import streamlit as st
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


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

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # kap med ?
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = str(int(float(r["Antal platser"])))
        except:
            kap[r["Skolenhet"]] = "?"

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    # ===== PLATSER =====
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

    # ===== PLACERING =====
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        if not rows_r:
            continue

        kapasitet = {sk: 0 for sk in skolor_r}
        for r in rows_r:
            kapasitet[r["Skola"]] += 1

        usage = {sk: {"År1":0,"År2":0,"År3":0,"År4":0} for sk in skolor_r}

        def place(student, year, sk):
            for r in rows_r:
                if r["Skola"] == sk and r[year] == "":
                    r[year] = student
                    usage[sk][year] += 1
                    return True
            return False


        for i, student in enumerate(stud_r):

            A = skolor_r[i % len(skolor_r)]
            B = skolor_r[(i+1) % len(skolor_r)]
            C = skolor_r[(i+2) % len(skolor_r)] if len(skolor_r) > 2 else B

            # År1
            place(student, "År1", A)

            # ✅ År2 + År3 SAMMA
            placed = False
            for sk in [B] + skolor_r:
                if (
                    usage[sk]["År2"] < kapasitet[sk] and
                    usage[sk]["År3"] < kapasitet[sk]
                ):
                    ok2 = place(student, "År2", sk)
                    ok3 = place(student, "År3", sk)
                    if ok2 and ok3:
                        placed = True
                        break

            # ✅ År4 – UNDVIK SAMMA SKOLA
            used = set()
            for r in rows_r:
                for y in ["År1","År2","År3"]:
                    if r[y] == student:
                        used.add(r["Skola"])

            placed4 = False

            # försök NYA skolor först
            for sk in skolor_r:
                if sk in used:
                    continue

                if usage[sk]["År4"] < kapasitet[sk]:
                    if place(student, "År4", sk):
                        placed4 = True
                        break

            # fallback
            if not placed4:
                for sk in skolor_r:
                    if usage[sk]["År4"] < kapasitet[sk]:
                        place(student, "År4", sk)
                        break

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
        "WarningLight": PatternFill("solid", fgColor="FFECEC"),
    }

    ws.append(["Skola","År1","År2","År3","År4"])

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

            for r in rows:
                if r["Skola"] == skola:
                    ws.append(["", r["År1"], r["År2"], r["År3"], r["År4"]])

                    for c in range(2,6):
                        cell = ws.cell(row_idx,c)
                        cell.alignment = left

                        if is_unknown:
                            cell.fill = fills["WarningLight"]

                    row_idx += 1

            ws.append([])
            row_idx += 1

    # kolumnbredd
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 14)

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
