
import streamlit as st
 openpyxl styles import Font, PatternFill, Alignmentimport pandas as pd

from openpyxl import Workbook


st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# ===== REGION =====
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

    kap = {
        r["Skolenhet"]: int(r["Antal platser"])
        for _, r in skolor.iterrows()
        if pd.notna(r["Antal platser"])
    }

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    # ===== PLATSRADER =====
    rows = []
    for _, r in skolor.iterrows():
        for _ in range(int(r["Antal platser"])):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "", "År2": "", "År3": "", "År4": ""
            })

    not_placed = []

    # ===== LOGIK =====
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        if not rows_r:
            not_placed += stud_r
            continue

        kapasitet = {sk: 0 for sk in skolor_r}
        for r in rows_r:
            kapasitet[r["Skola"]] += 1

        usage = {
            sk: {"År1":0,"År2":0,"År3":0,"År4":0}
            for sk in skolor_r
        }

        def try_place(student, start_index):

            A = skolor_r[start_index]

            if len(skolor_r) <= 2:
                B = skolor_r[(start_index+1) % len(skolor_r)]
                schedule = {"År1":A,"År2":B,"År3":A,"År4":B}
            else:
                B = skolor_r[(start_index+1) % len(skolor_r)]
                C = skolor_r[(start_index+2) % len(skolor_r)]
                schedule = {"År1":A,"År2":B,"År3":B,"År4":C}

            for y in schedule:
                if usage[schedule[y]][y] >= kapasitet[schedule[y]]:
                    return False

            for year in schedule:
                sk = schedule[year]
                for r in rows_r:
                    if r["Skola"] == sk and r[year] == "":
                        r[year] = student
                        usage[sk][year] += 1
                        break
            return True


        def fallback_place(student):
            for year in ["År1","År2","År3","År4"]:
                for sk in skolor_r:
                    if usage[sk][year] < kapasitet[sk]:
                        for r in rows_r:
                            if r["Skola"] == sk and r[year] == "":
                                r[year] = student
                                usage[sk][year] += 1
                                break
                        break

        for i, student in enumerate(stud_r):

            placed = False

            for shift in range(len(skolor_r)):
                if try_place(student, (i+shift) % len(skolor_r)):
                    placed = True
                    break

            if not placed:
                fallback_place(student)


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

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
    }

    ws.append(["Skola","År1","År2","År3","År4"])

    for c in range(1,6):
        cell = ws.cell(1,c)
        cell.fill = fills["Header"]
        cell.font = bold
        cell.alignment = center

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

        skolor_region = skolor[skolor["Region"] == region]["Skolenhet"]

        for skola in skolor_region:

            ws.append([f"{skola} (max {kap[skola]})"])

            for c in range(1,6):
                cell = ws.cell(row_idx,c)
                cell.fill = fills["Skola"]
                cell.font = bold

            row_idx += 1

            for r in rows:
                if r["Skola"] == skola:
                    ws.append(["", r["År1"], r["År2"], r["År3"], r["År4"]])

                    # ✅ vänsterjustera namn
                    for c in range(2,6):
                        ws.cell(row_idx,c).alignment = left

                    row_idx += 1

            ws.append([])
            row_idx += 1

    # ✅ förbättrad kolumnbredd
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                length = len(str(cell.value))
                if length > max_len:
                    max_len = length

        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in students["Namn"]:
        status = "EJ PLACERAD" if s in not_placed else "OK"
        ws2.append([s, status])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
import pandas as pd
