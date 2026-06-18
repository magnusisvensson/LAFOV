
# =========================
# IMPORT
# =========================
import streamlit as st
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# =========================
# REGION
# =========================
def get_region(text):
    t = str(text).lower()

    if "oskarshamn" in t:
        return "Oskarshamn"

    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"

    return "Kalmar"


# =========================
# KÖR
# =========================
if system_file and form_file:

    skolor = pd.read_excel(system_file, engine="openpyxl")
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # kapacitet
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 0

    # =========================
    # STUDENTER
    # =========================
    students = pd.read_excel(form_file, sheet_name="Data", engine="openpyxl")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    alt_col = [c for c in students.columns if "alternativ" in c.lower()][0]
    pref_col = [c for c in students.columns if "helst utgå" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]

    def choose_loc(row):
        if "alternativ" in str(row[pref_col]).lower():
            if pd.notna(row[alt_col]):
                return row[alt_col]
        return row[bost]

    students["ChosenOrt"] = students.apply(choose_loc, axis=1)
    students["Region"] = students["ChosenOrt"].apply(get_region)

    # =========================
    # PLATSER
    # =========================
    rows = []

    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "", "År2": "", "År3": "", "År4": ""
            })

    # =========================
    # PLACERING
    # =========================
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

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
            C = skolor_r[(i+2) % len(skolor_r)] if len(skolor_r)>2 else B

            if program == "LGFRI":
                place(student,"År1",A)
                place(student,"År2",B)
                place(student,"År3",A)
                place(student,"År4",B)

            else:
                place(student,"År1",A)

                for sk in [B]+skolor_r:
                    if usage[sk]["År2"] < kap[sk] and usage[sk]["År3"] < kap[sk]:
                        if place(student,"År2",sk):
                            place(student,"År3",sk)
                            break

                used = set()
                for r in rows_r:
                    if student in [r["År1"], r["År2"], r["År3"]]:
                        used.add(r["Skola"])

                placed = False
                for sk in skolor_r:
                    if sk not in used and usage[sk]["År4"] < kap[sk]:
                        if place(student,"År4",sk):
                            placed = True
                            break

                if not placed:
                    for sk in skolor_r:
                        if usage[sk]["År4"] < kap[sk]:
                            place(student,"År4",sk)
                            break

    # =========================
    # EXCEL MED FORMATERING
    # =========================
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

        ws.merge_cells(start_row=row_idx,start_column=1,end_row=row_idx,end_column=5)
        cell = ws.cell(row_idx,1)
        cell.value = region.upper()
        cell.fill = fills[region]
        cell.font = header_font
        cell.alignment = center

        row_idx += 2

        for skola in skolor[skolor["Region"]==region]["Skolenhet"]:

            ws.append([skola])

            for c in range(1,6):
                ws.cell(row_idx,c).fill = fills["Skola"]
                ws.cell(row_idx,c).font = bold

            row_idx += 1

            for r in rows:
                if r["Skola"] == skola:
                    ws.append(["",r["År1"],r["År2"],r["År3"],r["År4"]])

                    for c in range(2,6):
                        ws.cell(row_idx,c).alignment = left

                    row_idx += 1

            ws.append([])
            row_idx += 1

    # =========================
    # BLAD 2 (FIXAD)
    # =========================
    ws2 = wb.create_sheet("Översikt studenter")

    ws2.append(["Student","Bostad","Alt","År1","År2/3","År4"])

    for _, s in students.iterrows():

        namn = s["Namn"]

        p1 = p2 = p3 = ""

        for r in rows:
            if r["År1"] == namn:
                p1 = r["Skola"]
            if r["År2"] == namn:
                p2 = r["Skola"]
            if r["År4"] == namn:
                p3 = r["Skola"]

        ws2.append([
            namn,
            s[bost],
            s[alt_col],
            p1,
            p2,
            p3
        ])

    # =========================
    # KLART
    # =========================
    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f)

else:
    st.info("Ladda upp båda filer")

