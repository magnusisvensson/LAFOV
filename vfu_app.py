
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

    # ========= DATA =========
    skolor = pd.read_excel(system_file, engine="openpyxl")
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 0

    students = pd.read_excel(form_file, sheet_name="Data", engine="openpyxl")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]
    alt_col = [c for c in students.columns if "alternativ" in c.lower()][0]
    pref_col = [c for c in students.columns if "helst utgå" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]

    def choose_loc(row):
        if "alternativ" in str(row[pref_col]).lower() and pd.notna(row[alt_col]):
            return row[alt_col]
        return row[bost]

    students["ChosenOrt"] = students.apply(choose_loc, axis=1)
    students["Region"] = students["ChosenOrt"].apply(get_region)

    # ========= PLATSER =========
    rows = []
    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "", "År2": "", "År3": "", "År4": ""
            })

    # ========= PLACERING =========
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
            B = skolor_r[(i+1) % len(skolor_r)] if len(skolor_r)>1 else A

            # ✅ LGFRI FIX
            if program == "LGFRI":
                place(student,"År1",A)
                place(student,"År2",A)
                place(student,"År3",B)
                continue

            # ✅ övriga program
            C = skolor_r[(i+2) % len(skolor_r)] if len(skolor_r)>2 else B

            place(student,"År1",A)

            for sk in [B]+skolor_r:
                if usage[sk]["År2"] < kap[sk] and usage[sk]["År3"] < kap[sk]:
                    if place(student,"År2",sk):
                        place(student,"År3",sk)
                        break

            used = {r["Skola"] for r in rows_r if student in [r["År1"],r["År2"],r["År3"]]}

            for sk in skolor_r:
                if sk not in used and usage[sk]["År4"] < kap[sk]:
                    if place(student,"År4",sk):
                        break

    # ========= EXCEL =========
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    left = Alignment(horizontal="left")

    fills = {
        "Kalmar": PatternFill("solid", "D9EAF7"),
        "Oskarshamn": PatternFill("solid", "DFF5DF"),
        "Karlskrona": PatternFill("solid", "FFF4CC"),
        "Skola": PatternFill("solid", "E7E7E7"),
    }

    ws.append(["Skola","År1","År2","År3","År4"])

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])
        current = ws.max_row

        ws.merge_cells(start_row=current,start_column=1,end_row=current,end_column=5)

        for c in range(1,6):
            ws.cell(current,c).fill = fills[region]

        ws.append([])

        for skola in skolor[skolor["Region"]==region]["Skolenhet"]:

            ws.append([skola])
            current = ws.max_row

            for c in range(1,6):
                ws.cell(current,c).fill = fills["Skola"]
                ws.cell(current,c).font = bold

            for r in rows:
                if r["Skola"] == skola:
                    ws.append(["",r["År1"],r["År2"],r["År3"],r["År4"]])
                    row_now = ws.max_row
                    for c in range(2,6):
                        ws.cell(row_now,c).alignment = left

            ws.append([])

    # ========= BLAD 2 =========
    ws2 = wb.create_sheet("Studenter")

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

        ws2.append([namn,s[bost],s[alt_col],p1,p2,p3])

    # ========= BLAD 3 =========
    ws3 = wb.create_sheet("Kontroll")

    ws3.append(["Student","Region","Ort","År1","År2/3","År4","Antal","Status"])

    for _, s in students.iterrows():

        namn = s["Namn"]
        region = s["Region"]
        ort = s["ChosenOrt"]

        skolor_set = set()
        p1 = p2 = p3 = ""

        for r in rows:
            if r["År1"] == namn:
                p1 = r["Skola"]; skolor_set.add(r["Skola"])
            if r["År2"] == namn:
                p2 = r["Skola"]; skolor_set.add(r["Skola"])
            if r["År4"] == namn:
                p3 = r["Skola"]; skolor_set.add(r["Skola"])

        antal = len(skolor_set)

        if p1=="" or p2=="":
            status="SAKNAR"
        elif antal==1:
            status="EN"
        elif antal==2:
            status="OK"
        else:
            status="BRA"

        ws3.append([namn,region,ort,p1,p2,p3,antal,status])


    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button(
            "⬇️ Ladda ner Excel",
            f,
            file_name=file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Ladda upp båda filer")
