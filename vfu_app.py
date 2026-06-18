
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

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data", engine="openpyxl")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    alt_bost_col = [c for c in students.columns if "alternativ" in c.lower()][0]
    pref_col = [c for c in students.columns if "helst utgå" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]

    # välj rätt ort
    def choose_location(row):
        val = str(row[pref_col]).lower()
        if "alternativ" in val and pd.notna(row[alt_bost_col]):
            return row[alt_bost_col]
        return row[bost]

    students["ChosenOrt"] = students.apply(choose_location, axis=1)
    students["Region"] = students["ChosenOrt"].apply(get_region)

    # ===== PLATSER =====
    rows = []
    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
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

        usage = {sk: {"År1":0,"År2":0,"År3":0,"År4":0} for sk in skolor_r}

        def place(student, year, skola):
            for r in rows_r:
                if r["Skola"] == skola and r[year] == "":
                    r[year] = student
                    usage[skola][year] += 1
                    return True
            return False

        for i, student in enumerate(stud_r):

            A = skolor_r[i % len(skolor_r)]
            B = skolor_r[(i+1) % len(skolor_r)] if len(skolor_r) > 1 else A
            C = skolor_r[(i+2) % len(skolor_r)] if len(skolor_r) > 2 else B

            if program == "LGFRI":
                place(student, "År1", A)
                place(student, "År2", B)
                place(student, "År3", A)
                place(student, "År4", B)

            else:
                place(student, "År1", A)

                for sk in [B] + skolor_r:
                    if usage[sk]["År2"] < kap[sk] and usage[sk]["År3"] < kap:
                        if place(student, "År2", sk):
                            place(student, "År3", sk)
                            break

                used = set()
                for r in rows_r:
                    if student in [r["År1"], r["År2"], r["År3"]]:
                        used.add(r["Skola"])

                placed4 = False
                for sk in skolor_r:
                    if sk not in used and usage[sk]["År4"] < kap[sk]:
                        if place(student, "År4", sk):
                            placed4 = True
                            break

                if not placed4:
                    for sk in skolor_r:
                        if usage[sk]["År4"] < kap[sk]:
                            place(student, "År4", sk)
                            break

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])
    for r in rows:
        ws.append([r["Skola"], r["År1"], r["År2"], r["År3"], r["År4"]])

    # ===== BLAD 2 =====
    ws2 = wb.create_sheet("Översikt studenter")
    ws2.append(["Student","Bostad","Alt","År1","År2/3","År4"])

    # ===== BLAD 3 (FÄRGKODAT) =====
    ws3 = wb.create_sheet("Kontroll")

    ws3.append([
        "Student","Region","Ort",
        "År1","År2/3","År4",
        "Antal skolor","Status"
    ])

    fills = {
        "OK": PatternFill("solid", "C6EFCE"),
        "MEDIUM": PatternFill("solid", "FFEB9C"),
        "WARNING": PatternFill("solid", "FFD966"),
        "BAD": PatternFill("solid", "FFC7CE")
    }

    for _, s in students.iterrows():

        namn = s["Namn"]
        region = s["Region"]
        ort = s["ChosenOrt"]

        p1 = p2 = p3 = ""
        skolor_set = set()

        for r in rows:
            if r["År1"] == namn:
                p1 = r["Skola"]
                skolor_set.add(r["Skola"])
            if r["År2"] == namn:
                p2 = r["Skola"]
                skolor_set.add(r["Skola"])
            if r["År4"] == namn:
                p3 = r["Skola"]
                skolor_set.add(r["Skola"])

        antal = len(skolor_set)

        if p1 == "" or p2 == "" or p3 == "":
            status = "SAKNAR PLACERING"
            fill = fills["BAD"]

        elif antal == 1:
            status = "EN SKOLA ALLA ÅR"
            fill = fills["WARNING"]

        elif antal == 2:
            status = "OK (2 skolor)"
            fill = fills["MEDIUM"]

        else:
            status = "BRA ROTATION"
            fill = fills["OK"]

        ws3.append([namn, region, ort, p1, p2, p3, antal, status])

        for c in range(1, 9):
            ws3.cell(ws3.max_row, c).fill = fill

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file, "rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
