
import streamlit as st
import pandas as pd
from openpyxl import Workbook

st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ]

    # filtrera bort regioner
    skolor = skolor[
        ~skolor["Partnerområde"].str.lower().str.contains("karlskrona|oskarshamn|ronneby", na=False)
    ]

    # bygg platser
    slots = []

    for _, r in skolor.iterrows():
        skola = r["Skolenhet"]
        try:
            k = int(r["Antal platser"])
        except:
            k = 0

        slots += [skola] * k

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])

    n_students = len(student_names)

    # säkerställ att vi inte överskrider
    if n_students > len(slots):
        st.error("För många studenter för givna platser")
        st.stop()

    # ===== ROTERA PLATSER =====
    def rotate(lst, n):
        return lst[n:] + lst[:n]

    slots_year1 = slots.copy()
    slots_year2 = rotate(slots, 1)
    slots_year4 = rotate(slots, 2)

    # ===== TILLDELA =====
    year = {1:{},2:{},3:{},4:{}}

    for i, s in enumerate(student_names):
        year[1][s] = slots_year1[i]
        year[2][s] = slots_year2[i]
        year[3][s] = slots_year2[i]
        year[4][s] = slots_year4[i]

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    skol_lista = list(dict.fromkeys(slots))

    for skola in skol_lista:

        max_p = slots.count(skola)
        ws.append([f"{skola} (max {max_p})"])

        # samla per år
        year_lists = {f"År{i}":[] for i in [1,2,3,4]}

        for s in student_names:
            for y in [1,2,3,4]:
                if year[y][s] == skola:
                    year_lists[f"År{y}"].append(s)

        for i in range(max_p):
            row = []
            for y in ["År1","År2","År3","År4"]:
                row.append(year_lists[y][i] if i < len(year_lists[y]) else "")
            ws.append([""] + row)

        ws.append([])

    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in student_names:
        ws2.append([s,"OK"])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner", f, file_name=file)
