
import streamlit as st
import pandas as pd
from openpyxl import Workbook

st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    # ✅ FILTRERA KULL + PROGRAM
    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    # ✅ TA BORT Karlskrona & Oskarshamn
    skolor = skolor[
        ~skolor["Partnerområde"].str.lower().str.contains(
            "karlskrona|ronneby|oskarshamn", na=False
        )
    ]

    skol_lista = list(skolor["Skolenhet"])

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])

    # ===== GRUPPER (2-2) =====
    grupper = []
    i = 0
    while i < len(student_names):
        grupper.append(student_names[i:i+2])
        i += 2

    # ===== ROTATION (A-B-B-C) =====
    year = {1:{},2:{},3:{},4:{}}

    n = len(skol_lista)

    for g_idx, grupp in enumerate(grupper):

        A = skol_lista[g_idx % n]
        B = skol_lista[(g_idx + 1) % n]
        C = skol_lista[(g_idx + 2) % n]

        for s in grupp:
            year[1][s] = A
            year[2][s] = B
            year[3][s] = B
            year[4][s] = C

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in skol_lista:

        ws.append([skola])

        # samla per år
        year_lists = {
            "År1": [],
            "År2": [],
            "År3": [],
            "År4": []
        }

        for s in student_names:
            for y in [1,2,3,4]:
                if year[y][s] == skola:
                    year_lists[f"År{y}"].append(s)

        max_len = max(len(v) for v in year_lists.values())

        for i in range(max_len):
            row = []
            for y in ["År1","År2","År3","År4"]:
                if i < len(year_lists[y]):
                    row.append(year_lists[y][i])
                else:
                    row.append("")
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
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
