
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

    # filtrera bort Karlskrona/Oskarshamn
    skolor = skolor[
        ~skolor["Partnerområde"].str.lower().str.contains(
            "karlskrona|oskarshamn|ronneby", na=False
        )
    ]

    # bygg platser
    slots = []
    kap = {}

    for _, r in skolor.iterrows():
        skola = r["Skolenhet"]

        try:
            k = int(r["Antal platser"])
        except:
            k = 0

        kap[skola] = k
        slots += [skola] * k

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])

    # ===== ROTATION =====
    def rotate(lst, n):
        return lst[n:] + lst[:n]

    slots_y1 = slots.copy()
    slots_y2 = rotate(slots, 1)
    slots_y4 = rotate(slots, 2)

    year = {1:{},2:{},3:{},4:{}}

    placed_students = set()

    # ✅ TILLDELA SÅ LÅNGT DET GÅR
    for i, s in enumerate(student_names):

        if i < len(slots):
            year[1][s] = slots_y1[i]
            year[2][s] = slots_y2[i]
            year[3][s] = slots_y2[i]
            year[4][s] = slots_y4[i]
            placed_students.add(s)
        else:
            # ❗ ingen plats
            year[1][s] = None
            year[2][s] = None
            year[3][s] = None
            year[4][s] = None


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    skol_lista = list(kap.keys())

    for skola in skol_lista:

        max_p = kap.get(skola, 0)
        ws.append([f"{skola} (max {max_p})"])

        year_lists = {f"År{i}":[] for i in [1,2,3,4]}

        for s in student_names:
            for y in [1,2,3,4]:
                if year[y].get(s) == skola:
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

        if s in placed_students:
            status = "OK"
        else:
            status = "EJ PLACERAD"

        ws2.append([s, status])


    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner", f, file_name=file)
