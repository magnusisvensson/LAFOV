
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

    skolor = skolor[
        ~skolor["Partnerområde"].str.lower().str.contains(
            "karlskrona|oskarshamn|ronneby", na=False)
    ]

    kap = {}
    skol_lista = []

    for _, r in skolor.iterrows():
        try:
            k = int(r["Antal platser"])
        except:
            k = 0

        if k > 0:
            kap[r["Skolenhet"]] = k
            skol_lista.append(r["Skolenhet"])


    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])


    # ===== KAPACITETS-KOLL =====
    usage = {y: {s:0 for s in skol_lista} for y in [1,2,3,4]}

    def place(student, year, preferred):

        for offset in range(len(skol_lista)):
            skola = skol_lista[(preferred + offset) % len(skol_lista)]

            if usage[year][skola] < kap[skola]:
                usage[year][skola] += 1
                return skola

        return None


    # ===== PLACERING =====
    year = {1:{},2:{},3:{},4:{}}

    not_placed = []

    for idx, s in enumerate(student_names):

        # År1
        A = place(s, 1, idx % len(skol_lista))
        if not A:
            not_placed.append(s)
            continue

        # År2
        B = place(s, 2, (skol_lista.index(A)+1))
        if not B:
            not_placed.append(s)
            continue

        # År3
        if usage[3][B] < kap[B]:
            C2 = B
            usage[3][B] += 1
        else:
            C2 = place(s, 3, skol_lista.index(B))

        # År4
        C = place(s, 4, (skol_lista.index(B)+1))
        if not C:
            not_placed.append(s)
            continue

        year[1][s] = A
        year[2][s] = B
        year[3][s] = C2
        year[4][s] = C


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in skol_lista:

        ws.append([f"{skola} (max {kap[skola]})"])

        year_lists = {f"År{i}":[] for i in [1,2,3,4]}

        for s in student_names:
            for y in [1,2,3,4]:
                if year[y].get(s) == skola:
                    year_lists[f"År{y}"].append(s)

        for i in range(kap[skola]):
            row = []
            for y in ["År1","År2","År3","År4"]:
                row.append(year_lists[y][i] if i < len(year_lists[y]) else "")
            ws.append([""] + row)

        ws.append([])


    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in student_names:
        if s in not_placed:
            ws2.append([s,"EJ PLACERAD"])
        else:
            ws2.append([s,"OK"])


    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner", f, file_name=file)
