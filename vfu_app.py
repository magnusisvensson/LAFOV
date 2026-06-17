
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

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    # ✅ ta bort Karlskrona / Oskarshamn
    skolor = skolor[
        ~skolor["Partnerområde"].str.lower().str.contains(
            "karlskrona|ronneby|oskarshamn", na=False
        )
    ]

    # kapacitet
    kap = {
        r["Skolenhet"]: int(r["Antal platser"])
        for _, r in skolor.iterrows()
        if pd.notna(r["Antal platser"])
    }

    skol_lista = list(kap.keys())


    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])

    # ===== GRUPPER =====
    grupper = []
    i = 0
    while i < len(student_names):
        grupper.append(student_names[i:i+2])
        i += 2


    # ===== HJÄLP =====
    def can_place(group, school, usage):
        return usage[school] + len(group) <= kap[school]


    # ===== FÖRDELA ÅR1 MED KAPACITET =====
    year = {1:{},2:{},3:{},4:{}}
    usage_y1 = {s:0 for s in skol_lista}

    assignments = []  # (grupp, skola)

    school_idx = 0

    for grupp in grupper:

        placed = False

        for _ in range(len(skol_lista)):
            skola = skol_lista[school_idx % len(skol_lista)]

            if can_place(grupp, skola, usage_y1):
                assignments.append((grupp, skola))
                usage_y1[skola] += len(grupp)
                school_idx += 1
                placed = True
                break

            school_idx += 1

        # fallback (måste alltid placeras)
        if not placed:
            skola = skol_lista[school_idx % len(skol_lista)]
            assignments.append((grupp, skola))
            usage_y1[skola] += len(grupp)


    # ===== ROTATION MED KAPACITET =====
    usage = {
        1: {s:0 for s in skol_lista},
        2: {s:0 for s in skol_lista},
        3: {s:0 for s in skol_lista},
        4: {s:0 for s in skol_lista},
    }

    n = len(skol_lista)

    for idx, (grupp, A) in enumerate(assignments):

        # hitta index för A
        base_idx = skol_lista.index(A)

        B = skol_lista[(base_idx + 1) % n]
        C = skol_lista[(base_idx + 2) % n]

        # placera A
        for s in grupp:
            year[1][s] = A
        usage[1][A] += len(grupp)

        # placera B
        for s in grupp:
            year[2][s] = B
            year[3][s] = B
        usage[2][B] += len(grupp)
        usage[3][B] += len(grupp)

        # placera C
        for s in grupp:
            year[4][s] = C
        usage[4][C] += len(grupp)


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in skol_lista:

        max_p = kap.get(skola, 0)
        ws.append([f"{skola} (max {max_p})"])

        year_lists = {f"År{i}":[] for i in [1,2,3,4]}

        for s in student_names:
            for y in [1,2,3,4]:
                if year[y].get(s) == skola:
                    year_lists[f"År{y}"].append(s)

        max_len = max(len(v) for v in year_lists.values())

        for i in range(max_len):
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
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
