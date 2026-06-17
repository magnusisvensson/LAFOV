
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


def get_region(text):
    t = str(text).lower()
    if "kalmar" in t:
        return "Kalmar"
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "OKÄND"


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 0

    skol_lista = list(kap.keys())

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    student_names = list(students["Namn"])

    # ===== GRUPPERA 2 OCH 2 =====
    grupper = [
        student_names[i:i+2]
        for i in range(0, len(student_names), 2)
    ]

    # ===== ÅR-DATA =====
    year = {1:{},2:{},3:{},4:{}}

    # ===== STEG 1: ÅR1 =====
    group_index = 0
    for skola, max_p in kap.items():

        platser = max_p // 2  # antal grupper

        for _ in range(platser):
            if group_index >= len(grupper):
                break

            grupp = grupper[group_index]

            for student in grupp:
                year[1][student] = skola

            group_index += 1

    # ===== HJÄLP =====
    def next_school(sk, step=1):
        i = skol_lista.index(sk)
        return skol_lista[(i+step) % len(skol_lista)]

    # ===== ÅR2 =====
    for grupp in grupper:
        A = year[1][grupp[0]]
        B = next_school(A)

        for s in grupp:
            year[2][s] = B

    # ===== ÅR3 =====
    year[3] = year[2].copy()

    # ===== ÅR4 =====
    for grupp in grupper:
        B = year[2][grupp[0]]
        C = next_school(B)

        for s in grupp:
            year[4][s] = C


    # ===== BYGG SKOLVY =====
    school_data = {}

    for s in student_names:
        for y in ["År1","År2","År3","År4"]:
            sk = year[int(y[-1])][s]

            school_data.setdefault(sk,{})
            school_data[sk].setdefault(s,{
                "År1":"","År2":"","År3":"","År4":""
            })

            school_data[sk][s][y] = s

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola, max_p in kap.items():

        ws.append([f"{skola} (max {max_p})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        rows=[{"År1":"","År2":"","År3":"","År4":""}
              for _ in range(int(max_p))]

        i = 0
        if skola in school_data:
            for student,data in school_data[skola].items():
                if i >= max_p:
                    break
                rows[i] = data
                i += 1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])


    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in student_names:

        student_region = students.loc[
            students["Namn"] == s, "Region"
        ].values[0]

        skolor_student = [year[y][s] for y in [1,2,3,4]]

        if student_region == "OKÄND":
            status = "OBS - OKÄND ORT - LÅNG PENDLING"
        else:
            long_commute = any(
                skolor.loc[skolor["Skolenhet"]==sk,"Region"].values[0]
                != student_region
                for sk in skolor_student
            )

            status = "OK - OBS LÅNG PENDLING" if long_commute else "OK"

        ws2.append([s, status])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
