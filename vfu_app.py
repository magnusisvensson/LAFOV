
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


# ===== REGION =====
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

    # skolregion
    partner_col = [c for c in skolor.columns if "partner" in c.lower()]
    if partner_col:
        skolor["Region"] = skolor[partner_col[0]].apply(get_region)
    else:
        skolor["Region"] = "Kalmar"

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    student_names = list(students["Namn"])

    # ===== ÅR1 (FYLL) =====
    year = {1:{},2:{},3:{},4:{}}

    idx = 0
    for skola, max_p in kap.items():
        for _ in range(max_p):
            if idx >= len(student_names):
                break
            year[1][student_names[idx]] = skola
            idx += 1


    # ===== HJÄLP: hitta index i skol_lista =====
    def next_school(current, step=1):
        i = skol_lista.index(current)
        return skol_lista[(i + step) % len(skol_lista)]


    # ===== ÅR2 (B) =====
    for s in student_names:
        A = year[1][s]
        B = next_school(A, 1)
        year[2][s] = B

    # ===== ÅR3 = ÅR2 =====
    year[3] = year[2].copy()

    # ===== ÅR4 (C) =====
    for s in student_names:
        B = year[2][s]
        C = next_school(B, 1)
        year[4][s] = C


    # ===== BYGG SKOLVY =====
    school_data = {}

    for s in student_names:
        for y in ["År1","År2","År3","År4"]:
            yr = int(y[-1])
            sk = year[yr][s]

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
