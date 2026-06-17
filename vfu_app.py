
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
    if "kalmar" in t: return "Kalmar"
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "OKÄND"


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    # robust kapacitet
    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 0

    skol_lista = [s for s in kap if kap[s] > 0]

    # skolregion
    if "Partnerområde" in skolor.columns:
        skolor["Region"] = skolor["Partnerområde"].apply(get_region)
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

    # ===== ÅR1 (GARANTERA ALLA) =====
    year = {1:{},2:{},3:{},4:{}}

    slots = []
    for skola, max_p in kap.items():
        slots += [skola] * int(max_p)

    while len(slots) < len(student_names):
        slots += skol_lista

    for i, student in enumerate(student_names):
        year[1][student] = slots[i]

    # ===== GRUPPER =====
    grupper = {}
    for s, sk in year[1].items():
        grupper.setdefault(sk, []).append(s)

    # ===== FIXA ENSAMMA =====
    for skola, grupp in list(grupper.items()):
        if len(grupp) == 1:
            for sk2, grp2 in grupper.items():
                if sk2 != skola and len(grp2) >= 3:
                    grupper[skola].append(grp2.pop())
                    break

    # ✅ SYNKA ÅR1
    for skola, grupp in grupper.items():
        for s in grupp:
            year[1][s] = skola

    # ===== ROTATION =====
    def next_school(skola, step=1):
        i = skol_lista.index(skola)
        return skol_lista[(i + step) % len(skol_lista)]

    for skola, grupp in grupper.items():
        B = next_school(skola, 1)
        C = next_school(skola, 2)

        for s in grupp:
            year[2][s] = B
            year[3][s] = B
            year[4][s] = C

    # ===== BYGG SKOLDATA =====
    school_data = {}

    for s in student_names:
        for y in [1,2,3,4]:
            sk = year[y][s]

            school_data.setdefault(sk, {})
            school_data[sk].setdefault(s, {
                "År1":"","År2":"","År3":"","År4":""
            })
            school_data[sk][s][f"År{y}"] = s

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola, max_p in kap.items():

        ws.append([f"{skola} (max {max_p})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        # ✅ NY LOGIK (KORREKT)
        year_lists = {
            "År1": [],
            "År2": [],
            "År3": [],
            "År4": []
        }

        for student, data in school_data.get(skola, {}).items():
            for y in ["År1","År2","År3","År4"]:
                if data[y] != "":
                    year_lists[y].append(student)

        rows = []
        for i in range(max_p):
            row = {}
            for y in ["År1","År2","År3","År4"]:
                if i < len(year_lists[y]):
                    row[y] = year_lists[y][i]
                else:
                    row[y] = ""
            rows.append(row)

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

else:
    st.info("Ladda upp båda filer")
