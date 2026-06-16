import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU – Placering (rotation korrekt)")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


def school_region(partner):
    p = str(partner).lower()
    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(school_region)

    kap = {
        r["Skolenhet"]: int(r["Antal platser"]) if pd.notna(r["Antal platser"]) else 0
        for _, r in skolor.iterrows()
    }

    skol_lista = list(skolor["Skolenhet"])

    def school_sort_key(s):
        region = skolor.loc[skolor["Skolenhet"]==s,"Region"].values[0]
        order = {"Kalmar":0,"Oskarshamn":1,"Karlskrona":2}
        return (order.get(region,3), s)


    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    student_names = list(students["Namn"])

    n = len(student_names)

    # ===== BYGG PLATSER =====
    def build_slots():
        slots = []
        for skola, k in kap.items():
            slots += [skola]*k
        return slots

    slots_y1 = build_slots()[:n]
    slots_y2 = build_slots()[:n]
    slots_y4 = build_slots()[:n]

    # ===== ROTATION =====
    # År2: rotera 1 steg
    slots_y2 = slots_y2[1:] + slots_y2[:1]

    # År4: rotera 2 steg
    slots_y4 = slots_y4[2:] + slots_y4[:2]

    # ===== STUDENTDATA =====
    student_rows = {}

    for i, namn in enumerate(student_names):

        student_rows[namn] = {
            "År1": slots_y1[i],
            "År2": slots_y2[i],
            "År3": slots_y2[i],  # A-B-B-C
            "År4": slots_y4[i]
        }

    # ===== SKOLVY =====
    school_data = {}

    for student, data in student_rows.items():

        for year, skola in data.items():

            school_data.setdefault(skola, {})
            school_data[skola].setdefault(student,{
                "År1":"","År2":"","År3":"","År4":""
            })

            school_data[skola][student][year] = student

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(kap.keys(), key=school_sort_key):

        max_platser = kap.get(skola,0)

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)

        rows = [
            {"År1":"","År2":"","År3":"","År4":""}
            for _ in range(max_platser)
        ]

        i = 0
        if skola in school_data:
            for student,data in school_data[skola].items():
                if i >= max_platser:
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
        ws2.append([s,"OK"])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel",f,file_name=file)

else:
    st.info("Ladda upp båda filer")
