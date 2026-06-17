
import streamlit as st
import pandas as pd
from openpyxl import Workbook

st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# ===== REGION =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
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

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    # kapacitet
    kap = {
        r["Skolenhet"]: int(r["Antal platser"])
        for _, r in skolor.iterrows()
        if pd.notna(r["Antal platser"]) and int(r["Antal platser"]) > 0
    }

    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    student_names = list(students["Namn"])


    # ===== FÖRDELA PER REGION =====
    year = {1:{},2:{},3:{},4:{}}
    not_placed = []

    regions = ["Kalmar","Karlskrona","Oskarshamn"]

    for region in regions:

        skolor_r = skolor[skolor["Region"] == region]
        studenter_r = students[students["Region"] == region]

        skole_list = list(skolor_r["Skolenhet"])
        kap_r = {
            s: kap[s] for s in skole_list if s in kap
        }

        # bygg slots
        slots = []
        for sk in skole_list:
            slots += [sk] * kap_r.get(sk, 0)

        def rotate(lst, n):
            return lst[n:] + lst[:n]

        slots_y1 = slots.copy()
        slots_y2 = rotate(slots, 1)
        slots_y4 = rotate(slots, 2)

        for i, (_, row) in enumerate(studenter_r.iterrows()):

            s = row["Namn"]

            if i < len(slots):
                year[1][s] = slots_y1[i]
                year[2][s] = slots_y2[i]
                year[3][s] = slots_y2[i]
                year[4][s] = slots_y4[i]
            else:
                year[1][s] = None
                year[2][s] = None
                year[3][s] = None
                year[4][s] = None
                not_placed.append(s)


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    skole_all = list(kap.keys())

    for skola in skole_all:

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
        if s in not_placed:
            ws2.append([s,"EJ PLACERAD"])
        else:
            ws2.append([s,"OK"])


    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
