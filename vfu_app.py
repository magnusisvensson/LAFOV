
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


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

    kap = {}
    for _, r in skolor.iterrows():
        try:
            k = int(r["Antal platser"])
        except:
            k = 0
        if k > 0:
            kap[r["Skolenhet"]] = k

    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    rows = []
    for _, r in skolor.iterrows():
        for _ in range(int(r["Antal platser"])):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "", "År2": "", "År3": "", "År4": ""
            })

    not_placed = []

    def rotate(lst, n):
        return lst[n:] + lst[:n]

    # ===== KÄRNLOGIK =====
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])

        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        if not rows_r:
            not_placed += stud_r
            continue

        n = len(rows_r)

        students_used = stud_r[:n]
        not_placed += stud_r[n:]

        # ✅ SKOL-SEKVENS
        school_seq = []
        i = 0
        while len(school_seq) < n:
            school_seq.append(skolor_r[i % len(skolor_r)])
            i += 1

        # ✅ ROTATION PER REGION
        if len(skolor_r) <= 2:
            schools_y1 = school_seq
            schools_y2 = rotate(school_seq, 1)
            schools_y3 = schools_y2
            schools_y4 = school_seq
        else:
            schools_y1 = school_seq
            schools_y2 = rotate(school_seq, 1)
            schools_y3 = schools_y2
            schools_y4 = rotate(school_seq, 2)

        names_y1 = students_used
        names_y2 = rotate(students_used, 1)
        names_y3 = names_y2
        names_y4 = rotate(students_used, 2)

        # ✅ SKAPA TOMMAPER SKOLA
        school_rows = {sk: [] for sk in skolor_r}
        for row in rows_r:
            school_rows[row["Skola"]].append(row)

        counters = {sk: 0 for sk in skolor_r}

        # ✅ FYLL RÄTT RADER
        for i in range(n):

            sk1 = schools_y1[i]
            idx = counters[sk1]
            if idx < len(school_rows[sk1]):
                school_rows[sk1][idx]["År1"] = names_y1[i]

            sk2 = schools_y2[i]
            idx = counters[sk2]
            if idx < len(school_rows[sk2]):
                school_rows[sk2][idx]["År2"] = names_y2[i]
                school_rows[sk2][idx]["År3"] = names_y3[i]

            sk4 = schools_y4[i]
            idx = counters[sk4]
            if idx < len(school_rows[sk4]):
                school_rows[sk4][idx]["År4"] = names_y4[i]

            counters[sk1] += 1

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    ws.append(["Skola","År1","År2","År3","År4"])

    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([f"--- {region.upper()} ---"])

        for skola in skolor[skolor["Region"]==region]["Skolenhet"]:
            ws.append([f"{skola} (max {kap[skola]})"])

            for r in rows:
                if r["Skola"] == skola:
                    ws.append(["", r["År1"], r["År2"], r["År3"], r["År4"]])

            ws.append([])

        ws.append([])

    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in students["Namn"]:
        ws2.append([s, "EJ PLACERAD" if s in not_placed else "OK"])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
