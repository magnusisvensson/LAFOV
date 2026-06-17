
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

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

    kap = {
        r["Skolenhet"]: int(r["Antal platser"])
        for _, r in skolor.iterrows()
        if pd.notna(r["Antal platser"]) and int(r["Antal platser"]) > 0
    }

    skol_lista = list(kap.keys())


    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)


    # ===== BYGG PLATS-RADER =====
    rows = []

    for _, r in skolor.iterrows():
        skola = r["Skolenhet"]
        k = int(r["Antal platser"])

        for _ in range(k):
            rows.append({
                "Skola": skola,
                "Region": r["Region"],
                "År1": None,
                "År2": None,
                "År3": None,
                "År4": None
            })


    not_placed = []

    # ===== FÖRDELA PER REGION =====
    for region in ["Kalmar", "Oskarshamn", "Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])

        if not rows_r:
            not_placed += stud_r
            continue

        def rotate(lst, n):
            return lst[n:] + lst[:n]

        names1 = stud_r[:len(rows_r)]
        names2 = rotate(names1, 1)
        names4 = rotate(names1, 2)

        for i, row in enumerate(rows_r):

            if i < len(names1):
                row["År1"] = names1[i]
                row["År2"] = names2[i]
                row["År3"] = names2[i]
                row["År4"] = names4[i]
            else:
                row["År1"] = ""
                row["År2"] = ""
                row["År3"] = ""
                row["År4"] = ""

        if len(stud_r) > len(rows_r):
            not_placed += stud_r[len(rows_r):]


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    header_font = Font(bold=True)

    ws.append(["Skola","År1","År2","År3","År4"])

    # ===== REGION BLOK =====
    for region in ["Kalmar", "Oskarshamn", "Karlskrona"]:

        ws.append([f"--- {region.upper()} ---"])
        ws.cell(ws.max_row, 1).font = header_font

        skolor_region = skolor[skolor["Region"] == region]["Skolenhet"]

        for skola in skolor_region:

            max_p = kap.get(skola, 0)
            ws.append([f"{skola} (max {max_p})"])

            rows_s = [r for r in rows if r["Skola"] == skola]

            for r in rows_s:
                ws.append(["", r["År1"], r["År2"], r["År3"], r["År4"]])

            ws.append([])

        ws.append([])


    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in students["Namn"]:
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
