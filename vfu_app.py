
import streamlit as st
import pandas as pd
import Font, PatternFill, Alignment

st.title("VFU-placeringssystem")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


# =========================
# REGION
# =========================
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    # =========================
    # DATA
    # =========================
    skolor = pd.read_excel(system_file, engine="openpyxl")
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(float(r["Antal platser"]))
        except:
            kap[r["Skolenhet"]] = 0

    students = pd.read_excel(form_file, engine="openpyxl")
    students.columns = students.columns.str.strip()


    # =========================
    # ✅ ENKEL KOLUMNMATCHNING
    # =========================
    def find_col(label, keyword):

        matches = [c for c in students.columns if keyword in c.lower()]

        if matches:
            return matches[0]

        st.warning(f"Kunde inte hitta '{label}' – välj manuellt")
        return st.selectbox(label, students.columns, key=label)


    fn = find_col("Förnamn", "förnamn")
    ln = find_col("Efternamn", "efternamn")
    bost = find_col("Bostadsort", "bostads")
    alt_col = find_col("Alternativ ort", "alternativ")
    pref_col = find_col("Val av ort", "utgå")


    students["Namn"] = students[fn] + " " + students[ln]


    def choose_loc(row):
        if "alternativ" in str(row[pref_col]).lower() and pd.notna(row[alt_col]):
            return row[alt_col]
        return row[bost]


    students["ChosenOrt"] = students.apply(choose_loc, axis=1)
    students["Region"] = students["ChosenOrt"].apply(get_region)


    # =========================
    # PLATSER
    # =========================
    rows = []

    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1": "", "År2": "", "År3": "", "År4": ""
            })


    # =========================
    # PLACERING
    # =========================
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"] == region]
        stud_r = list(students[students["Region"] == region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        usage = {sk: {"År1":0,"År2":0,"År3":0,"År4":0} for sk in skolor_r}

        def place(student, year, sk):
            for r in rows_r:
                if r["Skola"] == sk and r[year] == "":
                    r[year] = student
                    usage[sk][year] += 1
                    return True
            return False

        for i, student in enumerate(stud_r):

            A = skolor_r[i % len(skolor_r)]
            B = skolor_r[(i+1) % len(skolor_r)]

            # ✅ LGFRI
            if program == "LGFRI":
                place(student,"År1",A)
                place(student,"År2",A)
                place(student,"År3",B)
                continue

            place(student,"År1",A)

            for sk in [B] + skolor_r:
                if usage[sk]["År2"] < kap[sk] and usage[sk]["År3"] < kap[sk]:
                    if place(student,"År2",sk):
                        place(student,"År3",sk)
                        break

            used = {r["Skola"] for r in rows_r if student in [r["År1"],r["År2"],r["År3"]]}

            for sk in skolor_r:
                if sk not in used and usage[sk]["År4"] < kap[sk]:
                    if place(student,"År4",sk):
                        break


    # =========================
    # PENDLINGSKONTROLL
    # =========================
    st.subheader("🚶 Pendlingskontroll")

    student_input = st.text_input("Ange student")

    if student_input:

        match = students[students["Namn"].str.lower() == student_input.lower()]

        if len(match):

            sr = match.iloc[0]
            bostad = sr["ChosenOrt"]
            region = sr["Region"]

            st.write(f"Bostadsort: {bostad}")

            p1=p2=p3=""
            for r in rows:
                if r["År1"] == sr["Namn"]: p1=r["Skola"]
                if r["År2"] == sr["Namn"]: p2=r["Skola"]
                if r["År4"] == sr["Namn"]: p3=r["Skola"]

            for year, skola in [("År1",p1),("År2/3",p2),("År4",p3)]:

                if not skola:
                    continue

                val = st.radio(
                    f"{year}: OK pendling {bostad} → {skola}?",
                    ["Ja","Nej"],
                    key=f"{student_input}_{year}"
                )

                if val == "Nej":

                    alternatives = list(set([
                        r["Skola"] for r in rows
                        if r["Region"]==region and r["Skola"]!=skola
                    ]))

                    if alternatives:
                        st.selectbox("Förslag", alternatives)
                    else:
                        st.error("Ingen plats")


    # =========================
    # EXCEL
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fills = {
        "Kalmar": PatternFill("solid","D9EAF7"),
        "Oskarshamn": PatternFill("solid","DFF5DF"),
        "Karlskrona": PatternFill("solid","FFF4CC"),
        "Skola": PatternFill("solid","E7E7E7"),
        "Header": PatternFill("solid","CCCCCC"),
    }

    # HEADER
    if program=="LGFRI":
        ws.append(["Skola","År1","År2","År3"])
    else:
        ws.append(["Skola","År1","År2","År3","År4"])

    for c in range(1,ws.max_column+1):
        cell = ws.cell(1,c)
        cell.fill = fills["Header"]
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # DATA
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([])
        ws.append([region.upper()])
        row = ws.max_row

        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ws.max_column)

        for c in range(1,ws.max_column+1):
            ws.cell(row,c).fill = fills[region]

        ws.append([])

        for skola in skolor[skolor["Region"]==region]["Skolenhet"]:

            ws.append([skola])
            row = ws.max_row

            for c in range(1,ws.max_column+1):
                ws.cell(row,c).fill = fills["Skola"]
                ws.cell(row,c).font = Font(bold=True)

            school_rows = [r for r in rows if r["Skola"]==skola]

            for r in school_rows:

                if program=="LGFRI":
                    ws.append(["",r["År1"],r["År2"],r["År3"]])
                else:
                    ws.append(["",r["År1"],r["År2"],r["År3"],r["År4"]])

            ws.append([])

    # BLAD 2
    ws2 = wb.create_sheet("Studenter")
    ws2.append(["Student","Bostad","Alt"])

    for _,s in students.iterrows():
        ws2.append([s["Namn"],s[bost],s[alt_col]])

    # BLAD 3
    ws3 = wb.create_sheet("Kontroll")
    ws3.append(["Student","Antal skolor","Status"])

    for _,s in students.iterrows():

        skolset = set([r["Skola"] for r in rows if s["Namn"] in r.values()])
        antal = len(skolset)

        if antal == 0:
            status="SAKNAR"; color="FFC7CE"
        elif antal == 1:
            status="EN"; color="FFD966"
        elif antal == 2:
            status="OK"; color="FFEB9C"
        else:
            status="BRA"; color="C6EFCE"

        ws3.append([s["Namn"],antal,status])

        rN = ws3.max_row
        for c in range(1,4):
            ws3.cell(rN,c).fill = PatternFill("solid",color)

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button(
            "⬇️ Ladda ner Excel",
            f,
            file_name=file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

import pandas as pd

from openpyxl import Workbook
