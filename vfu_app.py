
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


st.title("VFU-placeringssystem")

system_file = st.file_uploader("Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("Formulärsvar", type=["xlsx"])

kull = st.number_input("Kull", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])


def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t or "rödeby" in t:
        return "Karlskrona"
    return "Kalmar"


if system_file and form_file:

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(get_region)

    kap = {
        r["Skolenhet"]: int(float(r["Antal platser"]))
        for _, r in skolor.iterrows()
    }

    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    # enkel matchning (funkar)
    def find_col(k):
        return [c for c in students.columns if k in c.lower()][0]

    fn = find_col("förnamn")
    ln = find_col("efternamn")
    bost = find_col("bostads")
    alt = find_col("alternativ")
    pref = find_col("utgå")

    students["Namn"] = students[fn] + " " + students[ln]

    def choose_loc(row):
        if "alternativ" in str(row[pref]).lower():
            return row[alt]
        return row[bost]

    students["ChosenOrt"] = students.apply(choose_loc, axis=1)

    # ✅ region kontroll (ny)
    regions = []

    for i, row in students.iterrows():
        region = get_region(row["ChosenOrt"])

        if region == "Kalmar" and "karlskrona" in row["ChosenOrt"].lower():
            region = st.selectbox(
                f"Välj region för {row['Namn']}",
                ["Kalmar","Karlskrona","Oskarshamn"],
                key=row["Namn"]
            )

        regions.append(region)

    students["Region"] = regions

    # PLATSER
    rows = []
    for _, r in skolor.iterrows():
        for _ in range(kap[r["Skolenhet"]]):
            rows.append({
                "Skola": r["Skolenhet"],
                "Region": r["Region"],
                "År1":"","År2":"","År3":"","År4":""
            })

    # PLACERING
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        rows_r = [r for r in rows if r["Region"]==region]
        stud_r = list(students[students["Region"]==region]["Namn"])
        skolor_r = list(dict.fromkeys([r["Skola"] for r in rows_r]))

        i = 0

        for student in stud_r:

            A = skolor_r[i % len(skolor_r)]
            B = skolor_r[(i+1) % len(skolor_r)]

            for r in rows_r:
                if r["Skola"] == A and r["År1"]=="":
                    r["År1"]=student
                    break

            if program=="LGFRI":
                for r in rows_r:
                    if r["Skola"] == A and r["År2"]=="":
                        r["År2"]=student
                        break
                for r in rows_r:
                    if r["Skola"] == B and r["År3"]=="":
                        r["År3"]=student
                        break

            else:
                for r in rows_r:
                    if r["Skola"] == B and r["År2"]=="" and r["År3"]=="":
                        r["År2"]=student
                        r["År3"]=student
                        break

            i += 1

    # ================= Excel =================
    wb = Workbook()
    ws = wb.active

    header = ["Skola","År1","År2","År3"]
    if program != "LGFRI":
        header.append("År4")

    ws.append(header)

    # format
    for col in range(1,len(header)+1):
        c=ws.cell(1,col)
        c.font=Font(bold=True)
        c.fill=PatternFill("solid","CCCCCC")

    # data
    for region in ["Kalmar","Oskarshamn","Karlskrona"]:

        ws.append([region.upper()])
        ws.merge_cells(start_row=ws.max_row,start_column=1,
                       end_row=ws.max_row,end_column=len(header))

        for skola in skolor[skolor["Region"]==region]["Skolenhet"]:

            ws.append([f"{skola} (max {kap[skola]})"])

            for r in rows:
                if r["Skola"]==skola:
                    row=[ "",r["År1"],r["År2"],r["År3"] ]
                    if program!="LGFRI":
                        row.append(r["År4"])
                    ws.append(row)

            ws.append([])

    # AUTO WIDTH
    for col in ws.columns:
        max_len=max(len(str(cell.value)) for cell in col)
        ws.column_dimensions[col[0].column_letter].width=max_len+2

    # ===== BLAD 2 =====
    ws2=wb.create_sheet("Studenter")
    ws2.append(["Student","Ort","År1","År2/3","År4"])

    for _, s in students.iterrows():

        p1=p2=p3=""

        for r in rows:
            if r["År1"]==s["Namn"]: p1=r["Skola"]
            if r["År2"]==s["Namn"]: p2=r["Skola"]
            if r["År4"]==s["Namn"]: p3=r["Skola"]

        ws2.append([s["Namn"],s["ChosenOrt"],p1,p2,p3])

    # ===== BLAD 3 =====
    ws3=wb.create_sheet("Kontroll")
    ws3.append(["Student","Region","Ort","År1","År2","År4","Antal","Status"])

    for _, s in students.iterrows():

        skol=set()
        p1=p2=p3=""

        for r in rows:
            if r["År1"]==s["Namn"]: p1=r["Skola"]; skol.add(r["Skola"])
            if r["År2"]==s["Namn"]: p2=r["Skola"]; skol.add(r["Skola"])
            if r["År4"]==s["Namn"]: p3=r["Skola"]; skol.add(r["Skola"])

        antal=len(skol)

        if antal==0:
            status="SAKNAR"; color="FFCCCC"
        elif antal==1:
            status="EN"; color="FFD9B3"
        elif antal==2:
            status="OK"; color="FFF2CC"
        else:
            status="BRA"; color="CCFFCC"

        ws3.append([s["Namn"],s["Region"],s["ChosenOrt"],p1,p2,p3,antal,status])

        for c in range(1,9):
            ws3.cell(ws3.max_row,c).fill=PatternFill("solid",color)

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("Ladda ner",f,file_name=file)
