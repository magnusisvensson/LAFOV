
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolenheter planerade för kull:", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])

def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "Kalmar"

if system_file is not None and form_file is not None:

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    kap = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    cap_used = {}
    skol_data = {}
    logg = []

    def add(skola, student, col):
        skol_data.setdefault(skola,{})
        skol_data[skola].setdefault(student,{"År1":"","År2":"","År3":"","År4":""})
        skol_data[skola][student][col]=student

    # ===== FUNKTION SOM FÖRSÖKER PLACERA EN LISTA =====
    def place_list(student_list):

        region = student_list[0]["Region"]
        namn_lista = [s["Namn"] for s in student_list]

        möjliga = skolor[
            skolor["Partnerområde"].str.contains(region, case=False, na=False)
        ]["Skolenhet"].tolist()

        if len(möjliga) < 2:
            möjliga = list(skolor["Skolenhet"])

        # ===== OSKARSHAMN/KARLSKRONA =====
        if region in ["Oskarshamn","Karlskrona"]:

            for i in range(len(möjliga)-1):

                A = möjliga[i]
                B = möjliga[i+1]

                if (
                    cap_used.get((A,1),0)+len(student_list) <= kap.get(A,999) and
                    cap_used.get((B,2),0)+len(student_list) <= kap.get(B,999) and
                    cap_used.get((A,3),0)+len(student_list) <= kap.get(A,999) and
                    cap_used.get((B,4),0)+len(student_list) <= kap.get(B,999)
                ):

                    for namn in namn_lista:
                        add(A,namn,"År1")
                        add(B,namn,"År2")
                        add(A,namn,"År3")
                        add(B,namn,"År4")

                    cap_used[(A,1)] = cap_used.get((A,1),0)+len(student_list)
                    cap_used[(B,2)] = cap_used.get((B,2),0)+len(student_list)
                    cap_used[(A,3)] = cap_used.get((A,3),0)+len(student_list)
                    cap_used[(B,4)] = cap_used.get((B,4),0)+len(student_list)

                    return True

        # ===== KALMAR =====
        else:

            for i in range(len(möjliga)-2):

                A = möjliga[i]
                B = möjliga[i+1]
                C = möjliga[i+2]

                if (
                    cap_used.get((A,1),0)+len(student_list) <= kap.get(A,999) and
                    cap_used.get((B,2),0)+len(student_list) <= kap.get(B,999) and
                    cap_used.get((B,3),0)+len(student_list) <= kap.get(B,999) and
                    cap_used.get((C,4),0)+len(student_list) <= kap.get(C,999)
                ):

                    for namn in namn_lista:
                        add(A,namn,"År1")
                        add(B,namn,"År2")
                        add(B,namn,"År3")
                        add(C,namn,"År4")

                    cap_used[(A,1)] = cap_used.get((A,1),0)+len(student_list)
                    cap_used[(B,2)] = cap_used.get((B,2),0)+len(student_list)
                    cap_used[(B,3)] = cap_used.get((B,3),0)+len(student_list)
                    cap_used[(C,4)] = cap_used.get((C,4),0)+len(student_list)

                    return True

        return False

    # ===== DYNAMISK GRUPPLOGIK =====
    student_list = students.to_dict("records")

    i = 0
    while i < len(student_list):

        # börja med grupp 3
        group = student_list[i:i+3]

        if place_list(group):
            for s in group:
                logg.append({"Student":s["Namn"],"Status":"OK"})
            i += 3
            continue

        # fallback grupp 2
        group = student_list[i:i+2]

        if place_list(group):
            for s in group:
                logg.append({"Student":s["Namn"],"Status":"OK"})
            i += 2
            continue

        # fallback singel
        group = [student_list[i]]

        if place_list(group):
            logg.append({"Student":group[0]["Namn"],"Status":"OK"})
        else:
            logg.append({"Student":group[0]["Namn"],"Status":"Får ej plats"})

        i += 1

    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active

    ws.column_dimensions["A"].width = 40
    for c in ["B","C","D","E"]:
        ws.column_dimensions[c].width = 25

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(skol_data):

        ws.append([f"{skola} (max {int(kap.get(skola,0))})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        for student, years in skol_data[skola].items():
            ws.append(["", years["År1"], years["År2"], years["År3"], years["År4"]])

        ws.append([])

    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for r in logg:
        ws2.append([r["Student"], r["Status"]])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file, "rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)
