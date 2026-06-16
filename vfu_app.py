
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

def school_region(partner, skola):
    p = str(partner).lower()
    s = str(skola).lower()

    # ✅ DIN JUSTERING
    if "ljungnäs" in s or "blomstermåla" in s:
        return "Kalmar"

    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"

if system_file is not None and form_file is not None:

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor.apply(
        lambda r: school_region(r["Partnerområde"], r["Skolenhet"]),
        axis=1
    )

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
        skol_data[skola].setdefault(student,
            {"År1":"","År2":"","År3":"","År4":""})
        skol_data[skola][student][col]=student

    # ===== FLEXIBEL PLACERING =====
    def place(student_list):

        region = student_list[0]["Region"]
        names = [s["Namn"] for s in student_list]

        regional = skolor[skolor["Region"] == region]["Skolenhet"].tolist()
        all_schools = list(skolor["Skolenhet"])

        möjliga = regional + [s for s in all_schools if s not in regional]

        # ===== 1. FÖRSÖK PERFEKT =====
        for i in range(len(möjliga)-2):
            A,B,C = möjliga[i], möjliga[i+1], möjliga[i+2]

            if (
                cap_used.get((A,1),0)+len(names) <= kap.get(A,999) and
                cap_used.get((B,2),0)+len(names) <= kap.get(B,999) and
                cap_used.get((B,3),0)+len(names) <= kap.get(B,999) and
                cap_used.get((C,4),0)+len(names) <= kap.get(C,999)
            ):
                for n in names:
                    add(A,n,"År1")
                    add(B,n,"År2")
                    add(B,n,"År3")
                    add(C,n,"År4")

                cap_used[(A,1)] = cap_used.get((A,1),0)+len(names)
                cap_used[(B,2)] = cap_used.get((B,2),0)+len(names)
                cap_used[(B,3)] = cap_used.get((B,3),0)+len(names)
                cap_used[(C,4)] = cap_used.get((C,4),0)+len(names)
                return True

        # ===== 2. FÖRSÖK 2 SKOLOR =====
        for i in range(len(möjliga)-1):
            A,B = möjliga[i], möjliga[i+1]

            if (
                cap_used.get((A,1),0)+len(names) <= kap.get(A,999) and
                cap_used.get((B,2),0)+len(names) <= kap.get(B,999)
            ):
                for n in names:
                    add(A,n,"År1")
                    add(B,n,"År2")
                    add(A,n,"År3")
                    add(B,n,"År4")

                cap_used[(A,1)] = cap_used.get((A,1),0)+len(names)
                cap_used[(B,2)] = cap_used.get((B,2),0)+len(names)
                return True

        # ===== 3. MAXFYLL =====
        for skola in möjliga:
            if cap_used.get((skola,1),0) < kap.get(skola,999):

                for n in names:
                    add(skola,n,"År1")

                cap_used[(skola,1)] = cap_used.get((skola,1),0)+len(names)
                return True

        return False

    # ===== DYNAMISK GRUPP =====
    stud_list = students.to_dict("records")

    i = 0
    while i < len(stud_list):

        if place(stud_list[i:i+3]):
            logg += [{"Student":s["Namn"],"Status":"OK"} for s in stud_list[i:i+3]]
            i += 3
        elif place(stud_list[i:i+2]):
            logg += [{"Student":s["Namn"],"Status":"OK"} for s in stud_list[i:i+2]]
            i += 2
        elif place([stud_list[i]]):
            logg.append({"Student":stud_list[i]["Namn"],"Status":"OK"})
            i += 1
        else:
            logg.append({"Student":stud_list[i]["Namn"],"Status":"Får ej plats"})
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
