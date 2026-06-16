
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolenheter planerade för kull:", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])

# ===== REGIONLOGIK =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "Kalmar"

def school_region(partner, skola):
    p = str(partner).lower()
    s = str(skola).lower()

    # KORRIGERINGAR
    if "ljungnäs" in s or "blomstermåla" in s:
        return "Kalmar"

    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"

# ===== APP =====
if system_file is not None and form_file is not None:

    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor.apply(
        lambda r: school_region(r["Partnerområde"], r["Skolenhet"]), axis=1
    )

    kap = dict(zip(skolor["Skolenhet"], skolor["Antal platser"]))

    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn = [c for c in students.columns if "förnamn" in c.lower()][0]
    ln = [c for c in students.columns if "efternamn" in c.lower()][0]
    bost = [c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn] + " " + students[ln]
    students["Region"] = students[bost].apply(get_region)

    skol_data = {}
    cap_used = {}
    logg = []

    def add(skola, student, col):
        skol_data.setdefault(skola,{})
        skol_data[skola].setdefault(student,{"År1":"","År2":"","År3":"","År4":""})
        skol_data[skola][student][col] = student

    # ===== HJÄLP: CHECK PER ÅR =====
    def has_space(skola, år, n):
        return cap_used.get((skola,år),0) + n <= kap.get(skola,999)

    def use(skola, år, n):
        cap_used[(skola,år)] = cap_used.get((skola,år),0) + n

    # ===== KÄRNPLACERING =====
    def place(students_list):

        region = students_list[0]["Region"]
        names = [s["Namn"] for s in students_list]
        n = len(names)

        regional = skolor[skolor["Region"] == region]["Skolenhet"].tolist()
        all_schools = list(skolor["Skolenhet"])

        search_lists = [
            regional,
            regional + [s for s in all_schools if s not in regional],
            all_schools
        ]

        for möjliga in search_lists:

            # ✅ 1: FULL ROTATION
            for i in range(len(möjliga)-2):
                A,B,C = möjliga[i], möjliga[i+1], möjliga[i+2]

                if all([
                    has_space(A,1,n),
                    has_space(B,2,n),
                    has_space(B,3,n),
                    has_space(C,4,n)
                ]):
                    for s in names:
                        add(A,s,"År1")
                        add(B,s,"År2")
                        add(B,s,"År3")
                        add(C,s,"År4")

                    use(A,1,n); use(B,2,n); use(B,3,n); use(C,4,n)
                    return True

            # ✅ 2: TVÅ SKOLOR
            for i in range(len(möjliga)-1):
                A,B = möjliga[i], möjliga[i+1]

                if all([
                    has_space(A,1,n),
                    has_space(B,2,n),
                    has_space(A,3,n),
                    has_space(B,4,n)
                ]):
                    for s in names:
                        add(A,s,"År1")
                        add(B,s,"År2")
                        add(A,s,"År3")
                        add(B,s,"År4")

                    use(A,1,n); use(B,2,n); use(A,3,n); use(B,4,n)
                    return True

        # ✅ 3: MAXFYLL PER ÅR
        for skola in all_schools:
            for år,col in [(1,"År1"),(2,"År2"),(3,"År3"),(4,"År4")]:
                if has_space(skola,år,n):
                    for s in names:
                        add(skola,s,col)
                    use(skola,år,n)
                    return True

        return False

    # ===== DYNAMISKA GRUPPER =====
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
