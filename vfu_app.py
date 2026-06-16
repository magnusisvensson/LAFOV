import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU-system – Placering")

system_file = st.file_uploader("1. Ladda översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Ladda formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolenheter planerade för kull:", value=26)
program = st.selectbox("Program", ["LAFOV","LAGRV","LGFRI"])

# ===== REGION =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t: return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t: return "Karlskrona"
    return "Kalmar"

def school_region(partner, skola):
    s = str(skola).lower()
    p = str(partner).lower()

    if "ljungnäs" in s or "blomstermåla" in s:
        return "Kalmar"
    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"

# ===== DATA =====
if system_file and form_file:

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
    logg = {}
    placed_students = set()

    # ===== HELP =====
    def has_space(skola,år,n):
        return cap_used.get((skola,år),0)+n <= kap.get(skola,999)

    def use(skola,år,n):
        cap_used[(skola,år)] = cap_used.get((skola,år),0)+n

    def add(skola, student, col):
        skol_data.setdefault(skola,{})
        if student not in skol_data[skola]:
            skol_data[skola][student] = {
                "År1":"","År2":"","År3":"","År4":""
            }

        skol_data[skola][student][col] = student

    # ===== ROTATION =====
    def find_rotation(lista, region, n):

        for i in range(len(lista)-2):
            A = lista[i]
            B = lista[i+1]
            C = lista[i+2]

            if region in ["Oskarshamn","Karlskrona"]:

                if all([
                    has_space(A,1,n),
                    has_space(B,2,n),
                    has_space(A,3,n),
                    has_space(B,4,n)
                ]):
                    return ("ABAB",A,B,C)

            else:

                if all([
                    has_space(A,1,n),
                    has_space(B,2,n),
                    has_space(B,3,n),
                    has_space(C,4,n)
                ]):
                    return ("ABBC",A,B,C)

        return None

    def place(group):

        names = [s["Namn"] for s in group if s["Namn"] not in placed_students]
        if not names:
            return False

        region = group[0]["Region"]
        n = len(names)

        regional = skolor[skolor["Region"]==region]["Skolenhet"].tolist()
        alla = list(skolor["Skolenhet"])

        möjliga = regional + [s for s in alla if s not in regional]

        val = find_rotation(möjliga, region, n)

        if val:

            typ,A,B,C = val

            for s in names:

                if typ == "ABAB":
                    add(A,s,"År1")
                    add(B,s,"År2")
                    add(A,s,"År3")
                    add(B,s,"År4")
                else:
                    add(A,s,"År1")
                    add(B,s,"År2")
                    add(B,s,"År3")
                    add(C,s,"År4")

                placed_students.add(s)
                logg[s] = {"Status":"OK"}

            if typ == "ABAB":
                use(A,1,len(names))
                use(B,2,len(names))
                use(A,3,len(names))
                use(B,4,len(names))
            else:
                use(A,1,len(names))
                use(B,2,len(names))
                use(B,3,len(names))
                use(C,4,len(names))

            return True

        return False

    # ===== GRUPPER =====
    stud_list = students.to_dict("records")

    i = 0
    while i < len(stud_list):

        placed = False

        for size in [3,2,1]:

            group = stud_list[i:i+size]
            if len(group) < size:
                continue

            if place(group):
                i += size
                placed = True
                break

        if not placed:
            namn = stud_list[i]["Namn"]
            if namn not in placed_students:
                logg[namn] = {"Status":"Får ej plats"}
            i += 1

    # ===== STUDENTVY (FIXEN) =====
    student_rows = []

    for student in students["Namn"]:

        row = {
            "Student": student,
            "År1": "",
            "År2": "",
            "År3": "",
            "År4": ""
        }

        for skola in skol_data:
            if student in skol_data[skola]:

                data = skol_data[skola][student]

                if data["År1"]:
                    row["År1"] = skola
                if data["År2"]:
                    row["År2"] = skola
                if data["År3"]:
                    row["År3"] = skola
                if data["År4"]:
                    row["År4"] = skola

        student_rows.append(row)

    # ===== EXCEL =====
    wb = Workbook()

    ws = wb.active
    ws.title = "Studentöversikt"

    ws.append(["Student","År1","År2","År3","År4"])

    for r in student_rows:
        ws.append([r["Student"], r["År1"], r["År2"], r["År3"], r["År4"]])

    # ----- Rapport -----
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in logg:
        ws2.append([s, logg[s]["Status"]])

    file = "kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
