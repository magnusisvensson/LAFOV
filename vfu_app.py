import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font

st.title("VFU – Placering (A‑B‑B‑C korrekt)")

system_file = st.file_uploader("1. Översiktsfil", type=["xlsx"])
form_file = st.file_uploader("2. Formulärsvar", type=["xlsx"])

kull = st.number_input("Använd skolor planerade för kull:", value=26)
program = st.selectbox("Inom program:", ["LAFOV","LAGRV","LGFRI"])


# ===== REGION =====
def get_region(text):
    t = str(text).lower()
    if "oskarshamn" in t:
        return "Oskarshamn"
    if "karlskrona" in t or "ronneby" in t:
        return "Karlskrona"
    return "Kalmar"

def school_region(partner):
    p = str(partner).lower()
    if "oskarshamn" in p:
        return "Oskarshamn"
    if "karlskrona" in p or "ronneby" in p:
        return "Karlskrona"
    return "Kalmar"


# ===== MAIN =====
if system_file and form_file:

    # ===== SKOLOR =====
    skolor = pd.read_excel(system_file)
    skolor.columns = skolor.columns.str.strip()

    skolor = skolor[
        (skolor["Kull"] == kull) &
        (skolor["Inriktning"].str.upper() == program)
    ].copy()

    skolor["Region"] = skolor["Partnerområde"].apply(school_region)

    kap = {}
    for _, r in skolor.iterrows():
        try:
            kap[r["Skolenhet"]] = int(r["Antal platser"])
        except:
            kap[r["Skolenhet"]] = 0

    region_schools = {
        "Kalmar": skolor[skolor["Region"]=="Kalmar"]["Skolenhet"].tolist(),
        "Oskarshamn": skolor[skolor["Region"]=="Oskarshamn"]["Skolenhet"].tolist(),
        "Karlskrona": skolor[skolor["Region"]=="Karlskrona"]["Skolenhet"].tolist()
    }

    def school_sort_key(skola):
        region = skolor.loc[skolor["Skolenhet"]==skola,"Region"].values[0]
        order = {"Kalmar":0,"Oskarshamn":1,"Karlskrona":2}
        return (order.get(region,3), skola)


    # ===== STUDENTER =====
    students = pd.read_excel(form_file, sheet_name="Data")
    students.columns = students.columns.str.strip()

    fn=[c for c in students.columns if "förnamn" in c.lower()][0]
    ln=[c for c in students.columns if "efternamn" in c.lower()][0]
    bost=[c for c in students.columns if "bostadsort" in c.lower()][0]

    students["Namn"] = students[fn]+" "+students[ln]
    students["Region"] = students[bost].apply(get_region)

    student_list = students.to_dict("records")

    cap_used = {}
    skol_data = {}
    logg = {}

    # ===== HELP =====
    def has_space(s,y):
        return cap_used.get((s,y),0) < kap.get(s,0)

    def use(s,y):
        cap_used[(s,y)] = cap_used.get((s,y),0)+1

    def set_student_year(student, year, school):
        skol_data.setdefault(student,{
            "År1":"","År2":"","År3":"","År4":""
        })
        skol_data[student][year] = school


    # ===== PLACERING A-B-B-C =====
    for stud in student_list:

        namn = stud["Namn"]
        region = stud["Region"]

        skol_lista = region_schools.get(region, [])

        placed = False

        if len(skol_lista) < 3:
            logg[namn] = "Ej placerad"
            continue

        for i in range(len(skol_lista)):

            A = skol_lista[i]
            B = skol_lista[(i+1) % len(skol_lista)]
            C = skol_lista[(i+2) % len(skol_lista)]

            if all([
                has_space(A,1),
                has_space(B,2),
                has_space(B,3),
                has_space(C,4)
            ]):

                set_student_year(namn,"År1",A)
                set_student_year(namn,"År2",B)
                set_student_year(namn,"År3",B)
                set_student_year(namn,"År4",C)

                use(A,1)
                use(B,2)
                use(B,3)
                use(C,4)

                logg[namn] = "OK"
                placed = True
                break

        if not placed:
            logg[namn] = "Ej placerad"


    # ===== BYGG SKOLVY (KRITISK FIX) =====
    school_view = {}

    for student, data in skol_data.items():

        for year, school in data.items():

            if school == "":
                continue

            school_view.setdefault(school, [])
            school_view[school].append((student, year))


    # ===== EXCEL =====
    wb = Workbook()
    ws = wb.active
    ws.title = "Placeringar"

    fill = PatternFill(start_color="DDDDDD", fill_type="solid")

    ws.append(["Skola","År1","År2","År3","År4"])

    for skola in sorted(kap.keys(), key=school_sort_key):

        max_platser = int(kap.get(skola,0))

        ws.append([f"{skola} (max {max_platser})"])
        r = ws.max_row

        for c in range(1,6):
            ws.cell(r,c).fill = fill
            ws.cell(r,c).font = Font(bold=True)

        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5)

        rows = [{"År1":"","År2":"","År3":"","År4":""}
                for _ in range(max_platser)]

        if skola in school_view:

            students_here = {}

            for student, year in school_view[skola]:
                students_here.setdefault(student,{
                    "År1":"","År2":"","År3":"","År4":""
                })
                students_here[student][year] = student

            i=0
            for student,data in students_here.items():
                if i>=max_platser:
                    break
                rows[i] = data
                i+=1

        for row in rows:
            ws.append(["",row["År1"],row["År2"],row["År3"],row["År4"]])

        ws.append([])

    # ===== RAPPORT =====
    ws2 = wb.create_sheet("Rapport")
    ws2.append(["Student","Status"])

    for s in students["Namn"]:
        ws2.append([s, logg.get(s,"Ej placerad")])

    file="kull_resultat.xlsx"
    wb.save(file)

    with open(file,"rb") as f:
        st.download_button("⬇️ Ladda ner Excel", f, file_name=file)

else:
    st.info("Ladda upp båda filer")
``
